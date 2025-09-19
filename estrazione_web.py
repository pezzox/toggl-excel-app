import streamlit as st
from io import BytesIO
import pandas as pd
import pdfplumber, re, tempfile
from typing import Optional

# ---------- configurazione pagina ----------
st.set_page_config(
    page_title="Toggl → Excel",
    page_icon="📊",
    layout="centered"
)
# ATTENZIONE: il limite upload è in .streamlit/config.toml (maxUploadSize), NON qui.

# ---------- parsing ----------
DUR_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
PCT_RE = re.compile(r"^\d{1,3}(?:\.\d+)?%$")

def extract_words_page(page):
    """Ritaglia l’area utile e restituisce DataFrame parole con coordinate."""
    w, h = page.width, page.height
    bbox = (w * 0.04, h * 0.06, w * 0.96, h * 0.95)
    body = page.crop(bbox)
    words = body.extract_words(
        x_tolerance=2.0, y_tolerance=2.5,
        keep_blank_chars=False, use_text_flow=True
    )
    return pd.DataFrame(words) if words else pd.DataFrame()

def left_text_near(words_df, top, x_limit=200.0, y_tol=2.0) -> str:
    """Testo colonna sinistra (Project/Member) vicino alla riga (top)."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) &
        (words_df["x0"] < x_limit)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def find_client_x_min(words_df) -> Optional[float]:
    """Prova a trovare la x minima della colonna CLIENT dall'header."""
    if words_df.empty: 
        return None
    mask = words_df["text"].astype(str).str.strip().str.lower().eq("client")
    if mask.any():
        return float(words_df.loc[mask, "x0"].min()) - 4.0  # piccolo margine
    # fallback: prendi il 80° percentile delle x (parte destra della pagina)
    return float(words_df["x0"].quantile(0.80))

def right_text_near(words_df, top, x_min: float, y_tol=2.0) -> str:
    """Testo colonna destra (Client) vicino alla riga (top), a destra di x_min."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) &
        (words_df["x0"] >= x_min)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def classify_left(left: str):
    """Determina se la riga a sinistra è Project o Member."""
    if not left:
        return None, None
    lo = left.lower()
    if lo.startswith("total"):
        return "TOTAL", None
    if lo.startswith("without"):
        return "Without project", None
    m = re.search(r"\(\d+\)\s*$", left)
    if m:
        return re.sub(r"\s*\(\d+\)\s*$", "", left).strip(), None
    return None, left  # membro

def parse_page(page):
    """Ritorna righe con: PROJECT, MEMBER, DURATION, DURATION_%, CLIENT."""
    words = extract_words_page(page)
    if words.empty:
        return []

    # individua righe usando la durata come "ancora"
    dur = words[words["text"].astype(str).str.match(DUR_RE)].sort_values("top")
    pct = words[words["text"].astype(str).str.match(PCT_RE)].sort_values("top")

    # trova bordo sinistro/destro delle colonne semantiche
    client_x_min = find_client_x_min(words)

    rows = []
    # Accoppiamo duration e percentuale per ordine verticale (con stessa cardinalità tipica)
    for (_, d), (_, p) in zip(dur.iterrows(), pct.iterrows()):
        top = float(d["top"])
        left = left_text_near(words, top)
        proj, mem = classify_left(left)
        if proj is None and mem is None:
            continue

        client = right_text_near(words, top, x_min=client_x_min) if client_x_min is not None else ""
        rows.append({
            "PROJECT": proj,
            "MEMBER": mem,
            "DURATION": d["text"],
            "DURATION_%": p["text"],
            "CLIENT": client
        })
    return rows

@st.cache_data(show_spinner=False)
def process_pdf(file_bytes: bytes) -> pd.DataFrame:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
        tmp.write(file_bytes)
        tmp.flush()
        all_rows = []
        with pdfplumber.open(tmp.name) as pdf:
            for i in range(1, len(pdf.pages)):  # salta sempre pagina 1
                all_rows.extend(parse_page(pdf.pages[i]))
    df = pd.DataFrame(all_rows)
    if not df.empty:
        # tieni solo PROGETTI e TOTAL (niente righe "membro")
        df = df[df["MEMBER"].isna() | (df["MEMBER"].astype(str).str.strip() == "")]
        df = df.drop(columns=["MEMBER"])  # non più utile a questo punto
        # normalizza percentuale con punto
        df["DURATION_%"] = df["DURATION_%"].astype(str).str.replace(",", ".", regex=False)
        # ordina colonne come nello screenshot con CLIENT a destra
        df = df[["PROJECT", "DURATION", "DURATION_%", "CLIENT"]]
    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="breakdown", index=False)
    return output.getvalue()

# ---------- UI ----------
st.markdown("## 📊 Estrattore Toggl → Excel")
st.caption(
    "Carica il report PDF *Project & member breakdown*. "
    "La prima pagina viene ignorata automaticamente. "
    "Verranno estratti **solo i progetti e i totali**, con le colonne **Project, Duration, Duration % e Client**."
)

uploaded = st.file_uploader("Carica il tuo PDF", type=["pdf"], label_visibility="collapsed")

if uploaded:
    with st.spinner("⏳ Elaborazione del PDF in corso..."):
        df = process_pdf(uploaded.read())

    if df.empty:
        st.error("⚠️ Nessuna riga trovata. Verifica che il PDF sia il report giusto.")
    else:
        st.success(f"✅ Estratti {len(df)} progetti.")
        st.dataframe(df, use_container_width=True, height=360)

        xlsx_bytes = df_to_excel_bytes(df)
        st.download_button(
            "📥 Scarica Excel",
            data=xlsx_bytes,
            file_name="breakdown.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.info("⬆️ Carica un file PDF per iniziare.")
