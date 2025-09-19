import streamlit as st
from io import BytesIO
from typing import Optional
import pandas as pd
import pdfplumber
import re
import tempfile

# --------------------- Config pagina (tema in .streamlit/config.toml) ---------------------
st.set_page_config(
    page_title="Toggl → Excel",
    page_icon="📊",
    layout="centered",
)

# --------------------- Regex utili ---------------------
DUR_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
PCT_RE = re.compile(r"^\d{1,3}(?:\.\d+)?%$")

# --------------------- Helpers di parsing ---------------------
def extract_words_page(page) -> pd.DataFrame:
    """Ritaglia l’area utile e restituisce un DataFrame con parole + coordinate."""
    w, h = page.width, page.height
    bbox = (w * 0.04, h * 0.06, w * 0.96, h * 0.95)  # margini per togliere titolo/piedi
    body = page.crop(bbox)
    words = body.extract_words(
        x_tolerance=2.0,
        y_tolerance=2.5,
        keep_blank_chars=False,
        use_text_flow=True,
    )
    return pd.DataFrame(words) if words else pd.DataFrame()

def is_breakdown_page(words_df: pd.DataFrame) -> bool:
    """Riconosce la pagina della tabella 'Project and member breakdown'."""
    if words_df.empty:
        return False
    text = " ".join(words_df["text"].astype(str).str.lower().tolist())
    if "project and member breakdown" in text:
        return True
    # fallback: presenza delle intestazioni principali
    return all(k in text for k in ["project", "member", "duration"])

def left_text_near(words_df: pd.DataFrame, top: float, x_limit: float = 200.0, y_tol: float = 2.0) -> str:
    """Testo della colonna sinistra (Project/Member) vicino alla coordinata 'top'."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) & (words_df["x0"] < x_limit)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def find_client_x_min(words_df: pd.DataFrame) -> Optional[float]:
    """Trova la x minima della colonna CLIENT usando l'header; fallback: lato destro."""
    if words_df.empty:
        return None
    mask = words_df["text"].astype(str).str.strip().str.lower().eq("client")
    if mask.any():
        return float(words_df.loc[mask, "x0"].min()) - 4.0
    return float(words_df["x0"].quantile(0.80))  # fallback ragionevole sul lato destro

def right_text_near(words_df: pd.DataFrame, top: float, x_min: float, y_tol: float = 2.0) -> str:
    """Testo della colonna destra (Client) vicino alla coordinata 'top'."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) & (words_df["x0"] >= x_min)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def classify_left(left: str):
    """Determina se il testo a sinistra è un PROGETTO o un MEMBRO."""
    if not left:
        return None, None
    lo = left.lower()
    if lo.startswith("total"):
        return "TOTAL", None
    if lo.startswith("without"):
        return "Without project", None
    if re.search(r"\(\d+\)\s*$", left):
        return re.sub(r"\s*\(\d+\)\s*$", "", left).strip(), None
    return None, left  # membro

def parse_page(page):
    """Estrae righe: PROJECT, MEMBER, DURATION, DURATION_%, CLIENT dalla pagina."""
    words = extract_words_page(page)
    if words.empty:
        return []

    dur = words[words["text"].astype(str).str.match(DUR_RE)].sort_values("top")
    pct = words[words["text"].astype(str).str.match(PCT_RE)].sort_values("top")
    if dur.empty or pct.empty:
        return []

    client_x_min = find_client_x_min(words)

    rows = []
    # accoppia duration e percentuale per ordine verticale (stesso numero di righe)
    for (_, d), (_, p) in zip(dur.iterrows(), pct.iterrows()):
        top = float(d["top"])
        left = left_text_near(words, top)
        proj, mem = classify_left(left)
        if proj is None and mem is None:
            continue
        client = right_text_near(words, top, x_min=client_x_min) if client_x_min is not None else ""
        rows.append(
            {
                "PROJECT": proj,
                "MEMBER": mem,
                "DURATION": d["text"],
                "DURATION_%": p["text"],
                "CLIENT": client,
            }
        )
    return rows

# --------------------- Pipeline principale ---------------------
@st.cache_data(show_spinner=False)
def process_pdf(file_bytes: bytes) -> pd.DataFrame:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
        tmp.write(file_bytes)
        tmp.flush()

        def collect_rows(pages_range) -> list[dict]:
            rows = []
            with pdfplumber.open(tmp.name) as pdf:
                for i in pages_range:
                    page = pdf.pages[i]
                    words = extract_words_page(page)
                    if not is_breakdown_page(words):
                        continue
                    rows.extend(parse_page(page))
            return rows

        all_rows = []
        with pdfplumber.open(tmp.name) as pdf:
            n = len(pdf.pages)
        # 1) tipico: la tabella è in pagina 2 → parti da lì
        if n > 1:
            all_rows = collect_rows(range(1, n))
        # 2) fallback: se non trovato, prova tutte le pagine
        if not all_rows:
            all_rows = collect_rows(range(0, n))

    df = pd.DataFrame(all_rows)
    if not df.empty:
        # tieni solo PROGETTI e TOTAL (no membri)
        df = df[df["MEMBER"].isna() | (df["MEMBER"].astype(str).str.strip() == "")]
        df = df.drop(columns=["MEMBER"], errors="ignore")
        # normalizza percentuali con il punto
        df["DURATION_%"] = df["DURATION_%"].astype(str).str.replace(",", ".", regex=False)
        # ordina colonne finali (senza AMOUNT)
        df = df[["PROJECT", "DURATION", "DURATION_%", "CLIENT"]]
    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="breakdown", index=False)
    return buffer.getvalue()

# --------------------- UI ---------------------
st.markdown("## 📊 Estrattore Toggl → Excel")
st.caption(
    "Carica il report PDF *Project & member breakdown*. "
    "La prima pagina viene ignorata se non contiene la tabella; la ricerca parte da pagina 2. "
    "Vengono estratti **solo i progetti e i totali** con le colonne **Project, Duration, Duration % e Client**."
)

uploaded = st.file_uploader("Carica PDF", type=["pdf"], label_visibility="collapsed")

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
