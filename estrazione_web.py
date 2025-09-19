import streamlit as st
from io import BytesIO
from typing import Optional, List, Dict
import pandas as pd
import pdfplumber
import re
import tempfile

# --------------------- Config pagina ---------------------
st.set_page_config(page_title="Toggl → Excel", page_icon="📊", layout="centered")

# --------------------- Regex utili ---------------------
DUR_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
PCT_RE = re.compile(r"^\d{1,3}(?:[.,]\d+)?%$")  # accetta 24,13% o 24.13%
# amount: trattino, oppure valuta/testo numerico (simboli €, $, etc. + cifre, separatori)
AMT_RE = re.compile(r"^(-|[€$£]?\s?\d[\d\.\s,]*\d|\d[\d\.\s,]*[€$£])$")

# --------------------- Helpers di parsing ---------------------
def extract_words_page(page) -> pd.DataFrame:
    """Ritaglia l’area utile e restituisce un DataFrame parole+coordinate."""
    w, h = page.width, page.height
    bbox = (w * 0.04, h * 0.06, w * 0.96, h * 0.95)
    body = page.crop(bbox)
    words = body.extract_words(
        x_tolerance=2.0, y_tolerance=2.5, keep_blank_chars=False, use_text_flow=True
    )
    return pd.DataFrame(words) if words else pd.DataFrame(columns=["text","x0","x1","top","bottom"])

def is_breakdown_page(words_df: pd.DataFrame) -> bool:
    """Riconosce la pagina della tabella 'Project and member breakdown'."""
    if words_df.empty:
        return False
    text = " ".join(words_df["text"].astype(str).str.lower().tolist())
    if "project and member breakdown" in text:
        return True
    return all(k in text for k in ["project", "member", "duration"])

def left_text_near(words_df: pd.DataFrame, top: float, x_limit: float, y_tol: float = 2.0) -> str:
    """Testo della colonna sinistra (Project/Member) vicino alla coordinata 'top'."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) & (words_df["x0"] < x_limit)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def find_client_x_min(words_df: pd.DataFrame) -> Optional[float]:
    """Trova la x minima della colonna CLIENT usando l'header; fallback: lato destro."""
    mask = words_df["text"].astype(str).str.strip().str.lower().eq("client")
    if mask.any():
        return float(words_df.loc[mask, "x0"].min()) - 4.0
    return float(words_df["x0"].quantile(0.80)) if not words_df.empty else None

def text_in_band(words_df: pd.DataFrame, top: float, xmin: float, xmax: float, y_tol: float = 2.0) -> str:
    """Concatena il testo compreso tra xmin e xmax alla stessa altezza."""
    local = words_df[
        (words_df["top"].between(top - y_tol, top + y_tol)) &
        (words_df["x0"] >= xmin) & (words_df["x0"] <= xmax)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def classify_left(left: str):
    """Determina se il testo a sinistra è un PROGETTO o un MEMBRO."""
    if not left:
        return None, None
    lo = left.lower().strip()
    if lo.startswith("total"):
        return "TOTAL", None
    if lo.startswith("without"):
        return "Without project", None
    if re.search(r"\(\d+\)\s*$", left):
        return re.sub(r"\s*\(\d+\)\s*$", "", left).strip(), None
    return None, left  # membro

def parse_page(page) -> List[Dict]:
    """
    Estrae righe con tutte le colonne:
    PROJECT, MEMBER, DURATION, DURATION_%, AMOUNT, CLIENT
    """
    words = extract_words_page(page)
    if words.empty:
        return []

    # colonne ancorate sulle durate/percentuali
    dur = words[words["text"].astype(str).str.match(DUR_RE)].sort_values("top")
    pct = words[words["text"].astype(str).str.match(PCT_RE)].sort_values("top")

    if dur.empty or pct.empty:
        return []

    # coordinate colonne
    duration_x_min = float(dur["x0"].min())
    pct_x_min = float(pct["x0"].min())
    client_x_min = find_client_x_min(words) or float(words["x0"].quantile(0.80))

    # margini
    LEFT_MAX = min(duration_x_min, pct_x_min) - 10.0
    AMT_MIN = pct_x_min + 10.0
    AMT_MAX = client_x_min - 10.0

    rows = []
    # Accoppia per ordine verticale (funziona su PDF Toggl)
    for (_, d), (_, p) in zip(dur.iterrows(), pct.iterrows()):
        top = float(d["top"])
        # project / member
        left = left_text_near(words, top, x_limit=LEFT_MAX, y_tol=2.5)
        proj, mem = classify_left(left)
        if proj is None and mem is None:
            continue

        # amount (testo nella banda tra percentuale e client)
        amount_raw = text_in_band(words, top, AMT_MIN, AMT_MAX, y_tol=2.5)
        # se banda vuota, prova a prendere un token singolo che combaci con pattern amount
        if not amount_raw:
            cand = words[
                (words["top"].between(top - 2.5, top + 2.5)) &
                (words["x0"] > pct_x_min) & (words["x0"] < client_x_min)
            ].sort_values("x0")
            for t in cand["text"].astype(str):
                if AMT_RE.match(t.strip()):
                    amount_raw = t.strip()
                    break
        if not amount_raw:
            amount_raw = "-"

        # client (testo a destra)
        client = text_in_band(words, top, client_x_min, float(words["x0"].max()) + 5.0, y_tol=2.5)

        # normalizza percentuale con punto
        pct_text = str(p["text"]).replace(",", ".")
        rows.append(
            {
                "PROJECT": proj,
                "MEMBER": mem,
                "DURATION": d["text"],
                "DURATION_%": pct_text,
                "AMOUNT": amount_raw,
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

        def collect_rows(pages_range) -> List[Dict]:
            rows: List[Dict] = []
            with pdfplumber.open(tmp.name) as pdf:
                for i in pages_range:
                    page = pdf.pages[i]
                    words = extract_words_page(page)
                    if not is_breakdown_page(words):
                        continue
                    rows.extend(parse_page(page))
            return rows

        with pdfplumber.open(tmp.name) as pdf:
            n = len(pdf.pages)

        # 1) tipico: tabella in pagina 2 (indice 1)
        all_rows = collect_rows(range(1, n)) if n > 1 else []
        # 2) fallback: prova tutte le pagine
        if not all_rows:
            all_rows = collect_rows(range(0, n))

    df = pd.DataFrame(all_rows, columns=["PROJECT","MEMBER","DURATION","DURATION_%","AMOUNT","CLIENT"])
    if not df.empty:
        # normalizza percentuali
        df["DURATION_%"] = df["DURATION_%"].astype(str).str.replace(",", ".", regex=False)
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
    "La ricerca parte da pagina 2 (fallback: tutte). "
    "Si estraggono **tutte le colonne**: Project, Member, Duration, Duration %, Amount, Client."
)

uploaded = st.file_uploader("Carica PDF", type=["pdf"], label_visibility="collapsed")

if uploaded:
    with st.spinner("⏳ Elaborazione del PDF in corso..."):
        df = process_pdf(uploaded.read())

    if df.empty:
        st.error("⚠️ Nessuna riga trovata. Assicurati che il PDF sia il report giusto.")
    else:
        st.success(f"✅ Righe estratte: {len(df)}")
        st.dataframe(df, use_container_width=True, height=380)

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

else:
    st.info("⬆️ Carica un file PDF per iniziare.")
