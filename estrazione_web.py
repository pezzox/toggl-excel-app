import streamlit as st
from io import BytesIO
import pandas as pd
import pdfplumber, re, tempfile

# ---------------- UI / PAGE ---------------- #
st.set_page_config(
    page_title="Toggl → Excel",
    page_icon="📊",
    layout="centered"
)

st.markdown("## 📊 Estrattore Toggl → Excel")
st.caption(
    "Carica il report PDF *Project & member breakdown* esportato da Toggl. "
    "La prima pagina (copertina/riassunto) può essere presente o meno: l'app analizza tutte le pagine "
    "e estrae **solo i progetti e i totali** (non i membri)."
)

# ---------------- PATTERN & HELPERS ---------------- #
DUR_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
PCT_RE = re.compile(r"^\d{1,3}(?:\.\d+)?%$")

def extract_words_page(page):
    """Ritaglia il corpo della pagina ed estrae le parole con coordinate."""
    w, h = page.width, page.height
    # lascia un margine esterno, evita intestazioni/piedi
    bbox = (w * 0.04, h * 0.06, w * 0.96, h * 0.95)
    body = page.crop(bbox)
    words = body.extract_words(
        x_tolerance=2.0,
        y_tolerance=2.5,
        keep_blank_chars=False,
        use_text_flow=True
    )
    return pd.DataFrame(words) if words else pd.DataFrame()

def left_text_near(words_df, top, x_limit=260.0, y_tol=5.0):
    """Testo sul lato sinistro (project / member) sulla stessa riga verticale della durata."""
    local = words_df[
        (words_df["top"] >= top - y_tol) &
        (words_df["top"] <= top + y_tol) &
        (words_df["x0"] < x_limit)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def classify_left(left: str):
    """Classifica: PROJECT/TOTAL vs MEMBER."""
    if not left:
        return None, None
    lo = left.lower()
    if lo.startswith("total"):
        return "TOTAL", None
    if lo.startswith("without"):
        return "Without project", None
    # Righe progetto hanno spesso "(1)" a fine riga → rimuovi conteggio
    if re.search(r"\(\d+\)\s*$", left):
        return re.sub(r"\s*\(\d+\)$", "", left).strip(), None
    # altrimenti è un membro
    return None, left

def guess_client_xmin(words_df):
    """
    Stima il bordo sinistro della colonna CLIENT:
    - se trova l'header 'CLIENT', usa la sua x0 - piccolo margine,
    - altrimenti usa il quantile alto delle x come fallback robusto.
    """
    mask_hdr = words_df["text"].astype(str).str.strip().str.lower().eq("client")
    if mask_hdr.any():
        return float(words_df.loc[mask_hdr, "x0"].min()) - 5.0
    return float(words_df["x0"].quantile(0.85))

def right_text_near(words_df, top, x_min, y_tol=5.0):
    """Testo nella fascia destra (CLIENT) sulla stessa riga verticale."""
    local = words_df[
        (words_df["top"] >= top - y_tol) &
        (words_df["top"] <= top + y_tol) &
        (words_df["x0"] >= x_min)
    ].sort_values("x0")
    if local.empty:
        return ""
    texts = local["text"].astype(str)
    # Escludi il '-' della colonna AMOUNT
    texts = texts[~texts.eq("-")]
    return " ".join(texts.tolist()).strip()

def parse_page(page):
    """Parsa una singola pagina del report."""
    words = extract_words_page(page)
    if words.empty:
        return []

    # token durata e percentuale (può non allinearsi perfettamente in verticale)
    dur_df = words[words["text"].astype(str).str.match(DUR_RE)].copy()
    pct_df = words[words["text"].astype(str).str.match(PCT_RE)].copy()
    if dur_df.empty or pct_df.empty:
        return []

    dur_df = dur_df.sort_values("top").reset_index(drop=True)
    pct_df = pct_df.sort_values("top").reset_index(drop=True)

    client_xmin = guess_client_xmin(words)

    rows = []
    for _, d in dur_df.iterrows():
        line_top = float(d["top"])

        # percentuale più vicina per top
        pct_df["__dist"] = (pct_df["top"].astype(float) - line_top).abs()
        p_row = pct_df.loc[pct_df["__dist"].idxmin()]
        p_txt = str(p_row["text"])

        left = left_text_near(words, line_top)
        proj, mem = classify_left(left)
        if proj is None and mem is None:
            continue

        client_text = right_text_near(words, line_top, client_xmin)

        rows.append({
            "PROJECT": proj,
            "MEMBER": mem,
            "DURATION": d["text"],
            "DURATION_%": p_txt,
            "AMOUNT": "-",
            "CLIENT": client_text if proj else ""
        })
    return rows

@st.cache_data(show_spinner=False)
def process_pdf(file_bytes: bytes) -> pd.DataFrame:
    """Apre il PDF, scorre tutte le pagine e restituisce un DataFrame pulito."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
        tmp.write(file_bytes)
        tmp.flush()

        all_rows = []
        with pdfplumber.open(tmp.name) as pdf:
            for i in range(len(pdf.pages)):  # analizza tutte le pagine
                all_rows.extend(parse_page(pdf.pages[i]))

    df = pd.DataFrame(all_rows)
    if not df.empty:
        # tieni solo progetti (no membri)
        keep = df["MEMBER"].isna() | (df["MEMBER"].astype(str).str.strip() == "")
        df = df[keep].copy()

        # normalizza percentuali (virgola → punto) e pulisci client
        df["DURATION_%"] = df["DURATION_%"].astype(str).str.replace(",", ".", regex=False)
        df["CLIENT"] = df["CLIENT"].astype(str).str.strip()

        # ordina come nel report
        df = df.reset_index(drop=True)

    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Converte il DataFrame in un file Excel in memoria."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="breakdown", index=False)
    return output.getvalue()

# ---------------- MAIN UI ---------------- #
uploaded = st.file_uploader("Carica il tuo PDF", type=["pdf"], label_visibility="collapsed")

if uploaded:
    with st.spinner("⏳ Elaborazione del PDF in corso..."):
        df = process_pdf(uploaded.read())

    if df.empty:
        st.error("⚠️ Nessuna riga trovata. Assicurati che il PDF sia il report giusto e che contenga la tabella.")
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

