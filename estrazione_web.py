import streamlit as st
from io import BytesIO
import pandas as pd
import pdfplumber, re, tempfile



st.set_page_config(
    page_title="Toggl → Excel",
    page_icon="📊",
    layout="centered"
)
# ---------- configurazione pagina ---------- # st.set_option('server.maxUploadSize', 100)  # ← toglila

# ---------- parsing ----------
DUR_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
PCT_RE = re.compile(r"^\d{1,3}(?:\.\d+)?%$")

def extract_words_page(page):
    w, h = page.width, page.height
    bbox = (w * 0.04, h * 0.06, w * 0.96, h * 0.95)
    body = page.crop(bbox)
    words = body.extract_words(x_tolerance=2.0, y_tolerance=2.5,
                               keep_blank_chars=False, use_text_flow=True)
    return pd.DataFrame(words) if words else pd.DataFrame()

def left_text_near(words_df, top, x_limit=200.0, y_tol=2.0):
    local = words_df[
        (words_df["top"] >= top - y_tol) &
        (words_df["top"] <= top + y_tol) &
        (words_df["x0"] < x_limit)
    ].sort_values("x0")
    return " ".join(local["text"].astype(str).tolist()).strip()

def client_text_near(words_df, top, y_tol=2.0):
    # Cerca il testo nella zona CLIENT (parte più a destra della tabella)
    if words_df.empty:
        return ""
    
    # Usa una soglia fissa più sicura
    local = words_df[
        (words_df["top"] >= top - y_tol) &
        (words_df["top"] <= top + y_tol) &
        (words_df["x0"] > 500)  # Cerca tutto quello che inizia oltre x=500
    ].sort_values("x0")
    
    client_text = " ".join(local["text"].astype(str).tolist()).strip()
    return client_text if client_text else ""

def classify_left(left: str):
    if not left: return None, None
    lo = left.lower()
    if lo.startswith("total"): return "TOTAL", None
    if lo.startswith("without"): return "Without project", None
    if re.search(r"\(\d+\)\s*$", left): 
        return re.sub(r"\s*\(\d+\)$", "", left).strip(), None
    return None, left  # membro

def parse_page(page):
    words = extract_words_page(page)
    if words.empty: return []
    dur = words[words["text"].astype(str).str.match(DUR_RE)].sort_values("top")
    pct = words[words["text"].astype(str).str.match(PCT_RE)].sort_values("top")
    rows = []
    for (_, d), (_, p) in zip(dur.iterrows(), pct.iterrows()):
        left = left_text_near(words, float(d["top"]))
        client = client_text_near(words, float(d["top"]))
        proj, mem = classify_left(left)
        if proj is None and mem is None: continue
        rows.append({
            "PROJECT": proj,
            "MEMBER": mem,
            "DURATION": d["text"],
            "DURATION_%": p["text"],
            "AMOUNT": "-",
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
        # tieni solo progetti (no membri)
        df = df[df["MEMBER"].isna() | (df["MEMBER"].astype(str).str.strip() == "")]
        df["DURATION_%"] = df["DURATION_%"].astype(str).str.replace(",", ".", regex=False)
    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="breakdown", index=False)
    return output.getvalue()

# ---------- UI ----------
st.markdown("## 📊 Estrattore Toggl → Excel")
st.caption("Carica il report PDF *Project & member breakdown* esportato da Toggl. "
           "La prima pagina viene ignorata automaticamente. Verranno estratti **solo i progetti e i totali**.")

uploaded = st.file_uploader("Carica il tuo PDF", type=["pdf"], label_visibility="collapsed")

if uploaded:
    with st.spinner("⏳ Elaborazione del PDF in corso..."):
        df = process_pdf(uploaded.read())

    if df.empty:
        st.error("⚠️ Nessuna riga trovata. Assicurati che il PDF sia il report giusto.")
    else:
        st.success(f"✅ Estratti {len(df)} progetti.")
        st.dataframe(df, use_container_width=True, height=350)

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
