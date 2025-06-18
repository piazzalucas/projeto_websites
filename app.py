import streamlit as st
import pandas as pd
import io
from unidecode import unidecode
import re
from rapidfuzz import process, fuzz

# --- Configuração da página ---
st.set_page_config("Account Plan Intelligence", layout="wide")
st.title("🔍 Account Plan Intelligence")
st.markdown("""\
1. Faça upload da sua lista de empresas (.xlsx)  
2. Selecione o país  
3. Ajuste o nível de similaridade  
4. Faça upload dos relatórios NEWACC, WAFWON, WAFOPPS, APIWON, APIOPPS, GCWON e GCOPPS  
5. Clique em Processar e baixe o relatório completo  
""")

# --- Normalização de nomes ---
STOPWORDS = {'sa','ltda','inc','corp','supply','international','group','empresa','company','s','a'}
def normalize(name):
    text = unidecode(str(name)).lower()
    text = re.sub(r'[^a-z0-9 ]+', ' ', text)
    return " ".join(w for w in text.split() if w not in STOPWORDS)

# --- Carrega base de websites ---
@st.cache_data
def load_base():
    df = pd.read_csv("WEBSITES-COMPANYS-LATAM.csv", encoding="latin1")
    df['name_norm'] = df['Account Name'].map(normalize)
    return df
base = load_base()

# --- Seletor de país e lista de empresas ---
pais = st.selectbox("🌎 País", [""] + sorted(base['Primary Country'].dropna().unique()))
uploaded_list = st.file_uploader("📂 Lista de empresas (.xlsx)", type="xlsx")
threshold = st.slider("🔍 Similaridade mínima (%)", 50, 100, 85)

# --- Uploads de relatórios auxiliares ---
newacc_file   = st.file_uploader("🗂️ NEWACC (.xlsx/.csv)",   type=["xlsx","csv"])
wafwon_file  = st.file_uploader("🗂️ WAFWON (.xlsx/.csv)",  type=["xlsx","csv"])
wafopps_file = st.file_uploader("🗂️ WAFOPPS (.xlsx/.csv)", type=["xlsx","csv"])
apiwon_file  = st.file_uploader("🗂️ APIWON (.xlsx/.csv)",  type=["xlsx","csv"])
apiopps_file = st.file_uploader("🗂️ APIOPPS (.xlsx/.csv)", type=["xlsx","csv"])
gcwon_file   = st.file_uploader("🗂️ GCWON (.xlsx/.csv)",   type=["xlsx","csv"])
gcopps_file  = st.file_uploader("🗂️ GCOPPS (.xlsx/.csv)",  type=["xlsx","csv"])

# --- Processamento ---
if st.button("▶️ Processar"):
    if not pais or not uploaded_list:
        st.error("Selecione país e faça upload da lista de empresas.")
        st.stop()

    # 1. Processa lista de empresas
    df_list = pd.read_excel(uploaded_list)
    df_list['name_norm'] = df_list['EMPRESA'].map(normalize)

    # --- Carrega relatórios ---
    def load_report(f):
        if f is None:
            return pd.DataFrame()
        if f.name.lower().endswith('.csv'):
            return pd.read_csv(f)
        return pd.read_excel(f)

    newacc_df   = load_report(newacc_file)
    wafwon_df   = load_report(wafwon_file)
    wafopps_df  = load_report(wafopps_file)
    apiwon_df   = load_report(apiwon_file)
    apiopps_df  = load_report(apiopps_file)
    gcwon_df    = load_report(gcwon_file)
    gcopps_df   = load_report(gcopps_file)

    # Filtra base pelo país
    base_f = base[base['Primary Country'] == pais].copy()

    # --- Lookups por posição de coluna ---
    # NEWACC: website E(4), owner G(6), status J(9)
    def lookup_newacc_owner(w):
        df = newacc_df
        mask = df.iloc[:,4] == w
        return df.loc[mask].iloc[0,6] if mask.any() else None
    def lookup_newacc_status(w):
        df = newacc_df
        mask = df.iloc[:,4] == w
        return df.loc[mask].iloc[0,9] if mask.any() else None

    # OPP lookups: website D(3), opp name E(4)
    def lookup_opp(df, w):
        mask = df.iloc[:,3] == w
        return df.loc[mask].iloc[0,4] if mask.any() else None

    # 2. Preencher Account Owner / Status
    df_list['Account Owner']  = df_list['Website'].map(lookup_newacc_owner)
    df_list['Account Status'] = df_list['Website'].map(lookup_newacc_status)

    # 3. Preencher opportunities
    df_list['WAFWON']  = df_list['Website'].map(lambda w: lookup_opp(wafwon_df,  w))
    df_list['WAFOPPS'] = df_list['Website'].map(lambda w: lookup_opp(wafopps_df, w))
    df_list['APIWON']  = df_list['Website'].map(lambda w: lookup_opp(apiwon_df,  w))
    df_list['APIOPPS'] = df_list['Website'].map(lambda w: lookup_opp(apiopps_df, w))
    df_list['GCWON']   = df_list['Website'].map(lambda w: lookup_opp(gcwon_df,   w))
    df_list['GCOPPS']  = df_list['Website'].map(lambda w: lookup_opp(gcopps_df,  w))

    # 4. Status por produto
    def prod_status(r, won, opps):
        if pd.notna(r[won]):  return 'Customer'
        if pd.notna(r[opps]): return 'Partner'
        return 'Free'

    df_list['WAF Status'] = df_list.apply(lambda r: prod_status(r,'WAFWON','WAFOPPS'), axis=1)
    df_list['API Status'] = df_list.apply(lambda r: prod_status(r,'APIWON','APIOPPS'), axis=1)
    df_list['GC  Status'] = df_list.apply(lambda r: prod_status(r,'GCWON','GCOPPS'), axis=1)

    # 5. Exporta resultado
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_list.to_excel(writer, index=False, sheet_name='Account Plan')
    buffer.seek(0)
    st.download_button(
        "💾 Baixar Account Plan Completo",
        buffer,
        file_name=f"AccountPlan_{pais}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

