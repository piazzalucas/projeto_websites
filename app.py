import streamlit as st
import pandas as pd
from pathlib import Path
from rapidfuzz import process, fuzz
import io
from unidecode import unidecode
import re

# --- Page config ---
st.set_page_config("Account Plan Intelligence", layout="wide")
st.title("🔍 Account Plan Intelligence")
st.markdown("""\
1. Faça upload da sua lista de empresas (.xlsx)  
2. Selecione a coluna que contém o nome das empresas  
3. Selecione o país  
4. Ajuste o nível de similaridade  
5. Clique em Processar e baixe o relatório completo  
""")

# --- Normalize names ---
STOPWORDS = {'sa','ltda','inc','corp','supply','international','group','empresa','company','s','a'}
def normalize(name):
    text = unidecode(str(name)).lower()
    text = re.sub(r'[^a-z0-9 ]+', ' ', text)
    return " ".join(w for w in text.split() if w not in STOPWORDS)

# --- Load website base ---
@st.cache_data
def load_base():
    df = pd.read_csv("WEBSITES-COMPANYS-LATAM.csv", encoding="latin1")
    df['name_norm'] = df['Account Name'].map(normalize)
    df['Primary Country'] = df['Primary Country'].fillna('')
    return df
base = load_base()

# --- Load backend reports ---
@st.cache_data
def load_reports():
    data_folder = Path(__file__).parent / "data"
    return {
        'newacc':  pd.read_excel(data_folder / "newacc.xlsx"),
        'wafwon':  pd.read_excel(data_folder / "wafwon.xlsx"),
        'wafopps': pd.read_excel(data_folder / "wafopps.xlsx"),
        'apiwon':  pd.read_excel(data_folder / "apiwon.xlsx"),
        'apiopps': pd.read_excel(data_folder / "apiopps.xlsx"),
        'gcwon':   pd.read_excel(data_folder / "gcwon.xlsx"),
        'gcopps':  pd.read_excel(data_folder / "gcopps.xlsx"),
    }
reports = load_reports()

# --- Lookup functions ---
def lookup_newacc_owner(website):
    df = reports['newacc']
    mask = df.iloc[:,4] == website
    return df.loc[mask].iloc[0,6] if mask.any() else None

def lookup_newacc_status(website):
    df = reports['newacc']
    mask = df.iloc[:,4] == website
    return df.loc[mask].iloc[0,9] if mask.any() else None

def lookup_opp(key, website):
    df = reports[key]
    mask = df.iloc[:,3] == website
    return df.loc[mask].iloc[0,4] if mask.any() else None

# --- Matching preparation ---
@st.cache_data
def prepare_matching(country):
    df = base[base['Primary Country'] == country]
    return df['name_norm'].tolist(), df['Website'].tolist()

def buscar_site(norm_name, names, sites, threshold):
    # exact
    for nm, site in zip(names, sites):
        if norm_name == nm:
            return site
    # fuzzy token_set_ratio
    m1 = process.extractOne(norm_name, names, scorer=fuzz.token_set_ratio)
    if m1 and m1[1] >= threshold:
        return sites[names.index(m1[0])]
    # fuzzy partial_ratio
    m2 = process.extractOne(norm_name, names, scorer=fuzz.partial_ratio)
    if m2 and m2[1] >= threshold:
        return sites[names.index(m2[0])]
    return None

# --- UI Inputs ---
uploaded_list = st.file_uploader("📂 Sua lista de empresas (.xlsx)", type="xlsx")
if uploaded_list is None:
    st.stop()
# preview columns
df_preview = pd.read_excel(uploaded_list, nrows=0)
colunas = list(df_preview.columns)
col_sel = st.selectbox("Selecione a coluna de empresas", colunas)
threshold = st.slider("🔍 Similaridade mínima (%)", 50, 100, 85)
pais = st.selectbox("🌎 País", [""] + sorted(base['Primary Country'].unique()))

if st.button("▶️ Processar"):
    if not pais:
        st.error("Selecione o país.")
        st.stop()
    # load full df
    df_list = pd.read_excel(uploaded_list)
    # rename selected column to EMPRESA
    df_list.rename(columns={col_sel: 'EMPRESA'}, inplace=True)
    df_list['name_norm'] = df_list['EMPRESA'].map(normalize)

    # matching websites
    names_norm, sites_norm = prepare_matching(pais)
    df_list['Website'] = df_list['name_norm'].apply(
        lambda x: buscar_site(x, names_norm, sites_norm, threshold)
    )
    # lookup newacc
    df_list['Account Owner'] = df_list['Website'].map(lookup_newacc_owner)
    df_list['Account Status'] = df_list['Website'].map(lookup_newacc_status)
    # lookup opps
    for key in ['wafwon','wafopps','apiwon','apiopps','gcwon','gcopps']:
        df_list[key.upper()] = df_list['Website'].map(lambda w, k=key: lookup_opp(k, w))
    # product status
    def prod_status(row, won_col, opp_col):
        if pd.notna(row[won_col]): return 'Customer'
        if pd.notna(row[opp_col]): return 'Partner'
        return 'Free'
    df_list['WAF Status'] = df_list.apply(lambda r: prod_status(r,'WAFWON','WAFOPPS'), axis=1)
    df_list['API Status'] = df_list.apply(lambda r: prod_status(r,'APIWON','APIOPPS'), axis=1)
    df_list['GC Status']  = df_list.apply(lambda r: prod_status(r,'GCWON','GCOPPS'), axis=1)

    # export
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_list.to_excel(writer, index=False, sheet_name='Account Plan')
    buffer.seek(0)
    st.download_button("💾 Baixar Relatório Completo", buffer,
                       file_name=f"AccountPlan_{pais}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
