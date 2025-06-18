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
2. Selecione o país  
3. Ajuste o nível de similaridade  
4. Clique em Processar e baixe o relatório completo  
""")

# --- normalize strings ---
STOPWORDS = {'sa','ltda','inc','corp','supply','international','group','empresa','company','s','a'}
def normalize(name):
    text = unidecode(str(name)).lower()
    text = re.sub(r'[^a-z0-9 ]+', ' ', text)
    return " ".join(w for w in text.split() if w not in STOPWORDS)

# --- load websites base ---
@st.cache_data
def load_base():
    df = pd.read_csv("WEBSITES-COMPANYS-LATAM.csv", encoding="latin1")
    df['name_norm'] = df['Account Name'].map(normalize)
    df['Primary Country'] = df['Primary Country'].fillna('')
    return df

base = load_base()

# --- load fixed backend reports from data/ ---
@st.cache_data
def load_reports():
    data_folder = Path(__file__).parent / "data"
    return {
        'newacc':  pd.read_excel(data_folder / "NEWACC-MONTAGEM.xlsx"),
        'wafwon':  pd.read_excel(data_folder / "WAFWON-MONTAGEM.xlsx"),
        'wafopps': pd.read_excel(data_folder / "WAFOPPS-MONTAGEM.xlsx"),
        'apiwon':  pd.read_excel(data_folder / "APIWON-MONTAGEM.xlsx"),
        'apiopps': pd.read_excel(data_folder / "APIOPPS-MONTAGEM.xlsx"),
        'gcwon':   pd.read_excel(data_folder / "GCWON-MONTAGEM.xlsx"),
        'gcopps':  pd.read_excel(data_folder / "GCOPPS-MONTAGEM.xlsx"),
    }

reports = load_reports()

# --- lookup functions ---
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

# --- matching preparation ---
@st.cache_data
def prepare_matching(country):
    df = base[base['Primary Country'] == country]
    return df['name_norm'].tolist(), df['Website'].tolist()

def buscar_site(norm_name, names, sites, threshold):
    # exact match
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
threshold      = st.slider("🔍 Similaridade mínima (%)", 50, 100, 85)
pais           = st.selectbox("🌎 País", [""] + sorted(base['Primary Country'].unique()))

if st.button("▶️ Processar"):
    if not uploaded_list or not pais:
        st.error("Faça upload da lista e selecione o país.")
        st.stop()

    # 1) Load and normalize your list
    df_list = pd.read_excel(uploaded_list)
    df_list['name_norm'] = df_list['EMPRESA'].map(normalize)

    # 2) Prepare matching for selected country
    names_norm, sites_norm = prepare_matching(pais)

    # 3) Fill Website column
    df_list['Website'] = df_list['name_norm'].apply(
        lambda x: buscar_site(x, names_norm, sites_norm, threshold)
    )

    # 4) Lookup NEWACC owner & status
    df_list['Account Owner']  = df_list['Website'].map(lookup_newacc_owner)
    df_list['Account Status'] = df_list['Website'].map(lookup_newacc_status)

    # 5) Lookup opportunities for each product
    for key in ['wafwon','wafopps','apiwon','apiopps','gcwon','gcopps']:
        df_list[key.upper()] = df_list['Website'].map(lambda w, k=key: lookup_opp(k, w))

    # 6) Compute product status columns
    def prod_status(row, won_col, opp_col):
        if pd.notna(row[won_col]): return 'Customer'
        if pd.notna(row[opp_col]): return 'Partner'
        return 'Free'

    df_list['WAF Status'] = df_list.apply(lambda r: prod_status(r, 'WAFWON','WAFOPPS'), axis=1)
    df_list['API Status'] = df_list.apply(lambda r: prod_status(r, 'APIWON','APIOPPS'), axis=1)
    df_list['GC Status']  = df_list.apply(lambda r: prod_status(r, 'GCWON','GCOPPS'), axis=1)

    # 7) Export to Excel and provide download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_list.to_excel(writer, index=False, sheet_name='Account Plan')
    buffer.seek(0)

    st.download_button(
        "💾 Baixar Relatório Completo",
        buffer,
        file_name=f"AccountPlan_{pais}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


