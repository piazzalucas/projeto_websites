import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
from unidecode import unidecode
import re

# --- Configuração da página ---
st.set_page_config("Preencher Websites", layout="wide")
st.title("🔍 Preencher Websites por País")
st.markdown("""\
1. Faça upload da sua lista de empresas (.xlsx)  
2. Selecione o país  
3. Ajuste o nível de similaridade  
4. Clique em Processar e baixe o resultado  
""")

# --- Ordem de prioridade de Account Status ---
STATUS_IMPORTANCE = [
    'Lead',
    'Brand - Inactive',
    'Referral Partner - Inactive',
    'Tier 1 Reseller - Inactive',
    'Value Added Reseller - Inactiv',
    'Direct Customer - Inactive',
    'Indirect Customer - Inactive',
    'Multiple wo ISP - Inactive',
    'VAR Customer - Active',
    'Value Added Reseller - Active',
    'Tier 1 Reseller - Active',
    'Indirect Customer - Active',
    'ISP - Active',
    'Multiple wo ISP - Active',
    'Multiple w ISP - Active',
    'Direct Customer - Active',
    'NAP Master Agreement - Active',
    'Akamai Internal - Active'
]
STATUS_PRIORITY = {s: i for i, s in enumerate(STATUS_IMPORTANCE)}

# --- Normalização de nomes ---
STOPWORDS = {'sa', 'ltda', 'inc', 'corp', 'supply', 'international', 'group', 'empresa', 'company', 's', 'a'}
def normalize(name):
    text = unidecode(str(name)).lower()
    text = re.sub(r'[^a-z0-9 ]+', ' ', text)
    return " ".join(w for w in text.split() if w not in STOPWORDS)

# --- Carrega e prepara a base ---
@st.cache_data
def load_base():
    df = pd.read_csv("WEBSITES-COMPANYS-LATAM.csv", encoding="latin1")
    df['name_norm'] = df['Account Name'].map(normalize)
    # Mapeia prioridade de status, se a coluna existir
    if 'Account Status' in df.columns:
        df['status_prio'] = df['Account Status'].map(lambda s: STATUS_PRIORITY.get(s, -1))
    else:
        df['status_prio'] = 0
    return df
base = load_base()

# --- Seletor de país ---
paises = sorted(base['Primary Country'].dropna().unique())
pais = st.selectbox("🌎 País", [""] + paises)
if not pais:
    st.warning("Selecione um país para continuar.")
    st.stop()

# --- Upload da lista ---
uploaded = st.file_uploader("📂 Sua lista de empresas (Excel)", type="xlsx")
if not uploaded:
    st.stop()

# --- Slider de similaridade ---
threshold = st.slider("🔍 Similaridade mínima (%)", 50, 100, 85)

# --- Processamento ---
if st.button("▶️ Processar"):
    # Lê e normaliza entrada
    lista = pd.read_excel(uploaded, sheet_name=0)
    lista['name_norm'] = lista['EMPRESA'].map(normalize)

    # Filtra base pelo país
    bf = base[base['Primary Country'] == pais].copy()
    nomes_norm = bf['name_norm'].tolist()
    sites = bf['Website'].tolist()
    prios = bf['status_prio'].tolist()

    # Função de match com prioridade em Account Status
    def buscar_site(norm_name):
        # Fuzzy token_set_ratio: coleta todos acima do threshold
        matches = process.extract(norm_name, nomes_norm, scorer=fuzz.token_set_ratio, limit=None)
        # Filtra por score
        good = [(m, sc, idx) for m, sc, idx in matches if sc >= threshold]
        if good:
            # Encontra maior score
            max_sc = max(sc for _, sc, _ in good)
            # Filtra candidatos com esse score
            cands = [idx for _, sc, idx in good if sc == max_sc]
            # Seleciona aquele com maior prioridade de status
            best = max(cands, key=lambda i: prios[i])
            return sites[best]
        # Fallback partial_ratio
        matches2 = process.extract(norm_name, nomes_norm, scorer=fuzz.partial_ratio, limit=None)
        good2 = [(m, sc, idx) for m, sc, idx in matches2 if sc >= threshold]
        if good2:
            max_sc2 = max(sc for _, sc, _ in good2)
            cands2 = [idx for _, sc, idx in good2 if sc == max_sc2]
            best2 = max(cands2, key=lambda i: prios[i])
            return sites[best2]
        return 'nao encontrado'

    # Monta resultados
    resultados = []
    for orig, norm in zip(lista['EMPRESA'], lista['name_norm']):
        site = buscar_site(norm)
        # Top-3 sugestões
        sug = process.extract(norm, nomes_norm, scorer=fuzz.token_set_ratio, limit=3)
        top3 = [f"{bf['Account Name'].iat[idx]} ({score}%)" for _, score, idx in sug]
        resultados.append({'EMPRESA': orig, 'Website': site, 'Sugestoes': '; '.join(top3)})

    # Exporta para Excel
    df_out = pd.DataFrame(resultados)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name='Websites')
    buffer.seek(0)
    st.download_button("💾 Baixar Resultado", buffer,
                       file_name=f"{pais}_websites.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
