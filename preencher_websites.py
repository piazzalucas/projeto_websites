import pandas as pd
import difflib, re, os

# 1) Carrega a base de websites
base = pd.read_csv('WEBSITES-COMPANYS-LATAM.csv', encoding='latin1')

# 2) Mostra os países disponíveis
paises = sorted(base['Primary Country'].dropna().unique())
print("Países na base:")
for i, país in enumerate(paises, 1):
    print(f"{i:2d}. {país}")
escolha = int(input("Digite o número do país desejado: "))
pais_sel = paises[escolha-1]
print(f"> Você escolheu: {pais_sel}\n")

# 3) Pede o nome do arquivo de entrada
nome_lista = input("Nome do .xlsx em input_lists/ (ex: TESTE-ACCPLAN.xlsx): ")
caminho = os.path.join('input_lists', nome_lista)

# 4) Pede a sheet (ou Planilha1 por padrão)
sheet = input("Nome da sheet (Enter para 'Planilha1'): ") or 'Planilha1'

# 5) Lê a lista e limpa parênteses do nome
lista = pd.read_excel(caminho, sheet_name=sheet)
lista['EMPRESA_clean'] = (lista['EMPRESA']
    .astype(str)
    .str.replace(r"\s*\(.*\)", "", regex=True)
    .str.strip())

# 6) Filtra base pelo país e prepara lista de nomes
base_filtrada = base[base['Primary Country'] == pais_sel]
nomes_base = base_filtrada['Account Name'].tolist()

# 7) Função de fuzzy match
def buscar_site(nome, cutoff=0.7):
    m = difflib.get_close_matches(nome, nomes_base, n=1, cutoff=cutoff)
    return (base_filtrada.loc[
        base_filtrada['Account Name'] == m[0], 'Website'
    ].iloc[0] if m else 'nao encontrado')

# 8) Aplica e salva
lista['Website'] = lista['EMPRESA_clean'].apply(buscar_site)
resultado = lista[['EMPRESA', 'Website']]
out = f"{os.path.splitext(nome_lista)[0]}_{pais_sel}_websites.xlsx"
resultado.to_excel(out, index=False)

print(f"\n✔ Pronto! Arquivo gerado: {out}")
