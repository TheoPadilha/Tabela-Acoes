from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔹 Configuração do Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "norse-avatar-306414-f149a22bbcf1.json"  # Arquivo JSON da chave de serviço
SPREADSHEET_NAME = "tabela Ações"  # Nome da planilha no Google Sheets

# 🔹 Autenticação no Google Sheets
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).sheet1  # Abre a primeira aba da planilha

# 🔹 Inicializa o navegador Edge
service = Service()
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)
options.add_argument("--headless")  # Rodar sem abrir a janela do navegador
# 🔹 Lista de ações a serem consultadas
acoes = [
    "ABEV3", "BBAS3", "BBDC4", "BBSE3", "CMIG4", "EGIE3",
    "ITSA3", "KLBN4", "PETR4", "SAPR4", "TAEE11", "VALE3", "WEGE3"
]

# 🔹 Mapeamento específico para algumas ações
mapeamento_acoes = {
    "BBAS3": {232: "LPA", 224: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "BBDC4": {232: "LPA", 224: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "BBSE3": {224: "LPA", 216: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "KLBN4": {248: "LPA", 240: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"}
}

# Mapeamento padrão para ações sem um HTML específico
mapeamento_padrao = {248: "LPA", 240: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"}

# 🔹 Dividendos médios pagos por ação
dividendos_12m = {
    "ABEV3": 0.72,
    "BBAS3": 2.31,
    "BBDC4": 1.07,
    "BBSE3": 2.67,
    "CMIG4": 1.33,
    "EGIE3": 2.56,
    "ITSA3": 0.63,
    "KLBN4": 0.28,
    "PETR4": 10.55,
    "SAPR4": 0.28,
    "TAEE11": 4.08,
    "VALE3": 6.33,
    "WEGE3": 0.63
}

# 🔹 Apaga todos os dados na planilha antes de adicionar novos
sheet.clear()

# 🔹 Cabeçalhos da planilha
sheet.append_row(["TICKER", "PREÇO", "LPA", "VPA", "P/VP", "P/L", "DY", "PEG RATIO",
                  "Preço Teto (Graham)", "Preço Teto (Bazin - 6%)", "Preço Teto (Bazin - 8%)", 
                  "Preço Teto (Bazin - 10%)", "Preço Teto (Bazin - 12%)", 
                  "Preço Teto (Projetivo)", "Preço Teto (Peter Lynch)"])

# 🔹 Funções para calcular os preços teto
def calcular_preco_teto_graham(LPA, VPA):
    return (22.5 * LPA * VPA) ** 0.5

def calcular_preco_teto_bazin(dividendos_12m, rentabilidade):
    if dividendos_acao and rentabilidade:
        return dividendos_12m / rentabilidade
    return "N/A"

def calcular_preco_teto_projetivo(LPA, PL, DY):
    return (LPA * PL) / DY if LPA and PL and DY else "N/A"

def calcular_preco_teto_peter_lynch(LPA, PEG):
    return LPA * (PEG * 100) if LPA and PEG else "N/A"

# 🔹 Percorre todas as ações
for acao in acoes:
    print(f"\n--- Coletando dados para {acao} ---")

    # 🔸 1️⃣ Coleta dados do Investidor 10
    url_investidor10 = f"https://investidor10.com.br/acoes/{acao}"
    driver.get(url_investidor10)
    time.sleep(3)

    elementos = driver.find_elements(By.TAG_NAME, 'span')
    dados = {"TICKER": acao}

    # Usa o mapeamento específico se houver, senão usa o padrão
    mapa_indices = mapeamento_acoes.get(acao, mapeamento_padrao)

    for indice, rotulo in mapa_indices.items():
        if indice < len(elementos):
            dados[rotulo] = elementos[indice].text.strip()
        else:
            dados[rotulo] = "N/A"

    # 🔸 2️⃣ Coleta apenas o PEG Ratio do Status Invest
    url_statusinvest = f"https://statusinvest.com.br/acoes/{acao}"
    driver.get(url_statusinvest)
    time.sleep(3)

    elementos_status = driver.find_elements(By.TAG_NAME, 'strong')

    try:
        if acao == "TAEE11":
            dados["PEG RATIO"] = elementos_status[38].text.strip()  # Correção específica para TAEE11
        else:
            dados["PEG RATIO"] = elementos_status[36].text.strip()  # Para as outras ações, mantém o índice 36
    except:
        dados["PEG RATIO"] = "N/A"

    # 🔸 3️⃣ Conversão de valores para cálculos
    try:
        LPA = float(dados["LPA"].replace(",", ".")) if dados["LPA"] != "N/A" else None
        VPA = float(dados["VPA"].replace(",", ".")) if dados["VPA"] != "N/A" else None
        PL = float(dados["P/L"].replace(",", ".")) if dados["P/L"] != "N/A" else None
        DY = float(dados["DY"].replace(",", ".").replace("%", "")) / 100 if dados["DY"] != "N/A" else None
        PEG = float(dados["PEG RATIO"].replace(",", ".")) if dados["PEG RATIO"] != "N/A" else None
    except:
        LPA, VPA, PL, DY, PEG = None, None, None, None, None

    # 🔸 4️⃣ Calcula os preços teto
    preco_teto_graham = calcular_preco_teto_graham(LPA, VPA) if LPA and VPA else "N/A"
    
    # Agora o cálculo do Preço Teto Bazin usa os dividendos médios que você forneceu
    dividendos_acao = dividendos_12m.get(acao)
    
    # Calcula o Preço Teto Bazin para diferentes rentabilidades
    preco_teto_bazin_6 = calcular_preco_teto_bazin(dividendos_acao, 0.06)
    preco_teto_bazin_8 = calcular_preco_teto_bazin(dividendos_acao, 0.08)
    preco_teto_bazin_10 = calcular_preco_teto_bazin(dividendos_acao, 0.10)
    preco_teto_bazin_12 = calcular_preco_teto_bazin(dividendos_acao, 0.12)
    
    preco_teto_projetivo = calcular_preco_teto_projetivo(LPA, PL, DY) if LPA and PL and DY else "N/A"
    preco_teto_peter_lynch = calcular_preco_teto_peter_lynch(LPA, PEG) if LPA and PEG else "N/A"

    # 🔸 5️⃣ Adiciona os dados na planilha
    sheet.append_row([ 
        acao, dados["PREÇO"], dados["LPA"], dados["VPA"], dados["P/VP"], dados["P/L"], dados["DY"], dados["PEG RATIO"],
        preco_teto_graham, preco_teto_bazin_6, preco_teto_bazin_8, preco_teto_bazin_10, preco_teto_bazin_12, 
        preco_teto_projetivo, preco_teto_peter_lynch
    ])

    print(f"✅ Dados adicionados para {acao}")

    time.sleep(2)  # Pequena pausa entre requisições

# 🔹 Fecha o navegador
driver.quit()

print("\n🚀 Processo concluído! Os dados foram enviados para o Google Sheets.")
