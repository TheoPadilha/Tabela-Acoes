from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔹 Configuração do Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "norse-avatar-306414-f149a22bbcf1.json"
SPREADSHEET_NAME = "tabela Ações"

# 🔹 Autenticação no Google Sheets
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).sheet1

# 🔹 Inicializa o navegador Edge
service = Service()  # Se necessário, passe o caminho do WebDriver: Service("caminho/para/msedgedriver.exe")
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)

# 🔹 Lista de ações a serem consultadas
acoes = [ "ABEV3", "BBAS3", "BBDC4", "BBSE3", "CMIG4", "EGIE3",
    "ITSA3", "KLBN4", "PETR4", "SAPR4", "TAEE11", "VALE3", "WEGE3"]

# 🔹 Mapeamento dos índices para os rótulos desejados
mapeamento_acoes = {
     "BBAS3": {232: "LPA", 224: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "BBDC4": {232: "LPA", 224: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "BBSE3": {224: "LPA", 216: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"},
    "TAEE11": {248: "LPA", 240: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO"}
}

mapeamento_padrao = {248: "LPA", 240: "VPA", 136: "PEG RATIO", 48: "P/VP", 46: "P/L", 50: "DY", 42: "PREÇO", 352: "CAGR LUCRO"}
print(mapeamento_padrao)
# 🔹 Apaga todos os dados na planilha
sheet.clear()

# 🔹 Cabeçalhos da planilha
sheet.append_row(["TICKER", "PREÇO", "LPA", "VPA", "P/VP", "P/L", "DY", "PEG RATIO", 
                  "Preço Teto (Graham)", "Preço Teto (Bazin - 6%)", "Preço Teto (Bazin - 8%)", 
                  "Preço Teto (Bazin - 10%)", "Preço Teto (Bazin - 12%)", "Preço Teto (Projetivo 6%)",
                  "Preço Teto (Projetivo 8%)", "Preço Teto (Projetivo 10%)", "Preço Teto (Projetivo 12%)",
                  "Preço Teto (Peter Lynch)"])

# 🔹 Funções para calcular os preços teto
def calcular_preco_teto_graham(LPA, VPA):
    return (22.5 * LPA * VPA) ** 0.5 if LPA and VPA else "N/A"

def calcular_preco_teto_bazin(dividendos_acao, rentabilidade):
    return dividendos_acao / rentabilidade if dividendos_acao and rentabilidade else "N/A"

def calcular_preco_teto_peter_lynch(LPA, PEG_RATIO_CALCULADO, DY, PL):
    return LPA + DY / PL if LPA and PEG_RATIO_CALCULADO and PL else "N/A"

# 🔹 Percorre todas as ações
for acao in acoes:
    print(f"\n--- Coletando dados para {acao} ---")

    url_statusinvest = f"https://investidor10.com.br/acoes/{acao}"
    driver.get(url_statusinvest)
    time.sleep(3)

    elementos = driver.find_elements(By.TAG_NAME, 'span')
    mapa_indices = mapeamento_acoes.get(acao, mapeamento_padrao)

    dados = {"TICKER": acao}

    for indice, rotulo in mapeamento_padrao.items():
        dados[rotulo] = elementos[indice].text.strip() if indice < len(elementos) else "N/A"

    try:
        LPA = float(dados["LPA"].replace(",", ".")) if dados["LPA"] != "N/A" else None
        VPA = float(dados["VPA"].replace(",", ".")) if dados["VPA"] != "N/A" else None
        PL = float(dados["P/L"].replace(",", ".")) if dados["P/L"] != "N/A" else None
        DY = float(dados["DY"].replace(",", ".").replace("%", "")) / 100 if dados["DY"] != "N/A" else None
        PEG = float(dados.get("CAGR LUCRO", "0").replace(",", ".").replace("%", "")) / 100 * 100 if dados.get("CAGR LUCRO") != "N/A" else None
        print(f"CAGR LUCRO para {acao}: {PEG}")


        PEG_RATIO_CALCULADO = (PL / PEG) if (PL is not None and PEG is not None and PEG != 0) else "N/A"

    except Exception as e:
        print(f"Erro ao converter valores para {acao}: {e}")
        LPA, VPA, PL, DY, PEG_RATIO_CALCULADO = None, None, None, None, "N/A"

    preco_teto_graham = calcular_preco_teto_graham(LPA, VPA)
    dividendos_acao = 0.72  # Simulação de dividendo
    preco_teto_bazin_6 = calcular_preco_teto_bazin(dividendos_acao, 0.06)
    preco_teto_bazin_8 = calcular_preco_teto_bazin(dividendos_acao, 0.08)
    preco_teto_bazin_10 = calcular_preco_teto_bazin(dividendos_acao, 0.10)
    preco_teto_bazin_12 = calcular_preco_teto_bazin(dividendos_acao, 0.12)
    
    preco_teto_peter_lynch = calcular_preco_teto_peter_lynch(LPA, PEG_RATIO_CALCULADO, DY, PL)

    sheet.append_row([
        acao, dados["PREÇO"], dados["LPA"], dados["VPA"], dados["P/VP"], dados["P/L"], dados["DY"], PEG_RATIO_CALCULADO,
        preco_teto_graham, preco_teto_bazin_6, preco_teto_bazin_8, preco_teto_bazin_10, preco_teto_bazin_12,
        preco_teto_bazin_6, preco_teto_bazin_8, preco_teto_bazin_10, preco_teto_bazin_12,
        preco_teto_peter_lynch
    ])

    print(f"✅ Dados adicionados para {acao}")
    time.sleep(2)

driver.quit()
print("\n🚀 Processo concluído! Os dados foram enviados para o Google Sheets.")
