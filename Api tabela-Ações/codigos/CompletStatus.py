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
# options.add_argument('--headless')  # Adiciona a opção para rodar em headless
driver = webdriver.Edge(service=service, options=options)

# 🔹 Lista de ações a serem consultadas
acoes = [
    "ABEV3", "BBAS3", "BBDC4", "BBSE3", "CMIG4", "EGIE3",
    "ITSA3", "KLBN4", "PETR4", "SAPR4", "TAEE11", "VALE3", "WEGE3"
]

# 🔹 Mapeamento dos índices para os rótulos desejados
mapa_indices_padrao = {
    4 :"PREÇO",
    44: "LPA",
    42: "VPA",
    37: "P/VP",
    35: "P/L",
    34: "DY",
    67: "CAGR LUCRO",
   
}

# Mapeamento específico para TAEE11
mapa_indices_taesa = {
    4 :"PREÇO",
    46: "LPA",
    44: "VPA",
    39: "P/VP",
    37: "P/L",
    36: "DY",
    69: "CAGR LUCRO"
}

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
sheet.append_row(["TICKER", "PREÇO", "LPA", "VPA", "P/VP", "P/L", "DY","PEG RATIO", 
                  "Preço Teto (Graham)", "Preço Teto (Bazin - 6%)", "Preço Teto (Bazin - 8%)", 
                  "Preço Teto (Bazin - 10%)", "Preço Teto (Bazin - 12%)", "Preço Teto (Projetivo)", "Preço Teto (Projetivo)6%",
                  "Preço Teto (Projetivo)8%","Preço Teto (Projetivo)10%","Preço Teto (Projetivo)12%"
                  "Preço Teto (Peter Lynch)"])

# 🔹 Funções para calcular os preços teto
def calcular_preco_teto_graham(LPA, VPA):
    if LPA and VPA:
        return (22.5 * LPA * VPA) ** 0.5
    return "N/A"

def calcular_preco_teto_bazin(dividendos_acao, rentabilidade):
    if dividendos_acao and rentabilidade:
        return dividendos_acao / rentabilidade
    return "N/A"

def calcular_preco_teto_projetivo(DPA, rentabilidade):
    if DPA:
        return DPA / rentabilidade
    return "N/A"


def calcular_preco_teto_peter_lynch(LPA, PEG_RATIO_CALCULADO, DY , PL):
    if LPA and PEG_RATIO_CALCULADO:
        return LPA + DY / PL
    return "N/A"

# 🔹 Percorre todas as ações
for acao in acoes:
    print(f"\n--- Coletando dados para {acao} ---")

    url_statusinvest = f"https://statusinvest.com.br/acoes/{acao}"
    driver.get(url_statusinvest)
    time.sleep(3)

    # Obtém todos os elementos <strong> da página
    elementos = driver.find_elements(By.TAG_NAME, 'strong')

    # Escolhe o mapeamento correto
    mapa_indices = mapa_indices_taesa if acao == "TAEE11" else mapa_indices_padrao

    dados = {"TICKER": acao}

    # Acessa elementos específicos pelos índices e salva os dados
    for indice, rotulo in mapa_indices.items():
        if indice < len(elementos):
            dados[rotulo] = elementos[indice].text.strip()
        else:
            dados[rotulo] = "N/A"

    # 🔸 Conversão de valores para cálculos
    try:
        LPA = float(dados["LPA"].replace(",", ".")) if dados["LPA"] != "N/A" else None
        VPA = float(dados["VPA"].replace(",", ".")) if dados["VPA"] != "N/A" else None
        PL = float(dados["P/L"].replace(",", ".")) if dados["P/L"] != "N/A" else None
        DY = float(dados["DY"].replace(",", ".").replace("%", "")) / 100 if dados["DY"] != "N/A" else None
        PEG = float(dados["CAGR LUCRO"].replace(",", ".").replace("%", "")) / 100 if dados["CAGR LUCRO"] != "N/A" else None

        
        PEG_RATIO_CALCULADO = (PL / PEG) if (PL is not None and PEG is not None and PEG != 0) else "N/A"


    except Exception as e:
        print(f"Erro ao converter valores para {acao}: {e}")
        LPA, VPA, PL, DY, PEG_RATIO_CALCULADO = None, None, None, None, "N/A"


    # 🔸 Calcula os preços teto
    preco_teto_graham = calcular_preco_teto_graham(LPA, VPA)
    
    # Agora o cálculo do Preço Teto Bazin usa os dividendos médios que você forneceu
    dividendos_acao = dividendos_12m.get(acao, 0)
    
    # Calcula o Preço Teto Bazin para diferentes rentabilidades
    preco_teto_bazin_6 = calcular_preco_teto_bazin(dividendos_acao, 0.06)
    preco_teto_bazin_8 = calcular_preco_teto_bazin(dividendos_acao, 0.08)
    preco_teto_bazin_10 = calcular_preco_teto_bazin(dividendos_acao, 0.10)
    preco_teto_bazin_12 = calcular_preco_teto_bazin(dividendos_acao, 0.12)
    
    preco_teto_projetivo_6 = calcular_preco_teto_projetivo(dividendos_acao, 0.06)
    preco_teto_projetivo_8 = calcular_preco_teto_projetivo(dividendos_acao, 0.08)
    preco_teto_projetivo_10 = calcular_preco_teto_projetivo(dividendos_acao, 0.10)
    preco_teto_projetivo_12 = calcular_preco_teto_projetivo(dividendos_acao, 0.12)

    preco_teto_peter_lynch = calcular_preco_teto_peter_lynch(LPA, PEG_RATIO_CALCULADO, DY , PL)  # Considerando que não há PEG no StatusInvest

    # 🔸 Adiciona os dados na planilha
   
    sheet.append_row([ 
        acao, dados["PREÇO"], dados["LPA"], dados["VPA"], dados["P/VP"], dados["P/L"], dados["DY"], PEG_RATIO_CALCULADO,
        preco_teto_graham, preco_teto_bazin_6, preco_teto_bazin_8, preco_teto_bazin_10, preco_teto_bazin_12, 
        preco_teto_projetivo_6, preco_teto_projetivo_8, preco_teto_projetivo_10, preco_teto_projetivo_12, 
        preco_teto_peter_lynch
    ])



    print(f"✅ Dados adicionados para {acao}")

    time.sleep(2)  # Pequena pausa entre requisições

# 🔹 Fecha o navegador
driver.quit()

print("\n🚀 Processo concluído! Os dados foram enviados para o Google Sheets.")
