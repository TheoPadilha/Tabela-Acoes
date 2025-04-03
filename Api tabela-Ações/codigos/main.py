from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuração do Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "norse-avatar-306414-f149a22bbcf1.json"  # Nome do seu arquivo JSON da chave de serviço
SPREADSHEET_NAME = "tabela Ações"  # Nome da sua planilha no Google Sheets

# Autentica no Google Sheets
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).sheet1  # Abre a primeira aba da planilha

# Inicializa o navegador Edge
service = Service()
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)

# Lista de ações que serão consultadas
acoes = [
    "ABEV3", "BBAS3", "BBDC4", "BBSE3", "CMIG4", "EGIE3", 
    "ITSA3", "KLBN4", "PETR4", "SAPR4", "TAEE11", "VALE3", "WEGE3"
]

# Mapeamento dos índices para os rótulos desejados
mapa_indices_padrao = {
    4: "PREÇO",
    44: "LPA",
    42: "VPA",
    37: "P/VP",
    35: "P/L",
    34: "DY",
    36: "PEG RATIO",
}

# Mapeamento específico para TAEE11
mapa_indices_taesa = {
    4: "PREÇO",
    46: "LPA",
    44: "VPA",
    39: "P/VP",
    37: "P/L",
    36: "DY",
    38: "PEG RATIO",
}

# Cabeçalhos da planilha
sheet.append_row(["TICKER","PREÇO", "LPA", "VPA", "P/VP", "P/L", "DY", "PEG RATIO",
                  "Preço Teto (Graham)", "Preço Teto (Bazin)", "Preço Teto (Projetivo)", "Preço Teto (Peter Lynch)"])

# Funções para calcular os preços teto
def calcular_preco_teto_graham(LPA, VPA):
    return (22.5 * LPA * VPA) ** 0.5

def calcular_preco_teto_bazin(LPA, PL, DY):
    return (LPA * (8.5 + 2 * PL)) / (1 - DY)

def calcular_preco_teto_projetivo(LPA, PL, DY):
    return (LPA * PL) / DY

def calcular_preco_teto_peter_lynch(LPA, PEG):
    return LPA * (PEG * 100)

# Percorre todas as ações e extrai os dados
for acao in acoes:
    url = f'https://statusinvest.com.br/acoes/{acao}'
    driver.get(url)
    
    time.sleep(3)  # Espera a página carregar

    elementos = driver.find_elements(By.TAG_NAME, 'strong')

    # Define qual mapeamento de índices usar
    mapa_indices = mapa_indices_taesa if acao == "TAEE11" else mapa_indices_padrao

    # Coleta os dados
    dados = [acao]  # Começa com o nome da ação
    LPA, VPA, PL, DY, preco_acao, PEG = None, None, None, None, None, None

    for indice in mapa_indices:
        if indice < len(elementos):
            valor = elementos[indice].text
            dados.append(valor)
            # Atribui os valores necessários para os cálculos
            if mapa_indices[indice] == "LPA":
                LPA = float(valor.replace(",", "."))
            elif mapa_indices[indice] == "VPA":
                VPA = float(valor.replace(",", "."))
            elif mapa_indices[indice] == "P/L":
                PL = float(valor.replace(",", "."))
            elif mapa_indices[indice] == "DY":
                DY = float(valor.replace(",", ".").replace("%", "")) / 100
            elif mapa_indices[indice] == "PREÇO":
                preco_acao = float(valor.replace(",", "."))
            elif mapa_indices[indice] == "PEG RATIO":
                PEG = float(valor.replace(",", "."))

    # Calcular os preços teto
    preco_teto_graham = calcular_preco_teto_graham(LPA, VPA) if LPA and VPA else "N/A"
    preco_teto_bazin = calcular_preco_teto_bazin(LPA, PL, DY) if LPA and PL and DY else "N/A"
    preco_teto_projetivo = calcular_preco_teto_projetivo(LPA, PL, DY) if LPA and PL and DY else "N/A"
    preco_teto_peter_lynch = calcular_preco_teto_peter_lynch(LPA, PEG) if LPA and PEG else "N/A"

    # Adiciona os preços teto ao final dos dados
    dados.append(preco_teto_graham)
    dados.append(preco_teto_bazin)
    dados.append(preco_teto_projetivo)
    dados.append(preco_teto_peter_lynch)

    # Adiciona os dados na planilha
    sheet.append_row(dados)
    print(f"Dados adicionados para {acao}")

    time.sleep(2)  # Pequena pausa entre as requisições

# Fecha o navegador após coletar todas as informações
driver.quit()

print("Processo concluído! Os dados foram enviados para o Google Sheets.")
