from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time  

service = Service()  
options = webdriver.EdgeOptions()

# Inicializa o navegador
driver = webdriver.Edge(service=service, options=options)

# Lista de ações que serão consultadas
acoes = [
    "VALE3"]

# Mapeamento dos índices para os rótulos desejados
mapa_indices_padrao = {
    4 :"PREÇO",
    44: "LPA",
    42: "VPA",
    37: "P/VP",
    35: "P/L",
    34: "DY",
    67: "PEG RATIO"
}

# Mapeamento específico para TAEE11
mapa_indices_taesa = {
    4 :"PREÇO",
    46: "LPA",
    44: "VPA",
    39: "P/VP",
    37: "P/L",
    36: "DY",
    38: "PEG RATIO"
}

# Percorre todas as ações e extrai os dados
for acao in acoes:
    url = f'https://statusinvest.com.br/acoes/{acao}'
    driver.get(url)
    
    # Espera um tempo para garantir que a página carregue
    time.sleep(3)

    # Obtém todos os elementos <strong> da página
    elementos = driver.find_elements(By.TAG_NAME, 'strong')

    print(f"\n===== {acao} =====")
    
    # Escolhe o mapeamento correto
    mapa_indices = mapa_indices_taesa if acao == "TAEE11" else mapa_indices_padrao

    # Acessa elementos específicos pelos índices e imprime com o rótulo correspondente
    for indice, rotulo in mapa_indices.items():
        if indice < len(elementos):  # Verifica se o índice é válido
            print(f"{rotulo}: {elementos[indice].text}")
        else:
            print(f"{rotulo}: Informação não encontrada")
    
    time.sleep(2)  
# elementos2 = driver.find_elements(By.TAG_NAME, 'strong')

# # Extrai o texto de cada elemento
# textos = [elemento3.text for elemento3 in elementos2]
# print(textos)
# Fecha o navegador após coletar todas as informações
driver.quit()
