import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date
from sqlalchemy import create_engine
import urllib

print("Inicializando o navegador...")
options = Options()
options.binary_location = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
caminho_driver = r"C:\Users\adm.luiz.vinicius\Downloads\chromedriver.exe"

print("Inicializando o navegador manualmente...")
navegador = webdriver.Chrome(
    service=Service(caminho_driver),
    options=options
)

# Definir a URL e as credenciais
url = 'https://www.app-indecx.com/'
login = "ericka.nascimento@parvi.com.br"
senha = "Ericka@123"

print("Navegando até a URL...")
navegador.get(url)
navegador.maximize_window()

print("Realizando login...")
campo_login = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#email")))
campo_login.send_keys(login)

campo_senha = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#password")))
campo_senha.send_keys(senha)

botaologin = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn-block > span:nth-child(1)")))
botaologin.click()

print("Aplicando filtros...")
filtro = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#app > section > section > aside > div > ul > li:nth-child(2) > span > span")))
filtro.click()
time.sleep(6)

print("Selecionando NPS...")
cliqueNPS = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#pdf-dashboard > div > div.session__side-bar > ul > li:nth-child(2) > span > span")))
cliqueNPS.click()
time.sleep(6)

print("Fechando filtro...")
Cliqueparafecharfiltro = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#pdf-dashboard > div > div.session__side-bar > button")))
Cliqueparafecharfiltro.click()
time.sleep(6)

# Extrair o texto do elemento
neutro_boavista = WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#panel-main")))
texto_extraido = neutro_boavista.text
time.sleep(2)

# Processar o texto extraído
linhas = texto_extraido.split('\n')  # Dividir o texto em linhas

# Lista de índices das linhas que queremos pegar
indices = [1, 7, 13, 19, 26, 32, 38, 44, 50, 56, 63, 70, 77, 84, 91]

# Lista de empresas conforme fornecido
empresas = [
    "BOA VISTA", "BRASILIA", "CAMPOS", "FLORIANO", "MANAUS", "NOSSA SENHORA", "TERESINA", "SÃO LUIS", 
    "SÃO GONÇALO", "TANGUA", "LUZIANA", "PALMARES", "PETROPOLIS", "URUÇUI", "NACIONAL"
]

# Criar listas para armazenar as notas
notas = []

# Iterar sobre os índices desejados e capturar as notas
for i in range(len(indices)):
    index = indices[i]
    if index + 4 < len(linhas):  # Garantir que a linha de nota exista
        nota = linhas[index + 4].strip()  # Nota na 5ª linha

        # Verificar se a nota é "Não há dados neste relatório" e substituir por 0
        if "Não há dados neste relatório" in nota:
            nota = '0'

        notas.append(nota)

# Criar um DataFrame com as notas e as empresas
data_atualizacao = date.today().strftime('%Y-%m-%d')  # Data atual
df = pd.DataFrame({
    'Empresa': empresas[:len(notas)],  # Ajustar o tamanho da lista de empresas para coincidir com as notas
    'Nota Filial NPS Anual': notas,
    'data_atualizacao': data_atualizacao
})

# Exibir o DataFrame
print(df)
