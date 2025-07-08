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

def executar_script():
    try:
        # Configurar o serviço do WebDriver manualmente
        options = Options()
        options.binary_location = r"CC:\Program Files\Google\Chrome\Application\chrome.exe"  # caminho do seu Chrome
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=options)

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
        Cliqueparafecharfiltro = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#pdf-dashboard > div > div.session__side-bar > button > i")))
        Cliqueparafecharfiltro.click()
        time.sleep(6)

        print("Extraindo texto...")
        Neutro_BoaVista = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#panel-main")))
        texto_extraido = Neutro_BoaVista.text
        time.sleep(2)

        print("Processando dados...")
        linhas = texto_extraido.split('\n')
        indices = [1, 7, 13, 19, 26, 32, 38, 44, 50, 56, 63, 70, 77, 84, 91]
        empresas = [
            "BOA VISTA", "BRASILIA", "CAMPOS", "FLORIANO", "MANAUS", "NOSSA SENHORA",
            "TERESINA", "SÃO LUIS", "SÃO GONÇALO", "TANGUA", "LUZIANA", "PALMARES",
            "PETROPOLIS", "URUÇUI", "NACIONAL"
        ]

        notas = []
        for i in range(len(indices)):
            index = indices[i]
            if index + 4 < len(linhas):
                nota = linhas[index + 4].strip()
                if "Não há dados neste relatório" in nota:
                    nota = '0'
                notas.append(nota)

        data_atualizacao = date.today().strftime('%Y-%m-%d')
        df = pd.DataFrame({
            'Empresa': empresas[:len(notas)],
            'Nota Filial NPS Anual': notas,
            'data_atualizacao': data_atualizacao
        })

        print("Salvando dados no Excel...")
        df.to_excel('dados_nps.xlsx', index=False)

        print("Conectando ao banco de dados...")
        user = 'rpa_bi'
        password = 'Rp@_B&_P@rvi'
        host = '10.0.10.243'
        port = '54949'
        database = 'stage'

        params = urllib.parse.quote_plus(
            f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={host},{port};DATABASE={database};UID={user};PWD={password}'
        )
        connection_str = f'mssql+pyodbc:///?odbc_connect={params}'
        engine = create_engine(connection_str)
        table_name = "NPS_Vendas_Anual_IndeCX"

        with engine.connect() as connection:
            df.to_sql(table_name, con=connection, if_exists='replace', index=False)

        print(f"Dados inseridos com sucesso na tabela '{table_name}'!")

    except Exception as e:
        print(f" Erro durante a execução: {e}")

    finally:
        if 'navegador' in locals():
            print("Fechando o navegador...")
            navegador.quit()

executar_script()
