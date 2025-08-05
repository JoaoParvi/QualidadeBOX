
import time
import pandas as pd
import urllib
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from sqlalchemy import create_engine
from datetime import date

# Configurar o serviço do WebDriver manualmente
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Definir a URL e as credenciais
url = 'https://jlr.medallia.eu/sso/jlr/applications/ex_WEB-9/pages/280?roleId=557&f.timeperiod=363&alreftoken=%22dcd5378b593c41ce076edf5ee7e4cf8f%22'
login = "t-soare2"
senha = "Tatiana32"

def send_multiple_keys(navegador, key, times):
    for _ in range(times):
        navegador.switch_to.active_element.send_keys(key)
        time.sleep(1)

navegador.maximize_window()        
        
    # Navegar ate a URL
navegador.get(url)


# clicando no Retalier
cliqueRetailer= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#bySelection > div:nth-child(4) > div > span")))
cliqueRetailer.click()
time.sleep(7)

# Esperar o campo de login aparecer e preench�-lo
campo_login = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#user")))
campo_login.click()
campo_login.send_keys(login)
time.sleep(2)


# Esperar o campo de senha aparecer e preench�-lo
campo_senha = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#password")))
campo_senha.click()
campo_senha.send_keys(senha)
time.sleep(2)

# clicando no login
cliquelogin= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#signon")))
cliquelogin.click()
time.sleep(2)

time.sleep(10)


cliquefiltro= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#main-content-for-skip-link > div.sc-1s2t1lx-0.gGBlLA > div.sc-1s2t1lx-2.hJvoIP > div > div > div > div > div > div > div > div.sc-j2f5w4-3.fhVsgU.aui--remove-on-print > button > div > div.sc-rzscvc-4.sc-rzscvc-5.fSCtek.gXnbIR")))
cliquefiltro.click()
time.sleep(6)


cliqueWayJoaoPessoa= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#saved-filters-popper-content > div:nth-child(3) > ul > li:nth-child(1) > div.sc-1uun7m9-4.bdotrt > button")))
cliqueWayJoaoPessoa.click()
time.sleep(6)

send_multiple_keys(navegador, Keys.PAGE_DOWN, 4)

NotaFilialJP = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_JP = NotaFilialJP.text  
time.sleep(6)

NotaFilialJPJaguar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_JP_Jaguar = NotaFilialJPJaguar.text  # Extrair o texto do elemento
time.sleep(6)

NotaNacional = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/div/div/div[2]/label/span[2]')))
nota_Nacionall = NotaNacional.text  # Extrair o texto do elemento
time.sleep(6)

cliquefiltro= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#main-content-for-skip-link > div.sc-1s2t1lx-0.gGBlLA > div.sc-1s2t1lx-2.hJvoIP > div > div > div > div > div > div > div > div.sc-j2f5w4-3.fhVsgU.aui--remove-on-print > button > div > div.sc-rzscvc-4.sc-rzscvc-5.fSCtek.gXnbIR")))
cliquefiltro.click()
time.sleep(6)

 # clicando no login
cliqueWayManaus= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#saved-filters-popper-content > div:nth-child(3) > ul > li:nth-child(2) > div.sc-1uun7m9-4.bdotrt > button")))
cliqueWayManaus.click()
time.sleep(6)

send_multiple_keys(navegador, Keys.PAGE_DOWN, 6)

NotaFilialManaus = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_Manaus = NotaFilialManaus.text  # Extrair o texto do elemento
time.sleep(6)

cliquefiltro= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#main-content-for-skip-link > div.sc-1s2t1lx-0.gGBlLA > div.sc-1s2t1lx-2.hJvoIP > div > div > div > div > div > div > div > div.sc-j2f5w4-3.fhVsgU.aui--remove-on-print > button > div > div.sc-rzscvc-4.sc-rzscvc-5.fSCtek.gXnbIR")))
cliquefiltro.click()
time.sleep(6)


cliqueWayRecife= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#saved-filters-popper-content > div:nth-child(3) > ul > li:nth-child(3) > div.sc-1uun7m9-4.bdotrt > button")))
cliqueWayRecife.click()
time.sleep(2)

send_multiple_keys(navegador, Keys.PAGE_DOWN, 6)


NotaFilialRecife = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_Recife = NotaFilialRecife.text  # Extrair o texto do elemento
time.sleep(2)


cliquefiltro= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#main-content-for-skip-link > div.sc-1s2t1lx-0.gGBlLA > div.sc-1s2t1lx-2.hJvoIP > div > div > div > div > div > div > div > div.sc-j2f5w4-3.fhVsgU.aui--remove-on-print > button > div > div.sc-rzscvc-4.sc-rzscvc-5.fSCtek.gXnbIR")))
cliquefiltro.click()
time.sleep(6)


cliqueWaySalvador= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#saved-filters-popper-content > div:nth-child(3) > ul > li:nth-child(5) > div.sc-1uun7m9-4.bdotrt > button")))
cliqueWaySalvador.click()
time.sleep(2)

send_multiple_keys(navegador, Keys.PAGE_DOWN, 6)


NotaFilialSalvador = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_Salvador = NotaFilialSalvador.text  # Extrair o texto do elemento
time.sleep(2)


NotaFilialSalvadorJaguar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_Salvador_Jaguar = NotaFilialSalvadorJaguar.text  # Extrair o texto do elemento
time.sleep(2)


cliquefiltro= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#main-content-for-skip-link > div.sc-1s2t1lx-0.gGBlLA > div.sc-1s2t1lx-2.hJvoIP > div > div > div > div > div > div > div > div.sc-j2f5w4-3.fhVsgU.aui--remove-on-print > button > div > div.sc-rzscvc-4.sc-rzscvc-5.fSCtek.gXnbIR")))
cliquefiltro.click()
time.sleep(6)


cliqueWaySaoLuis= WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#saved-filters-popper-content > div:nth-child(3) > ul > li:nth-child(5) > div.sc-1uun7m9-4.bdotrt > button")))
cliqueWaySaoLuis.click()
time.sleep(2)

send_multiple_keys(navegador, Keys.PAGE_DOWN, 6)



NotaFilialSaoLuis = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_SaoLuis = NotaFilialSaoLuis.text  # Extrair o texto do elemento
time.sleep(2)


NotaFilialSaoLuisJaguar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[2]/div/div[9]/div/div[2]/div/div[2]/div[2]/div/section/div/div/div/div[1]/div/div[2]')))
nota_filial_SaoLuis_Jaguar = NotaFilialSaoLuisJaguar.text  # Extrair o texto do elemento
time.sleep(2)

print(nota_Nacionall)
print(nota_filial_JP)
print(nota_filial_JP_Jaguar)
print(nota_filial_Manaus)
print(nota_filial_Recife)
print(nota_filial_Salvador)
print(nota_filial_Salvador_Jaguar)
print(nota_filial_SaoLuis)
print(nota_filial_SaoLuis_Jaguar)



nota_Nacional1 = nota_Nacionall
nota_filial_JP1 = nota_filial_JP
nota_filial_JP_Jaguar1 = nota_filial_JP_Jaguar
nota_filial_Manaus1 = nota_filial_Manaus
nota_filial_Recife1 = nota_filial_Recife
nota_filial_Salvador1 = nota_filial_Salvador
nota_filial_Salvador_Jaguar1 = nota_filial_Salvador_Jaguar
nota_filial_SaoLuis1 = nota_filial_SaoLuis
nota_filial_SaoLuis_Jaguar1 = nota_filial_SaoLuis_Jaguar



# Cria uma lista com os valores das variáveis
notas = [
    nota_filial_JP1,
    nota_filial_JP_Jaguar1,
    nota_filial_Manaus1,
    nota_filial_Recife1,
    nota_filial_Salvador1,
    nota_filial_Salvador_Jaguar1,
    nota_filial_SaoLuis1,
    nota_filial_SaoLuis_Jaguar1
]

nota_Nacional = [
    nota_Nacional1
]

df1 = pd.DataFrame(notas, columns=["Nota"])
df2 = pd.DataFrame(nota_Nacional, columns=["Nota_Nacional"])
df3 = pd.DataFrame({"Segmento": ["Pos Vendas"]})

df4 = pd.concat([df1.reset_index(drop=True), df2.reset_index(drop=True)], axis=1)

df = pd.concat([df4.reset_index(drop=True), df3.reset_index(drop=True)], axis=1)
df["Segmento"] = df["Segmento"].fillna("Pos Vendas")

Empresa = {
    'nota_filial_JP1': ['LandJoaoPessoa'],
    'nota_filial_JP_Jaguar1': ['JaguarJoaoPessoa'],
    'nota_filial_Manaus1': ['LandManaus'],
    'nota_filial_Recife1': ['LandRecife'],
    'nota_filial_Salvador1': ['LandSalvador'],
    'nota_filial_Salvador_Jaguar1': ['JaguarManaus'],
    'nota_filial_SaoLuis1': ['LandSaoLuis'],
    'nota_filial_SaoLuis_Jaguar1': ['JaguarSaoLuis']
}
    
df5 = pd.DataFrame(Empresa)
df6 = df5.melt(var_name='filial', value_name='Empresa')

df7 = df6[['Empresa']]

df = pd.concat([df.reset_index(drop=True), df7.reset_index(drop=True)], axis=1)

# Adicionando a coluna "data_atualizacao" com a data atual
df['data_atualizacao'] = date.today()

df['Nota'] = df['Nota'].replace(to_replace=r'[^0-9,\.]', value='0', regex=True)

# Substituir vírgula por ponto (para conversão correta)
df['Nota'] = df['Nota'].str.replace(',', '.', regex=False)

# Converter para float
df['Nota'] = df['Nota'].astype(float)

# Preencher os NaNs da coluna Nota_Nacional com o valor presente (ex: 100)
valor_nacional = df['Nota_Nacional'].dropna().iloc[0]
df['Nota_Nacional'] = df['Nota_Nacional'].fillna(valor_nacional)

print("Conectando ao banco de dados...")
user = 'rpa_bi'
password = 'Rp@_B&_P@rvi'
host = '10.0.10.243'
port = '54949'
database = 'stage'

params = urllib.parse.quote_plus(
    f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={host},{port};DATABASE={database};UID={user};PWD={password}')
connection_str = f'mssql+pyodbc:///?odbc_connect={params}'
engine = create_engine(connection_str)
table_name = "QualidadeBox_PosVendas"

with engine.connect() as connection:
    df.to_sql(table_name, con=connection, if_exists='replace', index=False)

print(f"Dados inseridos com sucesso na tabela '{table_name}'!")
print("Fechando o navegador...")
navegador.quit()
