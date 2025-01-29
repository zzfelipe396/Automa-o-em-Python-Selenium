# Importando as Bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas as pd
import openpyxl
import os

# Lendo o DF
df = pd.read_excel(r'C:\Users\felip\OneDrive\Desktop\fretes.xlsx')

# Criando listas para armazenar as informações coletadas
prazos = []
precos = []

# Configurando o navegador
service = Service(ChromeDriverManager().install())
options = Options()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()

driver.get('https://www2.correios.com.br/sistemas/precosPrazos/')
print('- Abriu o navegador')
sleep(3)

# Criando um loop para coletar as informações
for index, row in df.iterrows():
    origem = driver.find_element(By.NAME, 'cepOrigem').send_keys(str(row['origem']))
    print('- Informou o CEP de origem')
    sleep(2)

    destino = driver.find_element(By.NAME, 'cepDestino').send_keys(str(row['destino']))
    print('- Informou o CEP de destino')
    sleep(2)

    Select(driver.find_element(By.NAME, 'servico')).select_by_index(15)
    print('- Selecionou o tipo de serviço')
    sleep(2)

    Select(driver.find_element(By.NAME, 'embalagem1')).select_by_index(2)
    print('- Selecionou a embalagem')
    sleep(2)

    driver.find_element(By.NAME, 'Altura').send_keys(int(row['altura']))
    print('- Informou a altura')
    sleep(2)

    driver.find_element(By.NAME, 'Largura').send_keys(int(row['largura']))
    print('- Informou a largura')
    sleep(2)

    driver.find_element(By.NAME, 'Comprimento').send_keys(int(row['comprimento']))
    print('- Informou o comprimento')
    sleep(2)

    driver.find_element(By.NAME, 'peso').send_keys(int(row['peso']))
    print('- Informou o peso')
    sleep(2)

    driver.find_element(By.NAME, 'ckValorDeclarado').click()
    print('- Marcou os serviços opcionais')
    sleep(2)

    driver.find_element(By.NAME, 'valorDeclarado').send_keys(int(row['valor']))
    print('- Informou o valor declarado')
    sleep(2)

    driver.find_element(By.NAME, 'Calcular').click()
    print('- Calculando o valor do frete')
    sleep(2)

    # Direcionando para a janela de Preços/Prazos
    driver.switch_to.window(driver.window_handles[1])
    sleep(2)

    tempo_entrega = driver.find_element(By.CLASS_NAME, 'destaque').find_element(By.TAG_NAME, 'td').text
    print('- Pegou o tempo de entrega')
    sleep(1)

    preco_entrega = driver.find_elements(By.CLASS_NAME, 'destaque')[1].find_element(By.TAG_NAME, 'td').text
    print('- Pegou o preço de entrega')
    sleep(1)

    prazos.append(tempo_entrega)
    precos.append(preco_entrega)
    sleep(3)

    driver.close() # Fechando a janela de Preços/Prazos
    driver.switch_to.window(driver.window_handles[0]) # Voltando para a janela de envio
    driver.refresh() # Atualizando a página

data = {'origem': df.iloc[:,0], 'destino': df.iloc[:,1], 'altura': df.iloc[:,2], 'largura': df.iloc[:,3], 'comprimento': df.iloc[:,4], 'peso': df.iloc[:,5]
        , 'valor': df.iloc[:,6], 'prazos': prazos, 'precos': precos}
dfTwo = pd.DataFrame(data)

dfTwo.to_excel('base_tratada.xlsx', index=False)
print(dfTwo)
