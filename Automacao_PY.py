#----------Iniciando BOT----------#
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select as Sel
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
import os
from os import listdir
from os.path import isfile, join
import os.path
import shutil
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from datetime import date
import win32com.client as win32
import time
import schedule
from ftplib import FTP

def job():

# criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

# criar um email
    email = outlook.CreateItem(0)

# configurar as informações do seu e-mail
    email.To = ""
    email.Subject = "E-mail automático BOT"
    email.HTMLBody = f"""
    <p>BOT inicializado ás: 05:00</p>

    <p>Realizando a baixa do arquivo CNAB Afinitty Itaú.</p>
    <p>Inicializado processos com sucesso!!</p>

    <p>Atenciosamente IA,</p>
    """

    # anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
    # email.Attachments.Add(anexo)

    email.Send()

    print("Email Enviado")

    data_e_hora_atuais = datetime.now() + timedelta(days = -1)
    data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y')
    opcao_input = "código do operador"
    codigo_input = "xxxxx"
    codigo_secreto = "xxxxx"   

#Atribuir pasta rede para download
    print("Atibuindo pasta rede para download do arquivo...")
    pasta = "xxx"
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {"download.default_directory": pasta, "download.prompt_for_download": False, "download.directory_upgrade": True,"plugins.always_open_pdf_externally": True})

# instala a versão atualizada do ChromeDriver
    print("Instalando a versão recente do ChromeDriver...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

#Abrir Chrome
    print("Abrindo navegador Chrome...")
    driver.maximize_window()
#Direcionar para o site QPROF
    print("Acessando Site do Itaú...")
    driver.get("https://www.itau.com.br/empresas")

#Clicar no botão +acessos
    print("Baixando Arquivo CNAB...")
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/header/div[2]/div/div[3]/form/button[3]"))).click()

# espera até o campo de código ser localizado na página
    drop_codigo = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/header/div[2]/div/div[4]/div/div/div/div[2]/form/div/select")))

# clica no item selecionado
    time.sleep(1)
    Sel(drop_codigo).select_by_visible_text(opcao_input)

#atribui a variável codigo_operador o campo código
    codigo_operador = driver.find_element(by=By.XPATH, value="/html/body/div[1]/header/div[2]/div/div[4]/div/div/div/div[5]/input")

#escreve no campo de Usuário
    codigo_operador.send_keys(codigo_input)
    time.sleep(2)
#Clica no botão acessar
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/header/div[2]/div/div[4]/div/div/div/div[7]/button"))).click()
    time.sleep(20)
#Mapear cada número do código, se constar no texto do botão então clica
    for letra in codigo_secreto:
            for user_num in range(1,6):
                botao_password = WebDriverWait(driver, 65).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/section/div/section/div[1]/div[2]/div/div[1]/div[1]/form/fieldset/div[2]/div[1]/a[" + str(user_num) + "]")))
                if letra in botao_password.text:
                    time.sleep(1)
                    botao_password.click()
            #print(letra + " | " + botao_password.text)

    time.sleep(3)
#Clica no botão acessar
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/div/section/div[1]/div[2]/div/div[1]/div[1]/form/fieldset/div[2]/div[2]/a"))).click()
    time.sleep(1)
#Clicar na flag acesso básico
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/section[2]/div/section/div/div/form/div/section/fieldset/ul/li[2]/p[1]/input"))).click()

#Clicar no botão continuar
    time.sleep(1)
    driver.execute_script("javascript:continuar()")
    time.sleep(5)
#Habilitar barra pesquisa
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/header/div[5]/div[1]/div[2]/nav/ul/li[1]/a"))).click()
#Atribuir barra pesquisa na variaveil buscar_pagina
    buscar_pagina =driver.find_element(by=By.XPATH, value="/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/fieldset/input[1]")
#Escrever o caminho para pesquisar
    buscar_pagina.send_keys("Recepcionar arquivos de retorno")
    time.sleep(3)
# atribui a variável html_list a unordered list que tem os fundos ( lista do drop down )
    #html_list = driver.find_element(by=By.XPATH, value="/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/div[2]/div/ul[3]/li[1]/a[2]")
# atribui a variável items os elementos com a tag "li" ( list item )
    #items = html_list.find_elements(by=By.TAG_NAME, value="li")
#atribui zero a variável elementos_lista
    #elementos_lista = 0

#for item in items:
    # incrementa a variável elementos_lista
    #elementos_lista = elementos_lista + 1
    # procurar o item na lista do drop down
    #if (driver.find_element(by=By.XPATH, value="/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/div[2]/div/ul[3]/li[1]/a[2]"+str(elementos_lista)+"]").text).find("Recepcionar arquivos de retorno") != -1:
        # se achar validar o tamanha do item, clicar no menor item (sem o texto ambiente teste)
        #if len(driver.find_element(by=By.XPATH, value="/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/div[2]/div/ul[3]/li[1]/a[2]"+str(elementos_lista)+"]").text) == 102:
            # caso o fundo não seja Afinitty, clica no dropdown
           # WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/div[2]/div/ul[3]/li[1]/a[2]"+str(elementos_lista)+"]"))).click()
   
#Clicar no primeiro item pesquisado
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/header/div[5]/div[1]/div[2]/div/div/div[2]/div/ul[3]/li[1]/a[2]/div"))).click()

#Selecionar o frame com a lista de data/arquivos
    time.sleep(2)
    driver.switch_to.frame("output_frame_recepcao")
#Expandir tabela com lista
    driver.find_element(by=By.XPATH, value="//html/body/form[1]/div[3]/div/div/div/div[1]/table/tbody/tr").click()

#Loop para mapear as datas e baixar o arquivo D1
    for lin_dt in range(2,7):
        datas = driver.find_element(by=By.XPATH, value="/html/body/form[1]/div[3]/div/div/div/div[1]/div/table[1]/tbody/tr[" + str(lin_dt) + "]/td[1]").text
        if datas == data_e_hora_em_texto:
            print('Data do arquivo baixado: ' + str(datas))
            driver.find_element(by=By.XPATH, value="/html/body/form[1]/div[3]/div/div/div/div[1]/div/table[1]/tbody/tr[" + str(lin_dt) + "]/td[5]").click()
            time.sleep(6)
            print("Arquivo Baixado com Sucesso!!")

                        # Caminho da pasta
    cpast = "xxxx"

# Obtém a lista de arquivos na pasta
    arquivos = os.listdir(cpast)

# Escolha o primeiro arquivo da lista (ou o arquivo desejado)
    nome_arquivo = arquivos[0]

# Imprime o nome do arquivo
    time.sleep(1)
    print('Arquivo localizado: ',nome_arquivo)

# Configurações do servidor FTP
    time.sleep(1)
    print('Conectando ao servidor de FTP...')
    ftp_host = 'xxxx'
    ftp_user = 'xxxx'
    ftp_password = 'xxx'

# Caminho local do arquivo a ser enviado
    local_file = 'xxxxx' + nome_arquivo

# Caminho remoto do diretório onde o arquivo será salvo
    remote_folder = 'xxxx'

# Nome do arquivo remoto (como será salvo no servidor FTP)
    remote_filename = 'xxxx'

# Conectar ao servidor FTP
    ftp = FTP(ftp_host)
    ftp.login(user=ftp_user, passwd=ftp_password)

# Navegar para o diretório remoto
    ftp.cwd(remote_folder)

# Abrir o arquivo local em modo de leitura binária
    with open(local_file, 'rb') as file:
    # Enviar o arquivo para o servidor FTP
        ftp.storbinary('STOR ' + remote_filename, file)
        time.sleep(1)
        print('Transferindo arquivo para FTP...')
# Fechar a conexão FTP
        ftp.quit()
        print('Arquivo transferido com sucesso!!')
#driver.quit
#pausar = input("Digite para finalizar: ")

horario_execucao = "11:12"

schedule.every().day.at(horario_execucao).do(job)

while True:
    schedule.run_pending()
    time.sleep(1)