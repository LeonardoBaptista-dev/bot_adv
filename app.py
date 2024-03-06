import time
import os
from datetime import datetime
import openpyxl
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from time import sleep
from auth import autenticar  # Importe a função de autenticação
from googleapiclient.discovery import build  # Importe a biblioteca do Google
from email.mime.text import MIMEText
import base64
import smtplib
from email.message import EmailMessage

# Abrir navegador
def entrar_chrome():
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico)
    driver.get("https://pje-consulta-publica.tjmg.jus.br")
    sleep(20)

    return (driver)

# Autenticar smtp gmail
def autenticar_gmail():
    with open('autenticacao.txt') as f:
        auth = f.readlines()

        f.close()

    senha_do_email = auth[0]
    email = auth[1]
    return (email, senha_do_email)

def extrair_dados(driver, senha_do_email, email):
    # Solicitar a OAB ao usuário
    # oab = input("Digite o numero da OAB do Adv que deseja fazer a automação: ")
    oab = 133864

    # Solicitar o estado ao usuário
    # estado = input("Digite o estado (exemplo 'SP'): ").strip().upper()
    estado = "SP" 

    # Verificar se o estado é válido
    estados_validos = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']
    if estado not in estados_validos:
        print("Estado inválido. Digite um estado válido.")
        exit()  # Encerra o programa

    # Digitar e-mail do advogado
    email_advogado = input("Digite o e-mail do Advogado: ").strip()

    # Digitar número da OAB e selecionar estado
    campo_oab = driver.find_element(By.XPATH, "//*[@id='fPP:Decoration:numeroOAB']")
    campo_oab.send_keys(oab)
    dropdown_estados = driver.find_element(By.XPATH, '//*[@id="fPP:Decoration:estadoComboOAB"]')
    opcoes_estados = Select(dropdown_estados)
    opcoes_estados.select_by_visible_text(f'{estado}')

    # Clicar em pesquisar
    pesquisar = driver.find_element(By.XPATH, '//input[@id="fPP:searchProcessos"]')
    pesquisar.click()
    sleep(20)

    # Criar uma lista de dicionários para armazenar os dados
    dados_processos = []

    # Entrar em cada um dos processos
    processos = driver.find_elements(By.XPATH, "//b[@class='btn-block']")
    for processo in processos:
        sleep(4)
        processo.click()
        sleep(20)
        janelas = driver.window_handles
        driver.switch_to.window(janelas[-1])
        driver.set_window_size(1280, 768)

        # Extrair os dados do processo
        nome_cliente = driver.find_elements(By.XPATH, "//td[@class='rich-table-cell ']")
        nome_cliente = nome_cliente[2].text

        nome_advogado = driver.find_elements(By.XPATH, "//div[@class='col-sm-12 ']")
        nome_advogado = nome_advogado[7].text

        numero_processo = driver.find_elements(By.XPATH, "//div[@class='col-sm-12 ']")
        numero_processo = numero_processo[0].text

        data_distribuicao = driver.find_elements(By.XPATH, "//div[@class='value col-sm-12 ']")
        data_distribuicao = data_distribuicao[1].text

        polo_passivo = driver.find_elements(By.XPATH, "//tr[@class='rich-table-row rich-table-firstrow ']")
        polo_passivo = polo_passivo[1].text

        classe_judicial = driver.find_elements(By.XPATH, "//div[@class='value col-sm-12 ']")
        classe_judicial = classe_judicial[2].text

        assunto = driver.find_elements(By.XPATH, "//div[@class='value col-sm-12 ']")
        assunto = assunto[3].text

        jurisdicao = driver.find_elements(By.XPATH, "//div[@class='value col-sm-12 ']")
        jurisdicao = jurisdicao[4].text

        # Extrair movimentações
        movimentacoes = driver.find_elements(By.XPATH, "//div[@id='j_id132:processoEventoPanel_body']//tr[contains(@class,'rich-table-row')]//td//div//div//span")
        lista_movimentacoes = [movimentacao.text for movimentacao in movimentacoes]

        # Adicionar dados à lista
        dados_processo = {
            "Nome Cliente": nome_cliente,
            "Nome Advogado": nome_advogado,  
            "E-mail Cliente": None,  # Inicialmente definido como None
            "E-mail Advogado": email_advogado,
            "Número Processo": numero_processo,
            "Data Distribuição": data_distribuicao,
            "Polo Passivo": polo_passivo,
            "Classe Judicial": classe_judicial,
            "Assunto": assunto,
            "Jurisdição": jurisdicao,
            "Movimentações": "\n".join(lista_movimentacoes)
            
        }
        dados_processos.append(dados_processo)
        print(f"Processo {numero_processo} extraído com sucesso")
        driver.close()
        sleep(5)
        driver.switch_to.window(driver.window_handles[0])

    # Salvar os dados em um arquivo Excel
    df = pd.DataFrame(dados_processos)

    # Verificar se a pasta 'planilhas' existe e, se não, criá-la
    pasta_planilhas = 'planilhas'
    if not os.path.exists(pasta_planilhas):
        os.makedirs(pasta_planilhas)

    # Carregar o último arquivo xlsx salvo na pasta 'planilhas'
    lista_arquivos = os.listdir(pasta_planilhas)
    lista_arquivos.sort(reverse=True)

    # Carregar os dados do último arquivo, se existir
    df_anterior = pd.DataFrame()
    if len(lista_arquivos) > 0:
        arquivo_anterior = os.path.join(pasta_planilhas, lista_arquivos[0])
        df_anterior = pd.read_excel(arquivo_anterior)


    # Coletar informações de e-mail e nome do cliente para cada processo
    for index, row in df.iterrows():
        if pd.isna(row['E-mail Cliente']) or pd.isna(row['Nome Cliente']):
            # Verificar se o e-mail já está cadastrado no arquivo anterior
            numero_processo_atual = row['Número Processo']
            if not df_anterior.empty and 'E-mail Cliente' in df_anterior.columns:
                if numero_processo_atual in df_anterior['Número Processo'].values:
                    email_cadastrado = df_anterior.loc[df_anterior['Número Processo'] == numero_processo_atual, 'E-mail Cliente'].values[0]
                    if not pd.isna(email_cadastrado):
                        row['E-mail Cliente'] = email_cadastrado
                        continue  # Pula para o próximo processo

            # Se não estiver cadastrado, solicita ao usuário inserir o e-mail
            email_cliente = input(f"Digite o e-mail do cliente para o processo {row['Nome Cliente']}: ")
            df.at[index, 'E-mail Cliente'] = email_cliente


    # Obter a data atual formatada 
    data_atual = datetime.now().strftime("%d-%m-%Y")


    # Salvar os dados em um arquivo Excel com o nome personalizado
    nome_arquivo = f'coleta_dados_{data_atual}.xlsx'
    df.to_excel(os.path.join(pasta_planilhas, nome_arquivo), index=False)

    # Ordena para obter o último arquivo gerado
    pasta_planilhas = 'planilhas'
    lista_arquivos = os.listdir(pasta_planilhas)
    lista_arquivos.sort(reverse=True)  
    if len(lista_arquivos) > 1:
        arquivo_anterior = os.path.join(pasta_planilhas, lista_arquivos[1])
        df_anterior = pd.read_excel(arquivo_anterior)
    else:
        df_anterior = pd.DataFrame()
    
    return (df, pasta_planilhas, df_anterior)


def enviar_emails(df_anterior, email, senha_do_email, df):
    for index, row in df.iterrows():
        movimentacoes = row['Movimentações']

        if not pd.isna(row['E-mail Cliente']):
            if not df_anterior.empty and row['Número Processo'] in df_anterior['Número Processo'].values:
                movimentacoes_anterior = df_anterior.loc[df_anterior['Número Processo'] == row['Número Processo'], 'Movimentações'].values[0]
                
                if movimentacoes != movimentacoes_anterior:
                    msg = EmailMessage()
                    msg['Subject'] = f"Relatório diário de processos - {row['Número Processo']}"
                    msg['From'] = email
                    msg['To'] = row['E-mail Cliente']

                    if movimentacoes:
                        msg.set_content(f"Saudações, {row['Nome Cliente']}!\n\nForam registradas novas movimentações no processo {row['Número Processo']}:\n{movimentacoes}")
                    else:
                        msg.set_content(f"Saudações, {row['Nome Cliente']}!\n\nNão houve movimentações recentes no processo {row['Número Processo']}.")

                    '''with open(f'planilhas/{nome_arquivo}', 'rb') as content_file:
                        content = content_file.read()
                        msg.add_attachment(content, maintype='application', subtype='xlsx', filename=nome_arquivo)'''

                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email, senha_do_email)
                        smtp.send_message(msg)
                        print(f"E-mail enviado para {row['Nome Cliente']} ({row['E-mail Cliente']}) para o processo {row['Número Processo']}")

                    

                else:
                    # Se não houve alterações, envia e-mail informando sobre a falta de modificações
                    msg = EmailMessage()
                    msg['Subject'] = f"Relatório diário de processos - {row['Número Processo']}"
                    msg['From'] = email
                    msg['To'] = row['E-mail Cliente']
                    msg.set_content(f"Saudações, {row['Nome Cliente']}!\n\nNão houve movimentações recentes no processo {row['Número Processo']}.")

                    '''with open(f'planilhas/{nome_arquivo}', 'rb') as content_file:
                        content = content_file.read()
                        msg.add_attachment(content, maintype='application', subtype='xlsx', filename=nome_arquivo)'''

                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email, senha_do_email)
                        smtp.send_message(msg)
                        print(f"E-mail enviado para {row['Nome Cliente']} ({row['E-mail Cliente']}) informando sobre a ausência de movimentações no processo {row['Número Processo']}")

def enviar_email_advogado(email_advogado, nome_arquivo, email, senha_do_email):
    msg = EmailMessage()
    msg['Subject'] = "Relatório diário de processos para Advogado"
    msg['From'] = email  # Email remetente
    msg['To'] = email_advogado  # Email do advogado

    msg.set_content("Olá, segue o relatório diário de processos em anexo.")
    with open(f'planilhas/{nome_arquivo}', 'rb') as content_file:
        content = content_file.read()
        msg.add_attachment(content, maintype='application', subtype='xlsx', filename=nome_arquivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, senha_do_email)
        smtp.send_message(msg)
        print(f"E-mail enviado para o Advogado ({email_advogado}) com o relatório diário de processos.") 

    



entrar_chrome()

autenticar_gmail()

extrair_dados(driver, senha_do_email, email)

enviar_emails(df_anterior, email, senha_email, df)

enviar_email_advogado(email_advogado, nome_arquivo)

input("Enter para fechar")
