import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
import numpy as np
from time import sleep
import datetime
import os
from openpyxl import Workbook
import math


def reiniciaSite():
    os.system('cls')
    sleep(1)
    navegador.get('<<URL>>')
    sleep(2.5)

def capturaDeLinhas():
    df = pd.read_excel('database.xlsx', sheet_name='db', usecols='A')
    return len(df.index)

def testaNew():
    if not os.path.exists('databaseNew.xlsx'):
        workbook = Workbook()
        workbook.create_sheet(title="db")
        workbook.save(filename='databaseNew.xlsx')

def saveDatabase(line):
    global database
    with open('ultimaLinha.txt', "w") as f:
        f.write("A última linha visitada foi a linha {} !".format(lastLine - 1))

    with pd.ExcelWriter("databaseNew.xlsx", mode='a', if_sheet_exists="overlay") as writer:
        database.to_excel(writer, sheet_name='db', startrow=line, startcol=0, header=False, index=False)

def getData(linhas):
    try:
        planilha = pd.read_excel('database.xlsx', sheet_name='db', dtype='object', nrows=299, skiprows=linhas-1,
                                 header=None)
    except Exception as e:
        print(f"[*] Falha na leitura da base de dados.")
        print(f"[-] Erro: {e}")
        exit()

    planilha = tratamentoDados(planilha)

    return planilha


def tratamentoDados(planilha):
    coluna_doc = planilha[0]
    tamanhos_linhas = coluna_doc.str.len()

    indices_inválidos = [i for i in range(len(tamanhos_linhas)) if tamanhos_linhas[i] != 14]

    planilha = planilha.drop(indices_inválidos)

    return planilha


def setList(tel):
    global database
    base = database.copy()
    base.drop(base.loc[0:(- 1)].index, inplace=True)
    dados = []
    for i in base.itertuples():
        dados.append(i[0:(3 + tel)])
    return dados


def setNumbers(qtd, keys):
    numeros = np.array(keys[3:(3 + qtd)])
    return numeros


def automation():
    global lastLine
    global keys
    global qtdtelefones
    global linha
    contador = linha

    try:
        reiniciaSite()
        for k in keys:
            numeros = setNumbers(qtdtelefones, k)
            lastLine = contador

            for numb in numeros:
                if str(numb) != 'nan':
                    answare = typeKeys(k[1], k[2], numb)

                    if (verifyAnsware(answare, contador, k[0]) == '0'):
                        clicking(
                            '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div[2]/div/button')  # Return form
                        break

                    clicking(
                        '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div[2]/div/button')  # Return form
                else:
                    continue
            contador += 1
    except KeyboardInterrupt:
        print("[*] Interrupcao Manual, salvando a Base de dados")
        saveDatabase(linha)
        navegador.quit()
        exit()
    except Exception as e:
        print("[!] Erro na funcao principal, salvando a Base de Dados")
        print(f"[-] Erro: {e}")
        saveDatabase(linha)
        navegador.quit()
        exit()


def clicking(path):
    WebDriverWait(navegador, 240).until(EC.presence_of_element_located((By.XPATH, path)))
    WebDriverWait(navegador, 240).until(EC.element_to_be_clickable((By.XPATH, path))).click()


def typeKeys(doc, ddd, tel):
    ancora = WebDriverWait(navegador, 240).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div/input')))
    cnpj = navegador.find_element(By.XPATH,
                                  '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div/input')
    number = navegador.find_element(By.XPATH,
                                    '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div/input')

    if ancora.is_displayed() and cnpj.is_enabled() and number.is_enabled():
        cnpj.clear()
        cnpj.send_keys(str(doc))
        number.clear()
        number.send_keys(str(ddd))
        number.send_keys(str(tel))

        clicking(
            '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[1]/form/div[4]/div/button')  # Submit
        return siteAnsware()


def siteAnsware():
    sleep(0.5)
    answare = WebDriverWait(navegador, 240).until(EC.presence_of_element_located((By.XPATH,
                                                                                  '//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div[1]/table/tbody/tr/td[4]')))
    if answare.is_displayed() and answare.size is not None and answare.text != '' and answare.get_attribute(
            "textContent") != '':
        try:
            return navegador.find_element(By.XPATH,'//*[@id="wrapper"]/div[4]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div[1]/table/tbody/tr/td[4]').text
        except:
            print("[!] Erro ao capturar a resposta")
            return 'ERRO'


def verifyAnsware(answare, cont, row):
    if (answare == 'TERMINAL APTO PARA COBRANCA'):
        database.loc[row, 'OBS'] = 'APTO'
        database.loc[row, 'DATA'] = str(today)
        print(f"[+] APTO {cont}")
        return '0'
    else:
        database.loc[int(row), 'OBS'] = 'X'
        database.loc[int(row), 'DATA'] = str(today)
        print(f"[-] INAPTO {cont}")
        return '1'


# MAIN CODE

linha = 0
today = datetime.date.today()
totalLinhas = capturaDeLinhas()

print("""
███╗░░░███╗███████╗██████╗░░█████╗░░█████╗░░██████╗██╗░░░██╗██╗░░░░░
████╗░████║██╔════╝██╔══██╗██╔══██╗██╔══██╗██╔════╝██║░░░██║██║░░░░░
██╔████╔██║█████╗░░██████╔╝██║░░╚═╝██║░░██║╚█████╗░██║░░░██║██║░░░░░
██║╚██╔╝██║██╔══╝░░██╔══██╗██║░░██╗██║░░██║░╚═══██╗██║░░░██║██║░░░░░
██║░╚═╝░██║███████╗██║░░██║╚█████╔╝╚█████╔╝██████╔╝╚██████╔╝███████╗
╚═╝░░░░░╚═╝╚══════╝╚═╝░░╚═╝░╚════╝░░╚════╝░╚═════╝░░╚═════╝░╚══════╝

░█████╗░██╗░░░██╗████████╗░█████╗░███╗░░░███╗░█████╗░████████╗██╗░█████╗░███╗░░██╗
██╔══██╗██║░░░██║╚══██╔══╝██╔══██╗████╗░████║██╔══██╗╚══██╔══╝██║██╔══██╗████╗░██║
███████║██║░░░██║░░░██║░░░██║░░██║██╔████╔██║███████║░░░██║░░░██║██║░░██║██╔██╗██║
██╔══██║██║░░░██║░░░██║░░░██║░░██║██║╚██╔╝██║██╔══██║░░░██║░░░██║██║░░██║██║╚████║
██║░░██║╚██████╔╝░░░██║░░░╚█████╔╝██║░╚═╝░██║██║░░██║░░░██║░░░██║╚█████╔╝██║░╚███║
╚═╝░░╚═╝░╚═════╝░░░░╚═╝░░░░╚════╝░╚═╝░░░░░╚═╝╚═╝░░╚═╝░░░╚═╝░░░╚═╝░╚════╝░╚═╝░░╚══╝ V2.0
By: https://github.com/WhiteCJbr""")
print()
print(f"""Bem vindo a automacao Mercosul. Selecione a opcao desejada:
        [ 1 ] Comecar do inicio da tabela
        [ 2 ] Comecar de uma linha especifica""")

opcao = -1
while opcao < 0 or opcao > 2:
    opcao = int(input("Opcao: "))

if opcao == 2:
    while (linha <= 0):
        linha = int(input("Digite a linha de início: "))
        if (linha > totalLinhas):
            print(f"[!] ATENCAO: A base de dados possui somente {totalLinhas} linhas !\n")
            print("Digite novamente !")
            linha = 0

qtdtelefones = int(input(f"Digite a quantidade de telefones na tabela: "))

try:
    chromedriver_path = "chromedriver.exe"

    service = Service(chromedriver_path)
    service.start()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--disable-gpu')
    navegador = webdriver.Chrome(options=options, service=service)
except Exception as e:
    print("[!] ERRO AO DEFINIR O CHROMEDRIVER ")
    print(f"[-] Erro: {e}")
    exit()


print(f"[+] Tecle Ctrl+C para finalizar a aplicacao")
try:
    navegador.get("<<URL>>")
except Exception as e:
    print('[!] Erro na requisicao ao site')
    print(f"[-] Erro: {e}")
    exit()

WebDriverWait(navegador, 360).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar"]/ul/li[5]/a')))  # Click1
sleep(1.5)
navegador.get('<<URL>>')

lastLine = 0
keys = []

testaNew()

try:
    database = getData(linha)
    keys = setList(qtdtelefones)
    automation()
    saveDatabase(linha)
    if linha + 300 < totalLinhas:
        linha += 300
    else:
        linha += (totalLinhas - linha)
    for i in range(0, math.ceil((totalLinhas - linha) / 300)):
        database = getData(linha)
        keys = setList(qtdtelefones)
        automation()
        saveDatabase(linha-1)
        if linha + 300 < totalLinhas:
            linha += 300
        else:
            linha += (totalLinhas - linha)
    navegador.quit()
except OSError as err:
    print("[!] OS error:", err)
    navegador.quit()
    saveDatabase(linha)
except ZeroDivisionError as err:
    print('[!] Erro em tempo de execucao:', err)
    saveDatabase(linha)
    navegador.quit()
