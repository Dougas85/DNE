import os
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import logging
import requests

def check_internet():
    try:
        _ = requests.get("http://www.google.com", timeout=5)
        return True
    except requests.ConnectionError:
        return False

def iniciar_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Rodar sem interface gráfica
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    navegador = webdriver.Chrome(options=options)
    navegador.get("https://apps3.correios.com.br/dne/#/")
    sleep(2)
    return navegador

def login(navegador):
    navegador.find_element(By.ID, "username").send_keys("81058616")
    navegador.find_element(By.ID, "password").send_keys("Raket**2324")
    navegador.find_element(By.NAME, "submitBtn").click()
    sleep(2)

def pesquisar_localidade(navegador, cidade):
    navegador.get("https://apps3.correios.com.br/dne/#/logradouro/consultar")
    sleep(1)

def cadastrar_faixa_cep(navegador, row, index, cidade):
    try:
        estado = row['ESTADO,C,61']
        bairro = row['BAIRRO,C,61']
        tipo = row['TIPO,C,61']
        titulo = row['TITULO,C,61']
        preposicao = row['PREPOSICAO,C,61']
        logradouro = row['LOGRADOURO,C,200']
        CEPinicial = str(row['CEP,N,9'])[:5]
        CEPfinal = str(row['CEP,N,9'])[-3:]
        adicional = row.get('ADIC,C,61', '')
        SEIe = row.get('SEI,C,15', '')

        navegador.find_element(By.ID, "btCadastrar").click()
        sleep(1)
        navegador.find_element(By.CLASS_NAME, 'fa-search').click()
        sleep(0.2)
        navegador.find_element(By.ID, "ufModal").send_keys(estado)
        sleep(0.3)
        navegador.find_element(By.ID, "nomeLocalidade").send_keys(cidade)
        sleep(0.5)
        navegador.find_element(By.ID, "pesquisarLocalidades").click()
        sleep(0.5)

        botoes = navegador.find_elements(By.NAME, "selecionarLocalidade")
        for botao in botoes:
            try:
                parent = botao.find_element(By.XPATH, "./ancestor::tr")
                texto = parent.find_element(By.XPATH, f".//td[2][text()=\"{cidade}\"]")
                if texto:
                    texto.click()
                    botao.click()
                    break
            except:
                continue

        navegador.find_element(By.ID, "bairroAutoComplete").send_keys(bairro)
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//ul/li/div[text()=\"{bairro}\"]"))
        ).click()

        navegador.find_element(By.ID, "tipoLogradouroCombo").send_keys(tipo)
        navegador.find_element(By.ID, "tituloPatenteCombo").send_keys('' if pd.isna(titulo) else titulo)
        navegador.find_element(
            By.XPATH,
            '//*[@id="geral"]/div/section/div[1]/div[1]/form/div[4]/div/div[2]/select'
        ).send_keys('' if pd.isna(preposicao) else preposicao)

        navegador.find_element(By.ID, "nomeLogradouroInput").send_keys(logradouro)
        navegador.find_element(By.ID, "InicioCep").send_keys(CEPinicial)
        navegador.find_element(By.ID, "finalCepInput").send_keys(CEPfinal)

        if not pd.isna(adicional):
            navegador.find_element(By.ID, "cnjInputTextForm").send_keys(adicional)
            sleep(0.5)

        if not pd.isna(SEIe):
            navegador.find_element(By.ID, "documentoSei").send_keys(SEIe)
            sleep(0.5)

        navegador.find_element(
            By.XPATH,
            '//*[@id="geral"]/div/section/div[1]/div[1]/form/div[7]/div/div[2]/input'
        ).click()
        sleep(0.5)

        navegador.find_element(By.ID, "btnSalvation").click()
        sleep(1)

    except Exception as e:
        logging.error(f"Linha {index}: {e}")

def processar_excel(path):
    navegador = None
    try:
        while True:
            if not check_internet():
                print("Sem internet. Aguardando reconexão...")
                if navegador:
                    navegador.quit()
                sleep(5)
                continue

            navegador = iniciar_navegador()
            login(navegador)
            df = pd.read_excel(path, dtype=str, engine='openpyxl')

            for index, row in df.iterrows():
                cidade = row['CIDADE,C,61']
                pesquisar_localidade(navegador, cidade)
                cadastrar_faixa_cep(navegador, row, index, cidade)

            navegador.quit()
            print("Processamento concluído com sucesso.")
            break

    except Exception as e:
        logging.error(f"Erro durante o processamento: {e}")
        print(f"Erro: {e}")

    finally:
        if navegador:
            navegador.quit()

def processar_arquivo(caminho):
    if os.path.exists(caminho):
        processar_excel(caminho)
        return "Processamento finalizado."
    else:
        return f"Arquivo não encontrado: {caminho}"

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python robo.py caminho_do_arquivo.xlsx")
    else:
        caminho = sys.argv[1]
        resultado = processar_arquivo(caminho)
        print(resultado)
