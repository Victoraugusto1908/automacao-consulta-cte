from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Lendo a planilha
caminho_arquivo = "relatorio.xlsx"
try:
    df = pd.read_excel(caminho_arquivo)
    print("Planilha lida com sucesso")
except:
    print("Não foi possível ler o arquivo")

# Selecinando a coluna das chaves de acesso
df_chaves = df['Chave acesso']
print("Coluna selecionada")

# Para abrir um navegador
navegador = webdriver.Chrome()
try:
    navegador.get("https://www.cte.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=cktLvUUKqh0=")
    print("Navegador aberto e página encontrada.")
except:
   print("Página não encontrada")

# Selecionando o campo de entrada
try:
    campo_input = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtChaveAcessoResumo")
    print("Campo selecionado com sucesso")
    try:
        campo_input.send_keys(df_chaves[1])

        # Clicando no captcha
        try:
            botao_captcha = navegador.find_element(By.CLASS_NAME, "h-captcha")
            print(f"Botão {'Captcha'} encontrado")
            botao_captcha.click()
            time.sleep(3)

            # Clicando para continuar
            try:
                botao_continuar = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnConsultarHCaptcha" )
                print(f"Botão {'Continuar'} encontrado")
                botao_continuar.click()
                time.sleep(5)
            except:
                print(f"Não foi possivel clicar no botão {'Continuar'}")
        except:
            print("Não foi possível validar o captcha")
            time.sleep(5)

    except: # Para quando não for possível realizar o input
        print("Não foi possível digitar este texto")
except: # Para quando não for possível selecionar o campo desejado
    print("Não foi possível selecionar o campo desejado")
