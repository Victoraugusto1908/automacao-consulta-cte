from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from undetected_chromedriver import Chrome, ChromeOptions
import pyautogui

# Abrindo o txt de log
log_file = open(r"C:\Users\victor.gomes\Desktop\automacoes\TESTE1\logfile.txt",'w')
log_file.write("Iniciando o processo...")
log_file.close()
log_file = open(r"C:\Users\victor.gomes\Desktop\automacoes\TESTE1\logfile.txt",'a')

# Lendo a planilha
caminho_arquivo = "relatorio.xlsx"
try:
    df = pd.read_excel(caminho_arquivo)
    print('Planilha lida com sucesso')
except:
    print("Não foi possível ler o arquivo")
    log_file.write("Não foi possível ler o arquivo do relatório.\n")


# Selecinando a coluna das chaves de acesso
df_chaves = df['Chaves']
print("Coluna selecionado com sucesso")
log_file.write("Relatório lido com sucesso...\n")

# Para abrir um navegadorn utilizando o undected_chromedriver
option = ChromeOptions()
navegador =Chrome(option=option)

# Para selecionar a pasta de download dos arquivos
option.add_experimental_option('prefs', {
    'download.default_directory': r'C:\Users\victor.gomes\Desktop\automacoes\TESTE1\Download',
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True})
log_file.write("Pasta para Download selecionada com sucesso...\n")

# Abrindo a pagina da consulta CTe
try:
    navegador.get("https://www.cte.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=cktLvUUKqh0=")
    print("Navegador aberto e página encontrada.")
    log_file.write("Navegador aberto e página encontrada...\n")
except:
    print("Página não encontrada")
    log_file.write("Erro na execução do navegador ou nõa foi possível encontrar a página.\n")
   
# Criando o laço para fazer isso durante toda a lista de chaves de acesso
for chave_acesso in df_chaves:
    print(f"Inicializando o download da chave: {chave_acesso}")
    log_file.write(f"\n -> Começando o processamento da chave: {chave_acesso}\n")

    # Selecionando o campo de entrada
    try:
        campo_input = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtChaveAcessoResumo")
        print("Campo selecionado com sucesso")
        log_file.write("Campo input_chave selecionado.\n")

        try:
            campo_input.send_keys(chave_acesso)
            log_file.write("Chave digitada no input.\n")

            # Clicando no captcha
            try:
                botao_captcha = navegador.find_element(By.CLASS_NAME, "h-captcha")
                print(f"Botão {'Captcha'} encontrado")
                botao_captcha.click()
                log_file.write("Captcha feito.\n")
                time.sleep(5)

                # Clicando para continuar
                try:
                    botao_continuar = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnConsultarHCaptcha" )
                    print(f"Botão {'Continuar'} encontrado")
                    botao_continuar.click()
                    log_file.write("Botão de continuar clicado.\n")

                except: # Se não conseguir continuar após a verificação do Captcha
                    print(f"Botão de Continuar não encontrado")
                    log_file.write("Botão de Continuar não encontrado.\n")

            except:
                print("Não foi possível validar o captcha")
                log_file.write("Erro no click do Captcha.\n")
                break

        except: # Para quando não for possível realizar o input
            print("Não foi possível digitar este texto")
            log_file.write("Não foi possível inserir este texto no input.\n")
            break

    except: # Para quando não for possível selecionar o campo desejado
        print("Não foi possível selecionar o campo desejado")
        log_file.write("Não achamos o input.\n")
        break

    # Clicando para baixar
    try:
        botao_download = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnDownload")
        print(f"Botão de {'Download'} encontrado")
        botao_download.click()
        log_file.write("Iniciando o download.\n")

        # Dando enter na tela interativa
        try:
            time.sleep(1)
            pyautogui.press('enter')
            print("Conseguimos pressionar o enter")

            # Clicando para voltar à uma nova consulta
            try:
                botao_voltar = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnVoltar")
                print(f"Botão{'Voltar'} encontrado")
                botao_voltar.click()
                log_file.write("Download desta chave finalizado com sucesso, inciando outra.\n")
                time.sleep(1)

            except: # Se não conseguir voltar para a pagina inicial
                print("Não foi possível voltar a pagina de consulta.")
                log_file.write("Não foi encontrado o botão de Voltar, verificar carregamento da página.\n")
                break

        except: # Para quando não for posível interagir com a tela de certificado.
            print("Não foi possível interagir com a tela de certificado")
            log_file.write("Não conseguimos interagir com a tela do certificado.\n")
            break

    except: # Para quando der algum erro no botão de download
        print("Não foi possível realizar o download")
        log_file.write("O código não encontrou o botão de download, muito provavél que o código não tenha acessado a página por conta do captcha. Vamos tentar quebrar o Captcha novamente.\n")
        falha_captcha = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_bltMensagensErroHCaptcha")
        if falha_captcha:
            time.sleep(1)

            # Fazendo todo o processo de captcha, escrever a chave e consultar novamente, mas com um tempo maior
            try:
                botao_limpar = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnLimparHCaptcha")
                botao_limpar.click()
                log_file.write("Botão de limpar clicado.\n")
                campo_input = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtChaveAcessoResumo")
                campo_input.send_keys(chave_acesso)
                log_file.write("Chave escrita no input.\n")
                botao_captcha = navegador.find_element(By.CLASS_NAME, "h-captcha")
                botao_captcha.click()
                log_file.write("Segunda tentativa do Captcha feita.\n")
                time.sleep(15)

                # Tentando avançar com o Captcha feito
                try:
                    botao_continuar = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnConsultarHCaptcha" )
                    botao_continuar.click()
                    log_file.write("Segunda tentativa foi efetuado com sucesso!\n")
                    print("Avançando de página")

                    botao_download = navegador.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnDownload")
                    botao_download.click()
                    log_file.write("Iniciando o download.\n")

                    time.sleep(1)
                    pyautogui.press('enter')

                    botao_voltar = navegador.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnVoltar")
                    botao_voltar.click()
                    log_file.write("Processamento concluído, inicializando outra chave.\n")

                except: # Se não foi possível avançar de página, é porque o Captcha travou, para processamento.
                    log_file.wirte("Mesmo após a segunda tentativa não foi possível continuar a operação, finalizando o processamento")
                    print("Fim do processo")
                
            except: # Caso não consiga clicar no Captcha novamente
                print("Botão do Captcha não foi encontrado.")
                log_file.write("Na segunda tentativa do Captcha não encontramos o botão.\n")
                break

        else:
            print("Não foi encontrada mensagem de erro de Captcha, parando o processamento.")
            log_file.write("Não foi encontrada mensagem de erro de Captcha, parando o processamento.")
            break

print("Fim do Processo")
navegador.quit()
