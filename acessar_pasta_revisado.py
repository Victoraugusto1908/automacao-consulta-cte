import os
import re

pasta = "speds"
com_movimento = []
sem_movimento = []

# Verifica se o caminho é um diretório
if os.path.isdir(pasta):
    # Lista os arquivos na pasta
    arquivos = os.listdir(pasta)

    # Caminhos completos para os arquivos de texto de saída
    caminho_completo_movimento = os.path.join(pasta, 'com_movimento.txt')
    caminho_completo_sem = os.path.join(pasta, 'sem_movimento.txt')

    # Itera sobre cada arquivo na pasta
    for nome_arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta, nome_arquivo)

        # Verifica se o caminho é um arquivo regular
        if os.path.isfile(caminho_arquivo):
            try:
                # Abre o arquivo
                with open(caminho_arquivo, 'r', encoding='utf-8', errors='ignore') as arquivo:
                    movimento_encontrado = False

                    # Verifica cada linha do arquivo
                    for linha in arquivo:
                        # Define o padrão para encontrar linhas com |C100|
                        padrao = r'^\|C100\|'
                        padrao2 = r'^\|C500\|'
                        padrao3 = r'^\|A100\|'
                        padrao4 = r'^\|F100\|'
                        padrao5 = r'^\|D100\|'
                        
                        # Verifica se a linha corresponde ao padrão
                        if re.search(padrao, linha):
                            movimento_encontrado = True
                            break  # Se encontrou, não precisa continuar verificando o arquivo
                        elif re.search(padrao2, linha):
                            movimento_encontrado = True
                            break
                        elif re.search(padrao3, linha):
                            movimento_encontrado = True
                            break  # Se encontrou, não precisa continuar verificando o arquivo
                        elif re.search(padrao4, linha):
                            movimento_encontrado = True
                            break  # Se encontrou, não precisa continuar verificando o arquivo
                        elif re.search(padrao5, linha):
                            movimento_encontrado = True
                            break  # Se encontrou, não precisa continuar verificando o arquivo

                    # Classifica o arquivo em com ou sem movimento
                    if movimento_encontrado:
                        com_movimento.append(nome_arquivo)
                        print(f'Arquivo "{nome_arquivo}" possui movimento e foi adicionado à lista.')
                    else:
                        sem_movimento.append(nome_arquivo)
                        print(f'Arquivo "{nome_arquivo}" NÃO possui movimento e foi adicionado à lista.')
            
            except Exception as e:
                print(f'Erro ao ler o arquivo "{nome_arquivo}": {str(e)}')

        else:
            print(f'O caminho "{pasta}" não é um diretório válido.')

    # Escreve as listas nos arquivos de texto
    with open(caminho_completo_movimento, 'w') as movimento_txt:
        for item in com_movimento:
            movimento_txt.write(f"{item}\n")

    with open(caminho_completo_sem, 'w') as sem_txt:
        for item in sem_movimento:
            sem_txt.write(f"{item}\n")

    print('Processo concluído.')

else:
    print(f'O caminho "{pasta}" não é um diretório válido.')