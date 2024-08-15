import os
import re

caminho_pasta = r"Z:\Operacional\Fiscoplan\CONFORMIDADE\J. C. M. NITEROI REFRIGERACAO LTDA\DOCUMENTOS\05-2024\SPED FISCAL"
cont = 1

# Verifica se o caminho é um diretório
if os.path.isdir(caminho_pasta):
    # Lista os arquivos na pasta
    arquivos = os.listdir(caminho_pasta)

    # Criando o arquivo de log
    log_file = os.path.join(caminho_pasta, "log_file.txt")
    with open(log_file, "w"):
        pass
    with open(log_file, "a") as log:
        log.write(f"Iniciando a análise dos arquivos da pasta {caminho_pasta}\n")
    
    # Itera sobre cada arquivo na pasta
    for nome_arquivo in arquivos:
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)

        try:
            # Formatando as strings a serem apresentadas
            infname_arquivo = nome_arquivo.split('-')
            cnpj = infname_arquivo[0]
            format_cnpj = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
            competencia = infname_arquivo[2]
            format_comp = f"{competencia[4:6]}/{competencia[:4]}"

            # Verifica se o caminho é um arquivo regular
            if os.path.isfile(caminho_arquivo):
                try:
                    # Abre o arquivo
                    with open(caminho_arquivo, 'r', encoding='utf-8', errors='ignore') as arquivo:
                        print(f"{cont}º: CNPJ:{format_cnpj}, {format_comp}")
                        with open(log_file, "a") as log:
                            log.write(f"\n\n{cont}º: CNPJ: {format_cnpj}, {format_comp}\n")

                            # Define os padrões para encontrar movimentos nos blocos do sped
                            padraoC100 = r'^\|C100\|'
                            padraoC500 = r'^\|C500\|'
                            padraoA100 = r'^\|A100\|'
                            padraoF100 = r'^\|F100\|'
                            padraoD100 = r'^\|D100\|'
                            padraoF550 = r'^\|F550\|'
                            nfce = r"^\|C100\|[^|]*\|[^|]*\|[^|]*\|65\|"

                            # Variáveis booleanas para testar linhas
                            encontrouC100 = False
                            encontrouC500 = False
                            encontrouA100 = False
                            encontrouF100 = False
                            encontrouD100 = False
                            encontrouF550 = False
                            encontrouNFCe = False

                            # Verifica cada linha do arquivo
                            for linha in arquivo:
                                
                                # Verifica se a linha corresponde ao padrão
                                if re.search(padraoC100, linha) and not encontrouC100:
                                    print(f"Movimento de C100 encontrado")
                                    log.write("|Possui movimento de C100|")
                                    encontrouC100 = True
                                if re.search(padraoC500, linha) and not encontrouC500:
                                    print(f"Movimento de C500 encontrado")
                                    log.write("|Possui movimento de C500|")
                                    encontrouC500 = True
                                if re.search(padraoA100, linha) and not encontrouA100:
                                    print(f"Movimento de A100 encontrado")
                                    log.write("|Possui movimento de A100|")
                                    encontrouA100 = True
                                if re.search(padraoF100, linha) and not encontrouF100:
                                    print(f"Movimento de F100 encontrado")
                                    log.write("|Possui movimento de F100|")
                                    encontrouF100
                                if re.search(padraoD100, linha) and not encontrouD100:
                                    print(f"|Movimento de D100 encontrado|")
                                    log.write("|Possui movimento de D100|")
                                    encontrouD100 = True
                                if re.search(padraoF550, linha) and not encontrouF550:
                                    print(f"Movimento de F550 encontrado")
                                    log.write("|Possui movimento de F550|")
                                    encontrouF550 = True
                                if re.search(nfce, linha) and not encontrouNFCe:
                                    print(f"Movimento de NFCe encontrado")
                                    log.write(f"\nPossui movimento de NFCe na linha {linha}")
                                    encontrouNFCe = True

                            if not (encontrouA100 or encontrouC100 or encontrouC500 or encontrouD100 or encontrouF100 or encontrouF550 or encontrouNFCe):
                                print("Arquivo não possui movimento")
                                log.write("O arquivo não possui movimento fiscal")
                            

                except Exception as e:
                    print(f'Erro ao ler o arquivo "{nome_arquivo}": {str(e)}')
                    with open(log_file, "a") as log:
                        log.write(f"Não foi possível ler o arquivo {nome_arquivo}")

                cont += 1

            else:
                print(f'O caminho "{arquivo}" não é um diretório válido.')
        except Exception as e:
            print(f"Não foi possível ler o arquivo {nome_arquivo}")
            with open(log_file, "a") as log:
                log.write(f"Não foi possível ler o arquivo {nome_arquivo}")
else:
    print(f'O caminho "{caminho_pasta}" não é um diretório válido.')