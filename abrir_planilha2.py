import os
import time
import pandas as pd
import subprocess

# Definir o diretório da pasta a ser monitorada
pasta_alvo = '/home/marcos/Downloads'

# Variáveis para armazenar o processo do LibreOffice e o arquivo anterior
processo_libreoffice = None
arquivo_anterior = None
planilha_anterior = None

while True:
    # Listar todos os arquivos na pasta
    arquivos = os.listdir(pasta_alvo)
    
    # Filtrar apenas os arquivos .xlsx
    arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.endswith('.xlsx')]
    
    if arquivos_xlsx:
        # Ordenar os arquivos por data de modificação (o mais recente primeiro)
        arquivos_xlsx.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_alvo, x)), reverse=True)
        
        # Pegar o nome do arquivo mais recente
        novo_arquivo = arquivos_xlsx[0]
        
        # Verificar se o arquivo atual é diferente do arquivo anterior
        if arquivo_anterior != novo_arquivo:
            # Fechar o arquivo anterior (caso exista)
            if arquivo_anterior is not None:
                processo_libreoffice.terminate()
                processo_libreoffice.wait()
            
            # Abrir o novo arquivo no LibreOffice
            novo_arquivo_caminho = os.path.join(pasta_alvo, novo_arquivo)
            processo_libreoffice = subprocess.Popen(['libreoffice', '--calc', '--norestore', novo_arquivo_caminho])

            
            # Aguardar um tempo para o LibreOffice abrir o arquivo
            time.sleep(5)  # Ajuste o tempo conforme necessário
            
            # Ler o conteúdo do arquivo usando o Pandas
            planilha = pd.read_excel(novo_arquivo_caminho)
            
            # Exemplo: imprimir o conteúdo da planilha
            print(planilha)
            
            # Atualizar o arquivo anterior
            arquivo_anterior = novo_arquivo
            planilha_anterior = planilha

            print('-----------------------')
            print('-----------------------')
    
    # Aguardar um intervalo de tempo antes de verificar novamente
    time.sleep(1)  # Intervalo de 10 segundos (pode ser ajustado conforme necessário)
