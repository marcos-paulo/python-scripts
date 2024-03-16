import os
import time
import pandas as pd

# Definir o diretório da pasta a ser monitorada
pasta_alvo = '/home/marcos/Downloads'

# Variável para armazenar o arquivo anterior
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
        if novo_arquivo != arquivo_anterior:
            # Fechar o arquivo anterior (caso exista)
            # if planilha_anterior is not None:
                # planilha_anterior.close()
                
            # Abrir o novo arquivo
            novo_arquivo_caminho = os.path.join(pasta_alvo, novo_arquivo)
            planilha = pd.read_excel(novo_arquivo_caminho)
            
            # Exemplo: imprimir o conteúdo da planilha
            print(planilha)
            
            # Atualizar o arquivo anterior
            arquivo_anterior = novo_arquivo
            planilha_anterior = planilha
            
    # Aguardar um intervalo de tempo antes de verificar novamente
    time.sleep(1)  # Intervalo de 10 segundos (pode ser ajustado conforme necessário)
