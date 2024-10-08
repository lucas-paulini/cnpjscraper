import glob
import pandas as pd
import openpyxl
import os

dados =glob.glob('planilhas\*.xlsx')

dados_historicos = pd.DataFrame()

for i in dados:
    tabela = pd.read_excel(i)
    dados_historicos = pd.concat([dados_historicos, tabela], axis=0, ignore_index=True)

dados_historicos.drop_duplicates('cnpj', keep='first', inplace=True)

# Define o nome base do arquivo
nome_arquivo = 'planilha-com-cnpj-unificado'

# Obtém o caminho do diretório atual
diretorio_atual = os.getcwd()

# Nome da pasta onde deseja salvar o arquivo
nome_pasta = 'planilhas unificadas'

# Concatena o caminho do diretório atual com o nome da pasta
caminho_pasta = os.path.join(diretorio_atual, nome_pasta)

# Verifica se a pasta existe e a cria se não existir
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)

for i in range(50):
    try:
        # Define o nome do arquivo com base no número de tentativas
        nome_arquivo_com_numero = nome_arquivo if i == 0 else f'{nome_arquivo}_{i}'

        # Define o caminho completo do arquivo
        caminho_arquivo = os.path.join(caminho_pasta, f'{nome_arquivo_com_numero}.xlsx')

        # Verifica se o arquivo com o nome atual já existe
        if os.path.exists(caminho_arquivo):
            # Se existir, continua para a próxima tentativa
            continue
        else:
            # Se não existir, salva o arquivo com o nome atual
            dados_historicos.to_excel(caminho_arquivo, index=False)
            print(f'Arquivo "{nome_arquivo_com_numero}.xlsx" salvo com sucesso.')
            break  # Sair do loop após salvar o arquivo com sucesso
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")
        continue