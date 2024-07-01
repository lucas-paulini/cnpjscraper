import glob
import pandas as pd
import openpyxl
import os

# Lista todos os arquivos Excel no diretório 'planilhas'
dados = glob.glob('planilhas/*.xlsx')

# Inicializa um DataFrame vazio para armazenar os dados unificados
dados_historicos = pd.DataFrame()

# Lê cada arquivo Excel e concatena os dados
for arquivo in dados:
    tabela = pd.read_excel(arquivo)
    dados_historicos = pd.concat([dados_historicos, tabela], axis=0, ignore_index=True)

# Remove duplicatas com base no CNPJ, mantendo apenas a primeira ocorrência
dados_historicos.drop_duplicates('cnpj', keep='first', inplace=True)

# Define o nome base do arquivo
nome_arquivo = 'planilha-com-cnpj-advogados-unificado'

# Obtém o caminho do diretório atual
diretorio_atual = os.getcwd()

# Nome da pasta onde deseja salvar o arquivo
nome_pasta = 'planilhas_unificadas'

# Concatena o caminho do diretório atual com o nome da pasta
caminho_pasta = os.path.join(diretorio_atual, nome_pasta)

# Verifica se a pasta existe e a cria se não existir
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)

# Verifica se a coluna 'uf' existe nos dados
if 'uf' in dados_historicos.columns:
    # Filtra apenas as linhas onde o valor da coluna 'uf' é igual a 'SP'
    dados_sp = dados_historicos[dados_historicos['uf'] == 'SP'].copy()
    
    # Salva os dados filtrados em um novo arquivo
    try:
        # Define o caminho completo do arquivo filtrado
        caminho_arquivo_sp = os.path.join(caminho_pasta, f'{nome_arquivo}_SP.xlsx')

        # Salva o arquivo com apenas os dados de São Paulo
        dados_sp.to_excel(caminho_arquivo_sp, index=False)
        print(f'Arquivo "{nome_arquivo}_SP.xlsx" salvo com sucesso.')
    except Exception as e:
        print(f"Erro ao salvar arquivo filtrado de SP: {e}")
else:
    print("A coluna 'uf' não existe nos dados.")
