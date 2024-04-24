import pandas as pd
import requests
from lxml import html
import time
import openpyxl
import os

# Define a URL onde vamos fazer as consultas
url = 'https://api.casadosdados.com.br/v2/public/cnpj/search'

# Define o HEADER para não sermos bloqueados
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
    'Content-Type': 'application/json'
}

# Cria o DataFrame vazio para usarmos daqui a pouco
df_final = pd.DataFrame()

print('Iniciando scrap de dados...')

# Loop que faz a paginação
for i in range(1, 28):
    data = {
    "query": {
        "termo": [],
        "atividade_principal": ["6911701"],
        "natureza_juridica": [],
        "uf": [
            "SP"
        ],
        "municipio": [],
        "bairro": [],
        "situacao_cadastral": "ATIVA",
        "cep": [],
        "ddd": []
    },
    "range_query": {
        "data_abertura": {
            "lte": None,
            "gte": "2023-06-01"
        },
        "capital_social": {
            "lte": None,
            "gte": None
        }
    },
    "extras": {
        "somente_mei": False,
        "excluir_mei": False,
        "com_email": True,
        "incluir_atividade_secundaria": False,
        "com_contato_telefonico": True,
        "somente_fixo": False,
        "somente_celular": False,
        "somente_matriz": False,
        "somente_filial": False
    },
    "page": i
}
    print(f'Raspando página {i} ', end='')
    # Realiza a solicitação POST
    response = requests.post(url, json=data, headers=headers)

    # Verifica se a solicitação foi bem-sucedida
    if response.status_code == 200:
        # Processa a resposta da API (pode variar dependendo da estrutura da resposta)
        resultado = response.json()
    else:
        print(f'Erro na solicitação (Código {response.status_code}): {response.text}')
        break

    df_provisorio = pd.json_normalize(resultado, ['data', 'cnpj'])
    df_final = pd.concat([df_final, df_provisorio], axis=0)

    print(f'- OK')

    #time.sleep(1)

print('Scraping inicial feito com suscesso!')

# Aqui inicia a busca por dados adicionais ---------------------------------------------------------------
print('Iniciando extração completa dos dados...')

# Vamos usar o DataFrame df_final como fonte dos dados
url = []
for razao, cnpj in zip(df_final['razao_social'], df_final['cnpj']):
    url.append('https://casadosdados.com.br/solucao/cnpj/' + razao.replace(' ', '-').replace('.', '').replace('&', 'and').replace('/', '').replace('*', '').replace('--', '-').lower() + '-' + cnpj)
    
# Inicia as listas vazias para receberem os dados
lista_email = []
lista_tel = []
lista_socio1 = []
lista_socio2 = []
lista_socio3 = []
lista_socio4 = []
lista_socio5 = []
lista_capital_social = []

# Função que verifica se uma variável é número - usaremos na validação do capital social
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

# Itera sobre a lista de URLs
for indice, link in enumerate(url):
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    headers = {'User-Agent': user_agent}

    response = requests.get(url[indice], headers=headers)

    print(f'Empresa {indice + 1}/{len(df_final)} ', end='')

    # Verificar se a requisição foi bem-sucedida
    if response.status_code == 200:
        # Parsear o conteúdo HTML da página
        page_content = html.fromstring(response.content)

        # Usar os XPath para encontrar os elementos desejados
        email = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[3]/div[2]/p[2]/a')
        tel = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[3]/div[1]/p[2]/a')

        # Verificar se o elemento EMAIL foi encontrado
        if email:
            # Acessar o texto do elemento e adicionar à lista de email
            lista_email.append(email[0].text_content().lower())
        else:
            lista_email.append('')
        # Verificar se o elemento TEL foi encontrado
        if tel:
            # Acessar o texto do elemento e adicionar à lista de tel
            lista_tel.append(tel[0].text_content())
        else:
            lista_email.append('')
        
        socio1 = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[4]/div/p[2]/b')
        socio2 = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[4]/div/p[3]/b')
        socio3 = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[4]/div/p[4]/b')
        socio4 = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[4]/div/p[5]/b')
        socio5 = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[4]/div/p[6]/b')

        if socio1:
            lista_socio1.append(socio1[0].text_content())
        else:
            lista_socio1.append('')

        if socio2:
            lista_socio2.append(socio2[0].text_content())
        else:
            lista_socio2.append('')

        if socio3:
            lista_socio3.append(socio3[0].text_content())
        else:
            lista_socio3.append('')

        if socio4:
            lista_socio4.append(socio4[0].text_content())
        else:
            lista_socio4.append('')

        if socio5:
            lista_socio5.append(socio5[0].text_content())
        else:
            lista_socio5.append('')

        capital_social = page_content.xpath('//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[1]/div[8]/p[2]')[0].text_content().replace('R$ ', '').replace('.', '').replace(',', '')

        if is_number(capital_social):
            lista_capital_social.append(int(capital_social))
        else:
            lista_capital_social.append('')

    else:
        lista_email.append('ERRO 404')
        lista_tel.append('ERRO 404')
        socio1.append('ERRO 404')
        socio2.append('ERRO 404')
        socio3.append('ERRO 404')
        socio4.append('ERRO 404')
        socio5.append('ERRO 404')
    
    print(f'- OK')
    #Espera 1 segundo antes de ir para a próxima iteração
    #time.sleep(1)

print('Dados adicionais extraídos com sucesso!')

df_dados_extraidos = pd.DataFrame({
    'TELEFONE': lista_tel,
    'EMAIL': lista_email,
    'SÓCIO 1': lista_socio1,
    'SÓCIO 2': lista_socio2,
    'SÓCIO 3': lista_socio3,
    'SÓCIO 4': lista_socio4,
    'SÓCIO 5': lista_socio5,
    'CAPITAL SOCIAL': lista_capital_social,
})

df_final = df_final.reset_index(drop=True)  # Redefine o índice
df_dados_extraidos = df_dados_extraidos.reset_index(drop=True)  # Redefine o índice
df_consolidado = pd.concat([df_final, df_dados_extraidos], axis=1)

print('Salvando no arquivo XLSX...')


# Define o nome base do arquivo
nome_arquivo = 'planilha-com-cnpj-advogados'

# Obtém o caminho do diretório atual
diretorio_atual = os.getcwd()

# Nome da pasta onde deseja salvar o arquivo
nome_pasta = 'planilhas'

# Concatena o caminho do diretório atual com o nome da pasta
caminho_pasta = os.path.join(diretorio_atual, nome_pasta)

# Verifica se a pasta existe e a cria se não existir
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)


for i in range(500):
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
            df_consolidado.to_excel(caminho_arquivo, index=False)
            print(f'Arquivo "{nome_arquivo_com_numero}.xlsx" salvo com sucesso.')
            break  # Sair do loop após salvar o arquivo com sucesso
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")
        continue