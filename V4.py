import pandas as pd
import requests
from lxml import html
import time
import os


def loopgeral():
    # Define a URL onde vamos fazer as consultas
    url = 'https://api.casadosdados.com.br/v2/public/cnpj/search'

    # Define o HEADER para não sermos bloqueados
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
        'Content-Type': 'application/json'
    }

    # Cria o DataFrame vazio para os resultados
    df_final = pd.DataFrame()

    print('Iniciando scraping de dados...')

    # Loop para paginação
    total_paginas = 50  # Número de páginas a serem raspadas
    for i in range(1, total_paginas + 1):
        data = {
            "query": {
                "termo": [],
                "atividade_principal": ["7500100"],  # Código CNAE
                "natureza_juridica": [],
                "uf": ["RJ"],  # Estado
                "municipio": [],
                "bairro": [],
                "situacao_cadastral": "ATIVA",  # Empresas ativas
                "cep": [],
                "ddd": []
            },
            "range_query": {
                "data_abertura": {
                    "lte": None,
                    "gte": "2020-01-01"  # Empresas abertas após XXXX
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
        
        print(f'Raspando página {i}/{total_paginas}...', end=' ')

        # Realiza a solicitação POST
        response = requests.post(url, json=data, headers=headers)

        if response.status_code == 200:
            resultado = response.json()
            # Normaliza e concatena os resultados
            df_provisorio = pd.json_normalize(resultado, ['data', 'cnpj'])
            df_final = pd.concat([df_final, df_provisorio], axis=0)
            print('OK')
        else:
            print(f'Erro {response.status_code}: {response.text}')
            break
        
        time.sleep(2)  # Pausa entre solicitações para evitar bloqueio (no caso, cloudflare)

    print('Scraping inicial feito com sucesso!')

    # Iniciar a busca por dados adicionais
    print('Iniciando extração completa dos dados...')

    # Lista de URLs para scraping adicional
    url_list = []
    for razao, cnpj in zip(df_final['razao_social'], df_final['cnpj']):
        url_list.append(f'https://casadosdados.com.br/solucao/cnpj/{razao.replace(" ", "-").lower()}-{cnpj}')

    # Listas para armazenar os dados
    lista_email = []
    lista_tel1 = []
    lista_tel2 = []
    lista_socio1 = []
    lista_socio2 = []
    lista_socio3 = []
    lista_socio4 = []
    lista_socio5 = []
    lista_capital_social = []

    # Função que verifica se uma variável é número
    def is_number(s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    # Loop para scraping adicional de cada URL
    total_empresas = len(df_final)
    for indice, link in enumerate(url_list):
        print(f'Extraindo dados da empresa {indice + 1}/{total_empresas}...', end=' ')

        response = requests.get(link, headers=headers)

        if response.status_code == 200:
            page_content = html.fromstring(response.content)

            # Extração dos campos por XPath
            email = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[19]/p/a')
            tel1 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[20]/p[1]/a[1]')
            tel2 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[20]/p[2]/a[1]')

            socio1 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[1]')
            socio2 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[2]')
            socio3 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[3]')
            socio4 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[4]')
            socio5 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[5]')

            # Capital social
            capital_social = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[10]/p')[0].text_content().replace('R$ ', '').replace('.', '').replace(',', '')

            # Preenchimento das listas
            lista_email.append(email[0].text_content().lower() if email else '')
            lista_tel1.append(tel1[0].text_content() if tel1 else '')
            lista_tel2.append(tel2[0].text_content() if tel2 else '')
            lista_socio1.append(socio1[0].text_content() if socio1 else '')
            lista_socio2.append(socio2[0].text_content() if socio2 else '')
            lista_socio3.append(socio3[0].text_content() if socio3 else '')
            lista_socio4.append(socio4[0].text_content() if socio4 else '')
            lista_socio5.append(socio5[0].text_content() if socio5 else '')
            lista_capital_social.append(int(capital_social) if is_number(capital_social) else '')

            print('OK')
        else:
            print(f'Erro {response.status_code}')
            lista_email.append('ERRO')
            lista_tel1.append('ERRO')
            lista_tel2.append('ERRO')
            lista_socio1.append('ERRO')
            lista_socio2.append('ERRO')
            lista_socio3.append('ERRO')
            lista_socio4.append('ERRO')
            lista_socio5.append('ERRO')
            lista_capital_social.append('ERRO')

        #time.sleep(1)

    # Construção do DataFrame final
    df_dados_extraidos = pd.DataFrame({
        'TELEFONE1': lista_tel1,
        'TELEFONE2': lista_tel2,
        'EMAIL': lista_email,
        'SÓCIO 1': lista_socio1,
        'SÓCIO 2': lista_socio2,
        'SÓCIO 3': lista_socio3,
        'SÓCIO 4': lista_socio4,
        'SÓCIO 5': lista_socio5,
        'CAPITAL SOCIAL': lista_capital_social,
    })

    df_final = df_final.reset_index(drop=True)
    df_dados_extraidos = df_dados_extraidos.reset_index(drop=True)
    df_consolidado = pd.concat([df_final, df_dados_extraidos], axis=1)

    print('Salvando no arquivo XLSX...')

    # Definição do caminho e salvamento do arquivo
    nome_arquivo = 'planilha-com-cnpj'
    caminho_pasta = os.path.join(os.getcwd(), 'planilhas')

    # Cria o diretório se não existir
    if not os.path.exists(caminho_pasta):
        os.makedirs(caminho_pasta)

    for i in range(5000):
        try:
            nome_arquivo_com_numero = nome_arquivo if i == 0 else f'{nome_arquivo}_{i}'
            caminho_arquivo = os.path.join(caminho_pasta, f'{nome_arquivo_com_numero}.xlsx')

            if not os.path.exists(caminho_arquivo):
                df_consolidado.to_excel(caminho_arquivo, index=False)
                print(f'Arquivo "{nome_arquivo_com_numero}.xlsx" salvo com sucesso.')
                break
        except Exception as e:
            print(f'Erro ao salvar arquivo: {e}')
            continue
    loopgeral()

loopgeral()
