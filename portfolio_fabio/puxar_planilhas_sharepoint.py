import os
import sys
import requests
import pandas as pd
from dotenv import load_dotenv
from office365_api.download_files import get_file

# carregar .env e tudo mais
load_dotenv()
ROOT = os.getenv('ROOT')
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))

# Adiciona o diretório correto ao sys.path
sys.path.append(PATH_OFFICE)

# puxar planilhas do sharepoint
def puxar_planilhas():
    inputs = os.path.join(ROOT, "inputs")
    up = os.path.join(ROOT, "up")
    url = 'https://apisidra.ibge.gov.br/values/t/1737/p/all/v/2266/N1/1?formato=json'
    apagar_arquivos_pasta(inputs)
    apagar_arquivos_pasta(up)

    get_file('portfolio.xlsx', 'DWPII/srinfo', inputs)
    get_file('projetos_empresas.xlsx', 'DWPII/srinfo', inputs)
    get_file('informacoes_empresas.xlsx', 'DWPII/srinfo', inputs)
    get_file('info_unidades_embrapii.xlsx', 'DWPII/srinfo', inputs)
    get_file('classificacao_projeto.xlsx', 'DWPII/srinfo', inputs)
    get_file('cnae_ibge.xlsx', 'DWPII/lookup_tables', inputs)
    get_file('pedidos_pi.xlsx', 'DWPII/srinfo', inputs)
    get_file('ibge_municipios.xlsx', 'DWPII/lookup_tables', inputs)

    # Fazer a requisição para obter os dados JSON
    response = requests.get(url)

    # Verificar se a requisição foi bem-sucedida
    if response.status_code == 200:
        # Carregar os dados JSON
        data = response.json()
    
        # Converter os dados para um DataFrame
        df = pd.DataFrame(data)

        df.to_excel(os.path.join(inputs, 'ipca.xlsx'), index=False)

    print('Downloads concluídos.')

def apagar_arquivos_pasta(caminho_pasta):
    try:
        # Verifica se o caminho é válido
        if not os.path.isdir(caminho_pasta):
            print(f"O caminho {caminho_pasta} não é uma pasta válida.")
            return
        
        # Lista todos os arquivos na pasta
        arquivos = os.listdir(caminho_pasta)
        
        # Apaga cada arquivo na pasta
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
    except Exception as e:
        print(f"Ocorreu um erro ao apagar os arquivos: {e}")

# puxar_planilhas()