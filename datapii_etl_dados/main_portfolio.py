import os
import sys
import pandas as pd
from dotenv import load_dotenv

# Adicione o caminho ao sys.path
sys.path.append(os.path.abspath(".."))

from puxar_planilhas_sharepoint import puxar_planilhas, apagar_arquivos_pasta
# Carregar variáveis de ambiente
load_dotenv()
ROOT = os.getenv('ROOT')
INPUT = os.path.join(ROOT, "datapii_etl_dados", "inputs")
OUTPUT = os.path.join(ROOT, "datapii_etl_dados", "output")

arquivos_dict = {
    'portfolio.xlsx': 'DWPII/srinfo',
}

def main_portfolio():
    apagar_arquivos_pasta(INPUT)
    puxar_planilhas(arquivos_dict, INPUT)
    caminho = os.path.join(INPUT, "portfolio.xlsx")
    destinho = os.path.join(OUTPUT, "portfolio_transformado.xlsx")
    etl_portfolio(caminho, "codigo_projeto", destinho)



def etl_portfolio(caminho, nome_chave, destino):
    # Carrega o arquivo Excel
    df = pd.read_excel(caminho)

    # Inicializa uma lista para armazenar os dados transformados
    dados_transformados = []

    # Itera sobre cada linha e campo do DataFrame original
    # Inicializa o contador para o campo "id"
    id_sequencial = 1

    # Itera sobre cada linha e campo do DataFrame original
    for index, row in df.iterrows():
        for coluna in df.columns:
            # Cria uma nova linha de dados transformados conforme a estrutura solicitada
            nova_linha = {
                "id": id_sequencial,  # número sequencial único
                "fonte": "srinfo",
                "tabela": "portfolio",
                "id_registro": row[nome_chave],  # valor da chave específica
                "campo": coluna,
                "tipo": str(type(row[coluna]).__name__),  # tipo de dado do campo
                "valor": row[coluna]  # valor do campo
            }
            # Adiciona a linha transformada à lista
            dados_transformados.append(nova_linha)
            # Incrementa o contador para o próximo ID
            id_sequencial += 1

    # Cria um novo DataFrame a partir da lista de dados transformados
    df_transformado = pd.DataFrame(dados_transformados)

    # Salva o DataFrame transformado no arquivo de saída
    df_transformado.to_excel(destino, index=False)
    print("Arquivo salvo em:", destino)

if __name__ == "__main__":
    main_portfolio()



