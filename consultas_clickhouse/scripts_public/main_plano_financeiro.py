from dotenv import load_dotenv
import os
import sys
from scripts_public.consulta_clickhouse import consulta_clickhouse
from scripts_public.processar_csv import processar_csv
import inspect
import pandas as pd

load_dotenv()

ROOT = os.getenv('ROOT')
sys.path.append(ROOT)

HOST = os.getenv('HOST')
PORT = os.getenv('PORT')
USER = os.getenv('USER')
PASSWORD = os.getenv('PASSWORD')

pasta = '1_data_raw'
nome_arquivo = 'plano_financeiro'

def processar_plano_financeiro():
    """
    Função que processa o arquivo de plano financeiro,
    renomeando as colunas e gerando um novo arquivo
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:

        print("Gerando planilha de plano financeiro")

        # Ler o arquivo
        df = pd.read_csv(os.path.join(ROOT, '1_data_raw', 'plano_financeiro.csv'))

        # Renomear um valor específico
        df['Carteira'] = df['Carteira'].replace('EMBRAPII', 'CG')

        # Salvar o arquivo
        df.to_csv(os.path.join(ROOT, '1_data_raw', 'plano_financeiro.csv'), index=False)

        # Definições dos caminhos e nomes de arquivos
        origem = os.path.join(ROOT, '1_data_raw')
        destino = os.path.join(ROOT, '2_data_processed')
        arquivo_origem = os.path.join(origem, nome_arquivo)
        arquivo_destino = os.path.join(destino, nome_arquivo)

        campos_interesse = [
              'Unidade',
              'Termo',
              'PlanoAcao',
              'Carteira',
              'Recurso',
              'Ano',
              'Valor'
        ]

        novos_nomes_e_ordem = {
            'Unidade': 'unidade_embrapii',
            'Termo': 'termo_cooperacao',
            'PlanoAcao': 'plano_acao',
            'Carteira': 'carteira',
            'Recurso': 'recurso',
            'Ano': 'ano',
            'Valor': 'valor',
        }

        campos_valor = ['valor']

        processar_csv(arquivo_origem = arquivo_origem, campos_interesse = campos_interesse, novos_nomes_e_ordem = novos_nomes_e_ordem,
                      arquivo_destino = arquivo_destino, campos_valor = campos_valor)
        
        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")

def main_plano_financeiro():
    """
    Função principal que realiza a consulta ao ClickHouse e processa o plano financeiro
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    
    try:

        # Consulta ao ClickHouse
        query = """
                SELECT * from db_ouro.lista_planofinanceiro lp 
        """

        consulta_clickhouse(HOST, PORT, USER, PASSWORD, query, pasta, nome_arquivo)
        processar_plano_financeiro()

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")

if __name__== "__main__":
      main_plano_financeiro()