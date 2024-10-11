import os
from dotenv import load_dotenv
from puxar_planilhas_sharepoint import puxar_planilhas
from gerar_planilha_nib import gerar_planilha_nib
from gerar_planilha_bmaisp import gerar_planilha_bmaisp
from office365_api.upload_files import upload_files

# carregar .env 
load_dotenv()
ROOT = os.getenv('ROOT')

PASTA_ARQUIVOS = os.path.abspath(os.path.join(ROOT, 'outputs'))

def gerar_planilhas():
    puxar_planilhas()
    gerar_planilha_nib()
    gerar_planilha_bmaisp()
    upload_files(PASTA_ARQUIVOS, "DWPII/nib_e_bmaisp")
    print(f'Planilhas enviadas para o Sharepoint com sucesso.')

gerar_planilhas()