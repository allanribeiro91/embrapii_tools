import os
from dotenv import load_dotenv
from office365_api.upload_files import upload_files

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

#Definição dos caminhos
PASTA_ARQUIVOS = os.path.abspath(os.path.join(ROOT, 'arquivo_pdf'))
SHAREPOINT_SITE = os.getenv('sharepoint_url_site')
SHAREPOINT_SITE_NAME = os.getenv('sharepoint_site_name')
SHAREPOINT_DOC = os.getenv('sharepoint_doc_library')
SHAREPOIN_PASTA_DESTINO = os.getenv('sharepoint_pasta_destino')


def levar_arquivos_sharepoint():
    upload_files(PASTA_ARQUIVOS, SHAREPOIN_PASTA_DESTINO, SHAREPOINT_SITE, SHAREPOINT_SITE_NAME, SHAREPOINT_DOC)

