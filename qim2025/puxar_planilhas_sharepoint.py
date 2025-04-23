import os
import sys
from dotenv import load_dotenv
from office365_api.download_files import get_file
import inspect

# carregar .env e tudo mais
load_dotenv()
ROOT = os.getenv('ROOT')
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))
UNIDADES_EMBRAPII = os.getenv('UNIDADES_EMBRAPII')
PROPOSTAS_TECNICAS = os.getenv('PROPOSTAS_TECNICAS')
PROSPECCOES = os.getenv('PROSPECCOES')
RECURSOS_HUMANOS_UES = os.getenv('RECURSOS_HUMANOS_UES')

# Adiciona o diretório correto ao sys.path
sys.path.append(PATH_OFFICE)

# puxar planilhas do sharepoint
def puxar_planilhas():
    print("🟡 " + inspect.currentframe().f_code.co_name)
    inputs = os.path.join(ROOT, "inputs")
    apagar_arquivos_pasta(inputs)

    get_file(UNIDADES_EMBRAPII, 'DWPII/srinfo', inputs)
    get_file(RECURSOS_HUMANOS_UES, 'DWPII/srinfo', inputs)
    get_file(PROSPECCOES, 'DWPII/srinfo', inputs)
    get_file(PROPOSTAS_TECNICAS, 'DWPII/srinfo', inputs)

    print("🟢 " + inspect.currentframe().f_code.co_name)

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
        print(f"🔴 Ocorreu um erro ao apagar os arquivos: {e}")

# puxar_planilhas()


