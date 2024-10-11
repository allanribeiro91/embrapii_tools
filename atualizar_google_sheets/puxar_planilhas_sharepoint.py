import os
import sys
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
    apagar_arquivos_pasta(inputs)

    get_file('portfolio.xlsx', 'DWPII/srinfo', inputs)
    get_file('projetos_empresas.xlsx', 'DWPII/srinfo', inputs)
    get_file('informacoes_empresas.xlsx', 'DWPII/srinfo', inputs)
    get_file('info_unidades_embrapii.xlsx', 'DWPII/srinfo', inputs)
    get_file('ue_linhas_atuacao.xlsx', 'DWPII/srinfo', inputs)
    get_file('macroentregas.xlsx', 'DWPII/srinfo', inputs)
    print('Download concluído')

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


