import os
import sys
from dotenv import load_dotenv

# carregar .env e tudo mais
load_dotenv()
ROOT = os.getenv('ROOT')

# Adiciona o diretório correto ao sys.path
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))
sys.path.append(PATH_OFFICE)

from office365_api.download_files import get_file

# puxar planilhas do sharepoint
def puxar_planilhas(arquivos, destino):  
    for arquivo, caminho in arquivos.items():
        get_file(arquivo, caminho, destino)
    
    print('Download concluído')


# Exemplo de uso da função:
arquivos_dict = {
    'portfolio.xlsx': 'DWPII/srinfo',
    'projetos_empresas.xlsx': 'DWPII/srinfo',
    'informacoes_empresas.xlsx': 'DWPII/srinfo',
    'info_unidades_embrapii.xlsx': 'DWPII/srinfo',
    'ue_linhas_atuacao.xlsx': 'DWPII/srinfo',
    'macroentregas.xlsx': 'DWPII/srinfo'
}


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

