import os
import shutil
from glob import glob
from datetime import datetime
from dotenv import load_dotenv
from conveter_pdf_to_png import pdf_para_imagem

load_dotenv()

PASTA_DOWNLOAD = os.getenv('PASTA_DOWNLOAD')
ROOT = os.getenv('ROOT')

def mover_arquivo(numero_arquivos = 1):

    nome_arquivo = 'embrapii_status_semanal_' + datetime.today().strftime('%Y_%m_%d')
    destino = os.path.join(ROOT, 'arquivo_pdf')

    apagar_arquivos_pasta(destino)

    #Lista todos os arquivos Excel na pasta Downloads
    files = glob(os.path.join(PASTA_DOWNLOAD, '*.pdf'))
    
    #Ordena os arquivos por data de modificação (mais recentes primeiro)
    files.sort(key=os.path.getmtime, reverse=True)
    
    #Seleciona os n arquivos mais recentes
    files_to_move = files[:numero_arquivos]
    
    # Move os arquivos selecionados para a pasta data_raw com renome
    for i, file in enumerate(files_to_move, start=1):
        novo_nome = f"{nome_arquivo}.pdf"
        novo_caminho = os.path.join(destino, novo_nome)
        try:
            shutil.move(file, novo_caminho)
        except Exception as e:
            print(f'Erro ao mover {file} para {novo_caminho}. Razão: {e}')
    
    pdf_para_imagem(novo_nome)




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