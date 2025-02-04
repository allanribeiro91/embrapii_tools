import os
import sys
from dotenv import load_dotenv
from scripts_public.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts_public.main_registros_financeiros import main_registros_financeiros
from scripts_public.main_anexo8 import main_anexo8
from scripts_public.levar_arquivos_sharepoint import levar_arquivos_sharepoint

load_dotenv()
ROOT = os.getenv('ROOT')
sys.path.append(ROOT)

ARQUIVOS_BRUTOS = os.path.abspath(os.path.join(ROOT, '1_data_raw'))
ARQUIVOS_PROCESSADOS = os.path.abspath(os.path.join(ROOT, '2_data_processed'))
BACKUP = os.path.abspath(os.path.join(ROOT, 'backup'))

def main(anexo8 = False, registros_financeiros = False, levar = False, por_unidade = False, por_projeto = False, por_mes = False,
         ano_especifico = None, mes_especifico = None, tirar_desqualificados = False):
    
    print("Apagando arquivos das pastas.")
    apagar_arquivos_pasta(ARQUIVOS_BRUTOS)
    apagar_arquivos_pasta(ARQUIVOS_PROCESSADOS)
    apagar_arquivos_pasta(BACKUP)

    if anexo8:
        main_anexo8(por_unidade, por_projeto, por_mes, mes_especifico, ano_especifico, tirar_desqualificados)
    
    if registros_financeiros:
        main_registros_financeiros()

    if levar:
        print("Levando arquivos processados para o Sharepoint")
        levar_arquivos_sharepoint()

if __name__ == "__main__":
    main(anexo8=True, registros_financeiros=True, levar=True,
         por_unidade=True, por_projeto=True, por_mes=True)