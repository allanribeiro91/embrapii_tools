import os
import sys
from dotenv import load_dotenv
from scripts_public.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts_public.main_registros_financeiros import main_registros_financeiros
from scripts_public.main_anexo8 import main_anexo8
from scripts_public.main_repasses import main_repasses
from scripts_public.main_plano_financeiro import main_plano_financeiro
from scripts_public.levar_arquivos_sharepoint import levar_arquivos_sharepoint
import inspect

load_dotenv()
ROOT = os.getenv('ROOT')
sys.path.append(ROOT)

ARQUIVOS_BRUTOS = os.path.abspath(os.path.join(ROOT, '1_data_raw'))
ARQUIVOS_PROCESSADOS = os.path.abspath(os.path.join(ROOT, '2_data_processed'))
BACKUP = os.path.abspath(os.path.join(ROOT, 'backup'))

def main(anexo8 = False, registros_financeiros = False, repasses = False, plano = False, levar = False, an8_por_unidade = False, an8_por_projeto = False, an8_por_mes = False,
         an8_ano_especifico = None, an8_mes_especifico = None, an8_tirar_desqualificados = False):
    """
    Função principal que chama as funções de processamento dos arquivos brutos.
        anexo8: Se True, processa o Anexo 8.
        registros_financeiros: Se True, processa os Registros Financeiros.
        repasses: Se True, processa os Repasses.
        plano: Se True, processa os Planos Financeiros.
        levar: Se True, leva os arquivos processados para o Sharepoint.
        an8_por_unidade: Se True, processa o Anexo 8 por Unidade.
        an8_por_projeto: Se True, processa o Anexo 8 por Projeto.
        an8_por_mes: Se True, processa o Anexo 8 por Mês.
        an8_ano_especifico: Lista de anos específicos a serem processados no Anexo 8 (se houver)
        an8_mes_especifico: Lista de meses específicos a serem processados no Anexo 8 (se houver)
        an8_tirar_desqualificados: Se True, tira os desqualificados do Anexo 8.
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
    
        if anexo8 or registros_financeiros or repasses:
            print("Apagando arquivos das pastas.")
            apagar_arquivos_pasta(ARQUIVOS_BRUTOS)
            apagar_arquivos_pasta(ARQUIVOS_PROCESSADOS)
            apagar_arquivos_pasta(BACKUP)

        if anexo8:
            main_anexo8(an8_por_unidade, an8_por_projeto, an8_por_mes, an8_mes_especifico, an8_ano_especifico, an8_tirar_desqualificados)
            # shutil.move(os.path.abspath(os.path.join(ARQUIVOS_BRUTOS, 'anexo8.csv')), ARQUIVOS_PROCESSADOS)

        if registros_financeiros:
            main_registros_financeiros()

        if repasses:
            main_repasses()

        if plano:
            main_plano_financeiro()

        if levar:
            print("Levando arquivos processados para o Sharepoint")
            levar_arquivos_sharepoint()

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")

if __name__ == "__main__":
    main(anexo8=True, registros_financeiros=True, repasses=True, plano=True,
         levar=True, an8_por_unidade=True, an8_por_projeto=True, an8_por_mes=True)