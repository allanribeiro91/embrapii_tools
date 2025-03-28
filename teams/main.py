from dotenv import load_dotenv
import os
from scripts.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts.mover_arquivos import mover_arquivos_excel
from scripts.buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from scripts.mensagem_chat_teams import mensagem_teams, enviar_mensagem_teams
import inspect

# Carregar as variáveis de ambiente
load_dotenv()
ROOT = os.getenv('ROOT')
PLANILHAS = os.path.abspath(os.path.join(ROOT, 'planilhas'))
ANTERIOR = os.path.abspath(os.path.join(ROOT, 'anterior'))

def main(atualizar = True, enviar = True):
    """
    Função geral para atualizar as planilhas de referência, obter os dados e enviar mensagem no Teams.
        atualizar: bool - Atualizar as planilhas de referência
        enviar: bool - Enviar mensagem
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:

        if atualizar:
            # Apagar arquivos da pasta "anterior"
            apagar_arquivos_pasta(ANTERIOR)

            # Mover arquivos de "planilhas" para "anterior"
            mover_arquivos_excel(2, PLANILHAS, ANTERIOR)

            # Buscar arquivos no Sharepoint
            buscar_arquivos_sharepoint()
        
        # Mensagem
        mensagem = mensagem_teams()
        
        # Enviar mensagem
        if enviar:
            enviar_mensagem_teams(mensagem)

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")

if __name__ == '__main__':
    main(atualizar = True, enviar = True)

