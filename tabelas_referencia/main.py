import os
from dotenv import load_dotenv
from puxar_planilhas_sharepoint import puxar_planilhas
from atualizacao_gsheet import atualizar_gsheet

# carregar .env 
load_dotenv()
ROOT = os.getenv('ROOT')

UNIDADES_EMBRAPII = os.path.abspath(os.path.join(ROOT, 'inputs', os.getenv('UNIDADES_EMBRAPII')))
RECURSOS_HUMANOS_UES = os.path.abspath(os.path.join(ROOT, 'inputs', os.getenv('RECURSOS_HUMANOS_UES')))
PROPOSTAS_TECNICAS = os.path.abspath(os.path.join(ROOT, 'inputs', os.getenv('PROPOSTAS_TECNICAS')))
PROSPECCOES = os.path.abspath(os.path.join(ROOT, 'inputs', os.getenv('PROSPECCOES')))

def main():
    puxar_planilhas()
    url = "https://docs.google.com/spreadsheets/d/1LAb7uMv6SyXX3aKES5lHCIZLinfAVncShYxxR1Lzp48/edit?usp=sharing"
    abas = {
        'raw_unidades_embrapii': UNIDADES_EMBRAPII,
        'raw_meta02': RECURSOS_HUMANOS_UES,
        'raw_meta05': PROSPECCOES,
        'raw_meta06': PROPOSTAS_TECNICAS,
    }

    for aba, caminho_arquivo in abas.items():
        atualizar_gsheet(url, aba, caminho_arquivo)
    

if __name__ == "__main__":
    main()
