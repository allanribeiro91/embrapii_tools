import os
import zipfile
from datetime import datetime
import inspect

def zipar_arquivos(origem, destino):
    """
    Função para zipar arquivos de uma pasta
        origem: str - Caminho da pasta com os arquivos a serem zipados
        destino: str - Caminho da pasta onde o arquivo zip será salvo
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
        
        current_datetime = datetime.now().strftime('%Y.%m.%d_%Hh%Mm%Ss')
        arquivo_zip = os.path.join(destino, f'consultas_clickhouse_{current_datetime}.zip')

        # Cria um objeto ZipFile no modo de escrita
        with zipfile.ZipFile(arquivo_zip, 'w') as zipf:
            # Percorre todos os arquivos da pasta
            for root, dirs, files in os.walk(origem):
                for file in files:
                    # Caminho completo do arquivo
                    file_path = os.path.join(root, file)
                    # Adiciona o arquivo ao ZIP, preservando a estrutura de pastas
                    zipf.write(file_path, os.path.relpath(file_path, origem))

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")
