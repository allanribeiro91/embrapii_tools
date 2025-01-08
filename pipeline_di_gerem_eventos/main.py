import os
import sys
import shutil
import pandas as pd
from dotenv import load_dotenv
from office365_api.download_files import get_file
from office365_api.upload_files import upload_files

# carregar .env e tudo mais
load_dotenv()
ROOT = os.getenv('ROOT')

# Adiciona o diretório correto ao sys.path
# PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))
# sys.path.append(PATH_OFFICE)

#Definição dos caminhos do SHAREPOINT
SHAREPOINT_SITE = os.getenv('sharepoint_url_site')
SHAREPOINT_SITE_NAME = os.getenv('sharepoint_site_name')
SHAREPOINT_DOC = os.getenv('sharepoint_doc_library')

#Definição das pastas locais
STEP1 = os.path.join(ROOT, "step_1_data_raw")
STEP2 = os.path.join(ROOT, "step_2_stage_area")
STEP3 = os.path.join(ROOT, "step_3_data_processed")

def puxar_planilhas():    
    apagar_arquivos_pasta(STEP1)
    apagar_arquivos_pasta(STEP2)
    apagar_arquivos_pasta(STEP3)

    get_file('gerem_registros.xlsx', 'Gerem_Eventos', STEP1)
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

def separar_abas_e_salvar(caminho_arquivo):
    """
    Função para separar abas específicas de uma planilha e salvá-las como novos arquivos em uma pasta.

    Parâmetros:
    caminho_arquivo (str): Caminho completo para o arquivo Excel de entrada.

    Retorno:
    None
    """
    try:
        # Lê o arquivo Excel
        planilha = pd.ExcelFile(caminho_arquivo)
        
        # Lista de abas a serem processadas
        abas_interesse = ["gerem_eventos", "gerem_interacao", "gerem_malling"]
        
        # Caminho da pasta de saída
        pasta_saida = "step_2_stage_area"
        os.makedirs(pasta_saida, exist_ok=True)

        # Itera sobre as abas e salva cada uma como uma nova planilha
        for aba in abas_interesse:
            if aba in planilha.sheet_names:
                df = planilha.parse(aba)
                caminho_saida = os.path.join(pasta_saida, f"{aba}.xlsx")
                df.to_excel(caminho_saida, index=False)
                print(f"Aba '{aba}' salva em: {caminho_saida}")
            else:
                print(f"Aba '{aba}' não encontrada no arquivo.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")



    # Ler a planilha gerem_eventos
    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")
        return

    # Selecionar as colunas id_evento e programa_iniciativa
    df_selecionado = df[['id_evento', 'programa_iniciativa']].dropna()

    # Criar uma lista para armazenar os dados processados
    eventos_programa_iniciativa_expandidos = []

    # Separar os valores da coluna ue e criar novas linhas
    for _, row in df_selecionado.iterrows():
        id_evento = row['id_evento']
        ues = row['programa_iniciativa'].split(';')  # Separar ues por ';'
        for ue in ues:
            eventos_programa_iniciativa_expandidos.append({'id_evento': id_evento, 'programa_iniciativa': ue.strip()})

    # Criar o DataFrame resultante
    df_eventos_ues = pd.DataFrame(eventos_programa_iniciativa_expandidos)

    caminho_saida = os.path.join(ROOT, "step_3_data_processed", "eventos_programas_iniciativas.xlsx")

    # Salvar o DataFrame em Excel
    try:
        df_eventos_ues.to_excel(caminho_saida, index=False)
        print(f"Arquivo salvo com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

def gerar_eventos_1paraN(caminho_arquivo, nome_coluna, nome_arquivo_saida):
    """
    Gera uma nova planilha com os valores de uma coluna separados e transpostos em linhas.
    
    :param caminho_arquivo: Caminho para o arquivo Excel original.
    :param nome_coluna: Nome da coluna a ser processada.
    :param nome_arquivo_saida: Nome do arquivo de saída (sem extensão).
    """
    try:
        # Ler a planilha gerem_eventos
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")
        return

    # Verificar se a coluna existe no DataFrame
    if nome_coluna not in df.columns:
        print(f"Coluna '{nome_coluna}' não encontrada no arquivo.")
        return

    # Selecionar as colunas id_evento e a coluna especificada
    df_selecionado = df[['id_evento', nome_coluna]].dropna()

    # Criar uma lista para armazenar os dados processados
    eventos_expandidos = []

    # Separar os valores da coluna especificada e criar novas linhas
    for _, row in df_selecionado.iterrows():
        id_evento = row['id_evento']
        valores = row[nome_coluna].split(';')  # Separar os valores por ';'
        for valor in valores:
            eventos_expandidos.append({'id_evento': id_evento, nome_coluna: valor.strip()})

    # Criar o DataFrame resultante
    df_eventos = pd.DataFrame(eventos_expandidos)

    # Definir o caminho de saída
    ROOT = os.getcwd()  # Definir o diretório atual como ROOT
    caminho_saida = os.path.join(ROOT, "step_3_data_processed", f"{nome_arquivo_saida}.xlsx")

    # Garantir que o diretório de saída exista
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)

    # Salvar o DataFrame em Excel
    try:
        df_eventos.to_excel(caminho_saida, index=False)
        print(f"Arquivo salvo com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

def gerar_responsaveis_embrapii(eventos, interacoes, nome_arquivo_saida):
    """
    Processa os arquivos eventos e interações, unifica os dados, e gera uma planilha de responsabilidades.
    
    :param eventos: Caminho para o arquivo de eventos.
    :param interacoes: Caminho para o arquivo de interações.
    :param nome_arquivo_saida: Nome do arquivo de saída (sem extensão).
    """
    try:
        # Ler os arquivos
        df_eventos = pd.read_excel(eventos)
        df_interacoes = pd.read_excel(interacoes)

        # Selecionar colunas relevantes
        df_eventos = df_eventos[['id_evento', 'data_inicio', 'responsavel_embrapii']]
        df_interacoes = df_interacoes[['id_interacao', 'data', 'responsavel_embrapii']]

        # Renomear colunas para unificação
        df_eventos = df_eventos.rename(columns={
            'id_evento': 'id_evento_interacao',
            'data_inicio': 'data'
        })
        df_interacoes = df_interacoes.rename(columns={
            'id_interacao': 'id_evento_interacao'
        })

        # Adicionar a coluna tipo
        df_eventos['tipo'] = 'Evento'
        df_interacoes['tipo'] = 'Interação'

        # Unificar os DataFrames
        df_unificado = pd.concat([df_eventos, df_interacoes], ignore_index=True)

        # Criar a coluna id_responsabilidade com valores únicos
        df_unificado['id_responsabilidade'] = range(1, len(df_unificado) + 1)

        # Separar os valores da coluna responsavel_embrapii e transpor os dados
        responsabilidades = []
        for _, row in df_unificado.iterrows():
            responsaveis = row['responsavel_embrapii'].split(';') if pd.notna(row['responsavel_embrapii']) else []
            for responsavel in responsaveis:
                responsabilidades.append({
                    'id_responsabilidade': row['id_responsabilidade'],
                    'id_evento_interacao': row['id_evento_interacao'],
                    'data': row['data'],
                    'tipo': row['tipo'],
                    'responsavel_embrapii': responsavel.strip()
                })

        # Criar DataFrame final
        df_responsabilidades = pd.DataFrame(responsabilidades)

        # Definir o caminho de saída
        ROOT = os.getcwd()  # Diretório atual
        caminho_saida = os.path.join(ROOT, "step_3_data_processed", f"{nome_arquivo_saida}.xlsx")

        # Garantir que o diretório de saída exista
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)

        # Salvar o DataFrame em Excel
        df_responsabilidades.to_excel(caminho_saida, index=False)
        print(f"Arquivo salvo com sucesso em: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao processar os arquivos: {e}")


def main():

    #1 Baixar a planilha de registros
    puxar_planilhas()

    #2 Separar bases
    caminho_planilha = os.path.join(ROOT, "step_1_data_raw", "gerem_registros.xlsx")
    separar_abas_e_salvar(caminho_planilha)

    #3 Gerar as bases finais
    #3.1. Definir a origem das planilhas
    original_eventos = os.path.join(ROOT, "step_2_stage_area", "gerem_eventos.xlsx")
    original_interacoes = os.path.join(ROOT, "step_2_stage_area", "gerem_interacao.xlsx")
    original_malling = os.path.join(ROOT, "step_2_stage_area", "gerem_malling.xlsx")
    
    #3.2. Definir o destino das planilhas
    eventos = os.path.join(ROOT, "step_3_data_processed", "eventos.xlsx")
    interacoes = os.path.join(ROOT, "step_3_data_processed", "interacoes.xlsx")
    malling = os.path.join(ROOT, "step_3_data_processed", "malling.xlsx")

    #3.3. Copiar as planilhas
    shutil.copy(original_eventos, eventos)
    shutil.copy(original_interacoes, interacoes)
    shutil.copy(original_malling, malling)
    
    #3.4. Gerar novas planilhas
    gerar_eventos_1paraN(original_eventos, 'tema_evento', 'eventos_temas')
    gerar_eventos_1paraN(original_eventos, 'ues_participantes', 'eventos_unidades_embrapii')
    gerar_eventos_1paraN(original_eventos, 'programa_iniciativa', 'eventos_programas_iniciativas')

    #3.5. Gerar a planilha de responsáveis da Embrapii
    eventos = os.path.join(ROOT, "step_3_data_processed", "eventos.xlsx")
    interacoes = os.path.join(ROOT, "step_3_data_processed", "interacoes.xlsx")
    gerar_responsaveis_embrapii(eventos, interacoes, 'responsaveis_embrapii')

    #4. Levar arquivos para o Sharepoint
    upload_files(STEP3, "DWPII/gerem", SHAREPOINT_SITE, SHAREPOINT_SITE_NAME, SHAREPOINT_DOC)

    
if __name__ == "__main__":
    main()