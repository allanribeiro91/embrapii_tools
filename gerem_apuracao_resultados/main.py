import os
import datetime
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv
from office365_api.download_files import get_file
from office365_api.upload_files import upload_files
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

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

    get_file('apuracao_resultados_2024.xlsx', 'DWPII/gerem', STEP1)
    get_file('prospeccao_prospeccao.xlsx', 'DWPII/srinfo', STEP1)
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

def stage_area_apuracao_resultados(caminho_arquivo):
    """
    Função para selecionar apenas as colunas desejadas, ajustar a ordem, renomear colunas
    e realizar ajustes no DataFrame.

    Parâmetros:
    caminho_arquivo (str): Caminho completo para o arquivo Excel de entrada.

    Retorno:
    None
    """

    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)

        # Dicionário com as colunas de interesse, na ordem desejada, e seus novos nomes
        colunas_interesse_ordem_novoNome = {
            "Data": "data_interacao",
            "TIPO DE AÇÃO": "tipo_interacao",
            "Formato": "formato",
            "Descrição": "descricao",
            "Empresa": "empresa",
            "Responsável EMBRAPII": "responsavel_embrapii"
        }

        # Selecionar as colunas de interesse, reordenar e renomear
        df = df[list(colunas_interesse_ordem_novoNome.keys())]
        df = df.rename(columns=colunas_interesse_ordem_novoNome)

        # Adicionar a coluna 'tipo_acao' com valor "Interação GEREM" para todas as linhas
        df.insert(0, 'tipo_acao', 'Interação GEREM')

        # Remover registros duplicados considerando todas as colunas
        df = df.drop_duplicates()

        # Caminho da pasta de saída
        pasta_saida = "step_2_stage_area"
        os.makedirs(pasta_saida, exist_ok=True)  # Garantir que a pasta exista
        nome_planilha = "gerem_interacao"
        caminho_saida = os.path.join(pasta_saida, f"{nome_planilha}.xlsx")

        # Salvar o DataFrame resultante em um arquivo Excel
        df.to_excel(caminho_saida, index=False)
        print(f"Planilha '{nome_planilha}' salva em: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

def maior_menor_data(caminho_arquivo, nome_coluna):
    """
    Função para pegar a maior e menor data presente na planilha.

    Parâmetros:
    caminho_arquivo (str): Caminho completo para o arquivo Excel de entrada.
    nome_coluna (str): Nome da coluna que contém as datas.

    Retorno:
    tuple: Uma tupla contendo a menor e a maior data (data_menor, data_maior).
    """

    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)

        # Converter a coluna especificada para datetime
        df[nome_coluna] = pd.to_datetime(df[nome_coluna], errors='coerce')

        # Remover valores inválidos (NaT)
        df = df.dropna(subset=[nome_coluna])

        # Obter a menor e a maior data
        data_menor = df[nome_coluna].min()
        data_maior = df[nome_coluna].max()

        # Retornar as datas como uma tupla
        return data_menor, data_maior

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        return None

def stage_area_prospeccao(caminho_arquivo, data_inicio, data_fim):
    """
    Função para selecionar apenas as colunas desejadas e o período especificado.

    Parâmetros:
    caminho_arquivo (str): Caminho completo para o arquivo Excel de entrada.
    data_inicio (date): Data de recorte inicial.
    data_fim (date): Data de recorte final.

    Retorno:
    None
    """

    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)

        # Lista de colunas de interesse
        colunas_interesse_e_ordem = ["data_prospeccao", "unidade_embrapii", "cnpj_empresa", "nome_empresa", "pessoa_contatada_empresa_nome"]

        # Selecionar apenas as colunas de interesse
        df = df[colunas_interesse_e_ordem]

        # Converter a coluna 'data_prospeccao' para datetime, assumindo o formato 'dd/mm/yyyy'
        df['data_prospeccao'] = pd.to_datetime(df['data_prospeccao'], format='%d/%m/%Y', errors='coerce')

        # Filtrar as linhas pelo período especificado
        df = df[(df['data_prospeccao'] >= pd.to_datetime(data_inicio)) & 
                (df['data_prospeccao'] <= pd.to_datetime(data_fim))]

        # Adicionar a coluna 'tipo_acao' com valor "Prospecção" para todas as linhas
        df.insert(0, 'tipo_acao', 'Prospecção')

        # Remover registros duplicados considerando todas as colunas
        df = df.drop_duplicates()

        # Caminho da pasta de saída
        pasta_saida = "step_2_stage_area"
        nome_planilha = "srinfo_prospeccao"
        caminho_saida = os.path.join(pasta_saida, f"{nome_planilha}.xlsx")

        # Salvar o DataFrame resultante em um arquivo Excel
        df.to_excel(caminho_saida, index=False)
        print(f"Planilha '{nome_planilha}' salva em: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")


def apuracao_srinfo_prospeccao(path_gerem, path_prospeccao):
    """
    Função para apurar a presença de empresas que a GEREM interagiu na base de prospecções do SRInfo,
    priorizando o primeiro nome na comparação.

    Parâmetros:
    path_gerem (str): Caminho para a planilha GEREM.
    path_prospeccao (str): Caminho para a planilha de prospecções do SRInfo.

    Retorno:
    Cria planilhas que indicam empresas que tiveram interações com a GEREM e aparecem na base do SRInfo.
    """
    print('#Apuração de Prospecções')
    try:
        # Lê os arquivos Excel
        df_gerem = pd.read_excel(path_gerem)
        df_prospeccao = pd.read_excel(path_prospeccao)

        # Adicionar IDs únicos
        df_gerem['id_gerem'] = range(1, len(df_gerem) + 1)
        df_prospeccao['id_prospeccao'] = range(1, len(df_prospeccao) + 1)

        # Criar df_gerem_empresas com as colunas id_gerem e empresa, eliminando duplicatas
        df_gerem_empresas = df_gerem[['id_gerem', 'empresa']].drop_duplicates(subset=['empresa']).rename(columns={"empresa": "gerem_empresa"})

        # Função para obter o primeiro nome de uma empresa
        def obter_primeiro_nome(nome):
            return nome.split()[0] if isinstance(nome, str) else ""

        # Realizar comparação por verossimilhança
        comparacoes = []
        for _, row_gerem in df_gerem_empresas.iterrows():
            empresa_gerem = row_gerem['gerem_empresa']
            id_gerem = row_gerem['id_gerem']
            primeiro_nome_gerem = obter_primeiro_nome(empresa_gerem)

            for _, row_prospeccao in df_prospeccao.iterrows():
                nome_empresa_prospeccao = row_prospeccao['nome_empresa']
                id_prospeccao = row_prospeccao['id_prospeccao']
                primeiro_nome_prospeccao = obter_primeiro_nome(nome_empresa_prospeccao)

                # Comparação usando peso maior para o primeiro nome
                peso_primeiro_nome = 0.7  # Peso maior para o primeiro nome
                peso_nome_completo = 0.3  # Peso menor para o nome completo

                grau_primeiro_nome = fuzz.token_sort_ratio(primeiro_nome_gerem, primeiro_nome_prospeccao)
                grau_nome_completo = fuzz.token_sort_ratio(empresa_gerem, nome_empresa_prospeccao)
                grau_final = (peso_primeiro_nome * grau_primeiro_nome) + (peso_nome_completo * grau_nome_completo)

                if grau_final > 50:  # Considerar apenas comparações acima de 50
                    comparacoes.append({
                        'id_gerem': id_gerem,
                        'gerem_empresa': empresa_gerem,
                        'id_prospeccao': id_prospeccao,
                        'prospeccao_nome_empresa': nome_empresa_prospeccao,
                        'grau_verossimilhanca': round(grau_final)
                    })

        # Criar DataFrame com os resultados da comparação
        df_comparacao = pd.DataFrame(comparacoes)

        # Exportar os dados para Excel
        pasta_saida = "step_3_data_processed"
        os.makedirs(pasta_saida, exist_ok=True)  # Garantir que a pasta exista

        # Exportar DataFrames
        df_gerem.to_excel(os.path.join(pasta_saida, "gerem_com_ids.xlsx"), index=False)
        df_prospeccao.to_excel(os.path.join(pasta_saida, "prospeccao_com_ids.xlsx"), index=False)
        df_comparacao.to_excel(os.path.join(pasta_saida, "comparacao_gerem_prospeccao.xlsx"), index=False)

        print("Processamento concluído. Arquivos exportados com sucesso.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")



def main():
    # Registrar o início do processo
    data_hora_inicio = datetime.now()
    print(f"Processo iniciado em: {data_hora_inicio.strftime('%Y-%m-%d %H:%M:%S')}")


    # 1 Baixar a planilha de registros
    puxar_planilhas()

    # 2. Tratar as bases de dados
    # 2.1. Gerem > Interações
    caminho_raw_gerem_interacoes = os.path.join(ROOT, "step_1_data_raw", "apuracao_resultados_2024.xlsx")
    stage_area_apuracao_resultados(caminho_raw_gerem_interacoes)

    # 2.2. SRInfo > Prospecção
    gerem_interacao_caminho = os.path.join(ROOT, "step_2_stage_area", "gerem_interacao.xlsx")
    data_menor, data_maior = maior_menor_data(gerem_interacao_caminho, "data_interacao")
    caminho_raw_prospeccao = os.path.join(ROOT, "step_1_data_raw", "prospeccao_prospeccao.xlsx")
    stage_area_prospeccao(caminho_raw_prospeccao, data_menor, data_maior)

    # 3. Apurar resultados
    srinfo_prospeccao = os.path.join(ROOT, "step_2_stage_area", "srinfo_prospeccao.xlsx")
    apuracao_srinfo_prospeccao(gerem_interacao_caminho, srinfo_prospeccao)

    # Registrar o término do processo
    data_hora_fim = datetime.now()
    print(f"Processo finalizado em: {data_hora_fim.strftime('%Y-%m-%d %H:%M:%S')}")

    # Exibir o tempo total de execução
    tempo_total = data_hora_fim - data_hora_inicio
    print(f"Tempo total de execução: {tempo_total}")

    print('Finalizado!')

  
if __name__ == "__main__":
    main()