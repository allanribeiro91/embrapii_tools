import os
import pandas as pd
from datetime import datetime
import inspect

def processar_csv(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino,
                  campos_data= None, campos_valor = None, mes_ano = None):
    """
    Função para processar um arquivo CSV, selecionando apenas as colunas de interesse,
    renomeando-as e reordenando-as conforme especificado.
        arquivo_origem: caminho completo do arquivo CSV a ser processado
        campos_interesse: lista com os nomes das colunas a serem mantidas
        novos_nomes_e_ordem: dicionário com os novos nomes das colunas e a ordem desejada
        arquivo_destino: caminho completo do arquivo Excel a ser gerado
        campos_data: lista com os nomes das colunas que representam datas
        campos_valor: lista com os nomes das colunas que representam valores
        mes_ano: lista com os nomes das colunas que representam mês e ano
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
    
        # Ler o arquivo Excel
        df = pd.read_csv(f'{arquivo_origem}.csv', delimiter=',')

        # Selecionar apenas as colunas de interesse
        df_selecionado = df[campos_interesse]

        # Renomear as colunas e definir a nova ordem
        df_renomeado = df_selecionado.rename(columns=novos_nomes_e_ordem)

        # Ajustar campos de data, se fornecidos
        if campos_data:
            for campo in campos_data:
                if campo in df_renomeado.columns:
                    df_renomeado[campo] = pd.to_datetime(df_renomeado[campo], format='%Y-%m-%d', errors='coerce')

        if mes_ano:
            for campo in mes_ano:
                df_renomeado[campo] = pd.to_datetime(df[campo])
                df_renomeado['ano'] = df_renomeado[campo].dt.year
                df_renomeado['mes'] = df_renomeado[campo].dt.month

        if campos_valor:
            for campo in campos_valor:
                df_renomeado[campo] = pd.to_numeric(df_renomeado[campo], errors='coerce').fillna(0)


        # Reordenar as colunas conforme especificado
        df_final = df_renomeado[list(novos_nomes_e_ordem.values())]

        today = datetime.now()
        df_final['data_extracao'] = today

        # Garantir que o diretório de destino existe
        os.makedirs(os.path.dirname(arquivo_destino), exist_ok=True)

        # Verificar se o arquivo de destino está sendo usado e remover se necessário
        if os.path.exists(arquivo_destino):
            os.remove(arquivo_destino)

        # Salvar o arquivo resultante
        df_final.to_excel(f'{arquivo_destino}.xlsx', index=False, sheet_name='Sheet1')

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")

