import os
import sys
from dotenv import load_dotenv
import pandas as pd
from datetime import datetime

# carregar .env 
load_dotenv()

def definir_recorte():
    # define datas de início e fim do recorte
    data_inicio = pd.to_datetime('2023-11-16')
    while True:
        try: 
            data_fim = pd.to_datetime(
                input('[BRASIL MAIS PRODUTIVO] Insira a data de fim do recorte que deseja\n(OBS: no formato AAAA-MM-DD): '),
                format='%Y-%m-%d'
                )
            break
        except ValueError:
            print('A data está no formato errado ou o dia não existe.\n')
    return [data_inicio,data_fim]

# recebe `recorte` que é uma lista de data inicio e data fim
def gerar_projetos_bmaisp(recorte):

    # carregar planilha
    projetos = pd.read_excel('inputs\\portfolio.xlsx')

    # converte coluna de data em tipo de data
    projetos['data_contrato'] = pd.to_datetime(projetos['data_contrato'])

    # aplica filtros de data
    projetos_filtro = projetos[
        (projetos['data_contrato'] > recorte[0]) & (projetos['data_contrato'] < recorte[1])
    ]

    # aplica outros filtros
    projetos_filtro = projetos_filtro[
        (projetos_filtro['brasil_mais_produtivo'].isin(['Sim']))
    ]
    
    #aplica filtro de status
    projetos_filtro = projetos_filtro[
        ~projetos_filtro['status'].isin(['Cancelado', 'Suspenso'])
    ]

    # transforma todos os status para "Em andamento"
    projetos_filtro['status'] = 'Em andamento'

    # projetos_filtro = projetos_filtro[[
    #     "codigo_projeto", "unidade_embrapii", "data_contrato", "data_inicio", "data_termino",
    #     "status", "tipo_projeto", "parceria_programa", "uso_recurso_obrigatorio", "tecnologia_habilitadora",
    #     "brasil_mais_produtivo", "area_aplicacao", "projeto", "trl_inicial", "trl_final", "valor_embrapii", "valor_empresa",
    #     "valor_unidade_embrapii", "titulo", "titulo_publico", "objetivo", "descricao_publica", "data_extracao_dados"]]
    projetos_filtro = projetos_filtro[[
        "codigo_projeto", "data_contrato", "status", "tecnologia_habilitadora",
        "valor_embrapii", "valor_empresa", "valor_sebrae", "valor_unidade_embrapii", "data_extracao_dados"]]
    
    # retorno da função
    return projetos_filtro

def gerar_projetos_empresas_bmaisp(projetos_filtro):

    # carregar planilha
    projetos_empresas = pd.read_excel('inputs\\projetos_empresas.xlsx')

    # filtra para somente empresas que estão em projetos NIB
    projetos_empresas_filtro = projetos_empresas[
        projetos_empresas['codigo_projeto'].isin(projetos_filtro['codigo_projeto'])
        ]
    
    # retorno da função
    return projetos_empresas_filtro

def gerar_empresas_bmaisp(projetos_empresas_filtro):

    # carregar planilha
    empresas = pd.read_excel('inputs\\informacoes_empresas.xlsx')

    # filtra para somente empresas que estão em projetos NIB
    empresas_filtro = empresas[
        empresas['cnpj'].isin(projetos_empresas_filtro['cnpj'])
        ]
    
    # outros filtros
    empresas_filtro = empresas_filtro.drop(columns=['novo'])

    # carregar planilha
    cnae_ibge = pd.read_excel('inputs\\cnae_ibge.xlsx')

    # adicionar colunas nome_secao e nome_divisao ao empresas_filtro usando merge
    empresas_filtro = empresas_filtro.merge(
        cnae_ibge[['subclasse2', 'nome_secao', 'nome_divisao']],
        left_on='cnae_subclasse',
        right_on='subclasse2',
        how='left'
    )
    
    # remover a coluna 'subclasse2' se necessário
    empresas_filtro = empresas_filtro.drop(columns=['subclasse2'])

        # aplica filtro para remover empresas com porte "Grande"
    empresas_filtro = empresas_filtro[
        ~empresas_filtro['porte'].isin(['Grande'])
    ]

    # retorno da função
    return empresas_filtro
    
def gerar_planilha_bmaisp():
    # define recorte e chama funções
    recorte = definir_recorte()
    projetos_filtro = gerar_projetos_bmaisp(recorte)
    projetos_empresas_filtro = gerar_projetos_empresas_bmaisp(projetos_filtro)
    empresas_filtro = gerar_empresas_bmaisp(projetos_empresas_filtro)

    # exporta excel
    today = datetime.today()
    destino_arquivo = f'outputs\\embrapii_portfolio_brasilmaisprodutivo_{recorte[1].strftime('%Y.%m.%d')}_gerado_em_{today.strftime('%Y.%m.%d')}.xlsx'
    with pd.ExcelWriter(destino_arquivo, engine='openpyxl') as writer:
        projetos_filtro.to_excel(writer, sheet_name='portfolio_projetos', index=False)  
        projetos_empresas_filtro.to_excel(writer, sheet_name='projetos_empresas', index=False)
        empresas_filtro.to_excel(writer, sheet_name='dados_empresas', index=False)

    # reporta sucesso
    print(f'Planilha filtrada do B+P gerada com sucesso.')
