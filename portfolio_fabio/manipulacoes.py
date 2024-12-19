import os
from datetime import datetime
import pandas as pd
import numpy as np
from unidecode import unidecode
from dotenv import load_dotenv

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

INPUTS = os.path.abspath(os.path.join(ROOT, 'inputs'))

from processar_excel import processar_excel

def juntar_planilhas():
    # lendo as planilhas
    port = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'portfolio.xlsx')))
    proj_emp = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'projetos_empresas.xlsx')))
    emp = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'informacoes_empresas.xlsx')))
    emp['municipio'] = emp['municipio'].astype(str)
    ues = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'info_unidades_embrapii.xlsx')))
    ues['assinatura_plano_acao'] = pd.to_datetime(ues['assinatura_plano_acao'], format = '%d/%m/%Y', errors = 'coerce')
    ues['ano_credenciamento'] = ues['assinatura_plano_acao'].dt.year
    territorial = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'ibge_municipios.xlsx')))
    cnae = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'cnae_ibge.xlsx')))

    # remover acentos das colunas de municipio
    def remove_acentos(text):
        return unidecode(text)
    
    ues['municipio'] = ues['municipio'].apply(remove_acentos).str.upper()
    ues['municipio'] = np.where(ues['municipio'].str.contains('BARRA DA TIJUCA|BOTAFOGO'), 'RIO DE JANEIRO', ues['municipio'])
    emp['municipio'] = emp['municipio'].apply(remove_acentos).str.upper()
    emp['municipio'] = np.where(emp['municipio'].str.contains('MOGI-GUACU'), 'MOGI GUACU', emp['municipio'])
    emp['municipio'] = np.where(emp['municipio'].str.contains('FORTALEZA DO TABOCAO'), 'TABOCAO', emp['municipio'])
    territorial['no_municipio'] = territorial['no_municipio'].apply(remove_acentos).str.upper()

    #juntando as planilhas
    port_emp = pd.merge(port, proj_emp, on = 'codigo_projeto', how = 'right')
    emp_municipio = pd.merge(emp, territorial, left_on = ['municipio','uf'], right_on = ['no_municipio','sg_uf'], how = 'left')
    emp_cnae = pd.merge(emp_municipio, cnae, left_on = 'cnae_subclasse', right_on = 'subclasse2', how = 'left')
    emp2 = pd.merge(port_emp, emp_cnae, on = 'cnpj', how = 'left')
    ues2 = pd.merge(ues, territorial, left_on = ['municipio','uf'], right_on = ['no_municipio','sg_uf'], how = 'left')
    merged = pd.merge(emp2, ues2, on = 'unidade_embrapii', how = 'left')

    return [merged, emp_cnae]


def ajustes(merged):

    # lendo a planilha pedidos de pi e a planilha geral
    ppi = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'pedidos_pi.xlsx')))
    planilha_geral = merged

    # fazendo as transformacoes necessarias
    planilha_geral['valor_total'] = planilha_geral['valor_embrapii'] + planilha_geral['valor_empresa'] + planilha_geral['valor_unidade_embrapii'] + planilha_geral['valor_sebrae']
    contagem_linhas = ppi.groupby('codigo_projeto').size().reset_index(name = 'pedidos_pi')

    merged2 = pd.merge(planilha_geral, contagem_linhas, on = 'codigo_projeto', how = 'left')

    return merged2

def combinar_dados(merged2):
     # concatenando os valores de empresas, para ter somente uma linha para cada codigo
    def concat_values(series):
         return '; '.join(series.astype(str))
    
    # agrupando o DataFrame pela coluna 'codigo_projeto'
    combinado = merged2.groupby('codigo_projeto').agg({
    'unidade_embrapii': 'first',
    'ano_credenciamento': 'first',
    'tipo_instituicao': 'first',
    'uf_y': 'first',
    'regiao_pais_y': 'first',
    'competencias_tecnicas': 'first',
    'empresa': concat_values,
    'cnpj': concat_values,
    'porte': concat_values,
    'uf_x': concat_values,
    'regiao_pais_x': concat_values,
    'agrupamento': concat_values,
    'divisao': concat_values,
    'nome_divisao': concat_values,
    'cnae_subclasse': concat_values,
    'nome_subclasse': concat_values,
    'tecnologia_habilitadora': 'first',
    'area_aplicacao': 'first',
    'missoes_cndi': 'first',
    'codigo_negociacao': 'first',
    'projeto': 'first',
    'titulo_publico': 'first',
    'objetivo': 'first',
    'descricao_publica': 'first',
    'tipo_projeto': 'first',
    'modalidade_financiamento': 'first',
    'parceria_programa': 'first',
    'call': 'first',
    'cooperacao_internacional': 'first',
    'macroentregas': 'first',
    'pct_aceites': 'first',
    'status': 'first',
    'data_contrato': 'first',
    'data_inicio': 'first',
    'data_termino': 'first',
    'uso_recurso_obrigatorio': 'first',
    'trl_inicial': 'first',
    'trl_final': 'first',
    'valor_total': 'first',
    'valor_embrapii': 'first',
    'valor_empresa': 'first',
    'valor_sebrae': 'first',
    'valor_unidade_embrapii': 'first',
    'pedidos_pi': 'first',
    'data_extracao_dados': 'first'
    }).reset_index()

    # Adicione uma coluna para a contagem de repetições
    combinado['numero_empresas_no_projeto'] = merged2.groupby('codigo_projeto').size().values


    # contando duracao em meses
    combinado['data_contrato'] = pd.to_datetime(combinado['data_contrato'], format = '%d/%m/%Y', errors = 'coerce')
    combinado['data_inicio'] = pd.to_datetime(combinado['data_inicio'], format = '%d/%m/%Y', errors = 'coerce')
    combinado['data_termino'] = pd.to_datetime(combinado['data_termino'], format = '%d/%m/%Y', errors = 'coerce')
    combinado['duracao_meses'] = ((combinado['data_termino'].dt.year - combinado['data_contrato'].dt.year) * 12 +
                                  (combinado['data_termino'].dt.month - combinado['data_contrato'].dt.month))
    
    # renomeando 'Encerrado' para 'Concluído'
    combinado['status'] = combinado['status'].apply(lambda x: 'Concluído' if x == 'Encerrado' else x)
        
    # mesclando com a contagem de linhas
    combinado['pedidos_pi'] = combinado['pedidos_pi'].fillna(0)
    
    # incluindo ipca
    ipca = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'ipca.xlsx')))

    # Definir a primeira linha como o título
    ipca.columns = ipca.iloc[0]

    # Remover a primeira linha agora que ela foi usada como título
    ipca = ipca.drop(0).reset_index(drop=True)

    ipca['Valor'] = ipca['Valor'].astype(float)

    def get_previous_month(date):
        if date.month == 1:  # Se for janeiro, pegar dezembro do ano anterior
            return f"{date.year - 1}12"
        else:
            previous_month = date - pd.DateOffset(months=1)
            return previous_month.strftime('%Y%m')

    # pegando mês anterior de contratação para cálculo do valor atualizado pelo IPCA
    combinado['mes_anterior'] = combinado['data_contrato'].apply(get_previous_month)
    ultimo_mes = ipca['Mês (Código)'].iloc[-1]
    ultimo_valor = ipca['Valor'].iloc[-1]

    # juntando os valores IPCA do mês anterior com a planilha dos projetos
    combinado = pd.merge(combinado, ipca[['Mês (Código)', 'Valor']], left_on = 'mes_anterior', right_on = 'Mês (Código)', how = 'left')
    combinado['Valor'] = combinado['Valor'].fillna(ultimo_valor)

    # calculando os valores atualizados
    combinado['valor_total_ipca'] = round((ultimo_valor/combinado['Valor'])*combinado['valor_total'],2)
    combinado['valor_embrapii_ipca'] = round((ultimo_valor/combinado['Valor'])*combinado['valor_embrapii'],2)
    combinado['valor_empresa_ipca'] = round((ultimo_valor/combinado['Valor'])*combinado['valor_empresa'],2)
    combinado['valor_sebrae_ipca'] = round((ultimo_valor/combinado['Valor'])*combinado['valor_sebrae'],2)
    combinado['valor_unidade_ipca'] = round((ultimo_valor/combinado['Valor'])*combinado['valor_unidade_embrapii'],2)

    # ordenando das mais novas para as mais antigas
    combinado = combinado.sort_values(by='data_contrato', ascending=False)

    # salvando o dataframe em excel
    combinado.to_excel(os.path.abspath(os.path.join(INPUTS, 'planilha_combinada.xlsx')), index = False)

    # pegando a data de extração dos dados
    extracao = min(combinado['data_extracao_dados']).strftime(format='%Y%m%d')

    # retornando a data do valor de referência utilizado para o cálculo do IPCA e a data de extração dos dados
    return ultimo_mes, extracao



def processar_dados(ultimo_mes, extracao):
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'inputs')
    nome_arquivo = "planilha_combinada.xlsx"
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(origem, f'portfolio_atualizado_{extracao}.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'codigo_projeto',
        'unidade_embrapii',
        'ano_credenciamento',
        'tipo_instituicao',
        'uf_y',
        'regiao_pais_y',
        'competencias_tecnicas',
        'empresa',
        'cnpj',
        'porte',
        'uf_x',
        'regiao_pais_x',
        'agrupamento',
        'divisao',
        'nome_divisao',
        'cnae_subclasse',
        'nome_subclasse',
        'numero_empresas_no_projeto',
        'tecnologia_habilitadora',
        'area_aplicacao',
        'missoes_cndi',
        'codigo_negociacao',
        'projeto',
        'titulo_publico',
        'objetivo',
        'descricao_publica',
        'tipo_projeto',
        'modalidade_financiamento',
        'parceria_programa',
        'call',
        'cooperacao_internacional',
        'macroentregas',
        'pct_aceites',
        'status',
        'data_contrato',
        'data_inicio',
        'data_termino',
        'duracao_meses',
        'uso_recurso_obrigatorio',
        'trl_inicial',
        'trl_final',
        'valor_total',
        'valor_embrapii',
        'valor_empresa',
        'valor_sebrae',
        'valor_unidade_embrapii',
        'pedidos_pi',
        'valor_total_ipca',
        'valor_embrapii_ipca',
        'valor_empresa_ipca',
        'valor_sebrae_ipca',
        'valor_unidade_ipca',
    ]

    novos_nomes_e_ordem = {
        'codigo_projeto': 'Código',
        'unidade_embrapii': 'Unidade EMBRAPII',
        'ano_credenciamento': 'Ano de Credenciamento',
        'tipo_instituicao': 'Tipo de Instituição',
        'uf_y': 'UF UE',
        'regiao_pais_y': 'Região UE',
        'empresa': 'Empresas',
        'cnpj': 'CNPJ',
        'porte': 'Porte da Empresa',
        'numero_empresas_no_projeto': 'Número de Empresas no Projeto',
        'uf_x': 'UF da Empresa',
        'regiao_pais_x': 'Região da Empresa',
        'agrupamento': 'Agrupamento Div CNAE',
        'divisao': 'CNAE Divisão',
        'nome_divisao': 'Nomenclatura CNAE Divisão',
        'cnae_subclasse': 'CNAE Classe',
        'nome_subclasse': 'Nomenclatura CNAE Classe',
        'competencias_tecnicas': 'Competência UE',
        'tecnologia_habilitadora': 'Tecnologias Habilitadoras',
        'area_aplicacao': 'Áreas de Aplicação',
        'missoes_cndi': 'Missões - CNDI final',
        'codigo_negociacao': 'Negociações',
        'projeto': 'Projeto',
        'titulo_publico': 'Título público',
        'objetivo': 'Objetivo',
        'descricao_publica': 'Descrição pública',
        'tipo_projeto': 'Tipo de projeto',
        'modalidade_financiamento': 'Modalidade de financiamento',
        'parceria_programa': 'Parceria / Programa',
        'call': 'Call',
        'cooperacao_internacional': 'Cooperação Internacional',
        'macroentregas': 'Macroentregas',
        'pct_aceites': '% de Aceites',
        'status': 'Status',
        'data_contrato': 'Data do contrato',
        'data_inicio': 'Data de início',
        'data_termino': 'Data de término',
        'duracao_meses': 'Tempo de duração (meses)',
        'uso_recurso_obrigatorio': 'É usada obrigatoriedade?',
        'trl_inicial': 'Nível de maturidade inicial',
        'trl_final': 'Nível de maturidade final',
        'valor_total': 'Valor total',
        'valor_embrapii': 'Valor aportado EMBRAPII',
        'valor_empresa': 'Valor aportado Empresa',
        'valor_sebrae': 'Valor aportado Sebrae',
        'valor_unidade_embrapii': 'Valor aportado Unidade',
        'pedidos_pi': 'Pedidos de Propriedade Intelectual',
        'valor_total_ipca': f'Valor total IPCA {ultimo_mes}',
        'valor_embrapii_ipca': f'Valor Embrapii IPCA {ultimo_mes}',
        'valor_empresa_ipca': f'Valor Empresa IPCA {ultimo_mes}',
        'valor_sebrae_ipca': f'Valor Sebrae IPCA {ultimo_mes}',
        'valor_unidade_ipca': f'Valor Unidade IPCA {ultimo_mes}',
    }

    # Campos de data e valor
    campos_data = ['Data do contrato', 'Data de início', 'Data de término']
    campos_valor = ['Valor total', 'Valor aportado EMBRAPII', 'Valor aportado Empresa', 'Valor aportado Sebrae', 'Valor aportado Unidade',
                    f'Valor total IPCA {ultimo_mes}', f'Valor Embrapii IPCA {ultimo_mes}', f'Valor Empresa IPCA {ultimo_mes}', f'Valor Sebrae IPCA {ultimo_mes}',
                    f'Valor Unidade IPCA {ultimo_mes}']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor, ordenar=True, campo_ordenar='data_contrato', crescente=False)


def empresas(emp, extracao):
    proj_emp = pd.read_excel(os.path.abspath(os.path.join(INPUTS, 'projetos_empresas.xlsx')))

    # concatenando os valores de empresas, para ter somente uma linha para cada cnpj
    def concat_values(series):
         return '; '.join(series.astype(str))
    
    # agrupando o DataFrame pela coluna 'cnpj'
    projetos = proj_emp.groupby('cnpj').agg({
    'codigo_projeto': concat_values,
    }).reset_index()

    projetos['numero_projetos'] = proj_emp.groupby('cnpj').size().values

    empresas = pd.merge(emp, projetos, on = 'cnpj', how = 'left')

    empresas.to_excel('portfolio_fabio/inputs/empresas.xlsx')

    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'inputs')
    nome_arquivo = "empresas.xlsx"
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(origem, f'informacoes_empresas_{extracao}.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'cnpj',
        'empresa',
        'regiao_pais',
        'uf',
        'municipio',
        'cod_uf',
        'cod_municipio_gaia',
        'porte',
        'faixa_faturamento',
        'divisao',
        'nome_divisao',
        'cnae_subclasse',
        'nome_subclasse',
        'agrupamento',
        'numero_projetos',
        'codigo_projeto',
    ]

    novos_nomes_e_ordem = {
        'cnpj': 'CNPJ',
        'empresa': 'Empresa',
        'regiao_pais': 'Região',
        'uf': 'UF',
        'municipio': 'Município',
        'cod_uf': 'Código IBGE UF',
        'cod_municipio_gaia': 'Código IBGE Município',
        'porte': 'Porte',
        'faixa_faturamento': 'Faixa de Faturamento',
        'divisao': 'CNAE Divisão',
        'nome_divisao': 'Nomenclatura CNAE Divisão',
        'cnae_subclasse': 'CNAE Classe',
        'nome_subclasse': 'Nomenclatura CNAE Classe',
        'agrupamento': 'Agrupamento div CNAE',
        'numero_projetos': 'Número de projetos',
        'codigo_projeto': 'Projetos',
    }

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino)


def gerar_planilha_unica(extracao):

    # lendo os arquivos
    portfolio = pd.read_excel(os.path.abspath(os.path.join(INPUTS, f'portfolio_atualizado_{extracao}.xlsx')))
    empresas = pd.read_excel(os.path.abspath(os.path.join(INPUTS, f'informacoes_empresas_{extracao}.xlsx')))

    # juntando na mesma planilha
    with pd.ExcelWriter(f'portfolio_fabio/up/Portfolio Trabalho {extracao}.xlsx') as writer:
        portfolio.to_excel(writer, sheet_name='Portfolio Trabalho', index=False)
        empresas.to_excel(writer, sheet_name='Informações Empresas', index=False)
