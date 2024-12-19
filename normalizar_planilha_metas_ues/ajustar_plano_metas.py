import pandas as pd

PLANILHA = 'Metas e Planos Financeiros UEs BKP.xlsx'
NOME_ARQUIVO_SAIDA = 'ues_planos_metas.xlsx'
NOME_ABA = 'Plano de Metas'

def ajustar_plano_metas():
    # Ler a planilha original
    plano_metas = pd.read_excel(PLANILHA, sheet_name=NOME_ABA)

    # Ajustar nomes das colunas
    plano_metas = plano_metas.rename(columns={
        'Unidade EMBRAPII': 'unidade_embrapii',
        'Título da meta': 'titulo_meta',
        'Responsável pela ADIÇÃO dos dados': 'responsavel_adicao_dados',
        'Responsável pela UE SRINFO': 'responsavel_ue_srinfo',
        'Número do TA': 'ta_numero',
        'Número do Apostilamento': 'apostilamento_numero',
        'Observações e comentários': 'observacoes'
    })

    # Converter todas as colunas de ano para string
    anos = [str(ano) for ano in range(2014, 2031)]
    plano_metas.columns = plano_metas.columns.astype(str)

    # Transpor colunas de ano em linhas
    anos = [str(ano) for ano in range(2014, 2031)]  # Garantindo que os anos sejam strings
    plano_metas_long = plano_metas.melt(
        id_vars=['unidade_embrapii', 'titulo_meta', 'responsavel_adicao_dados', 
                 'responsavel_ue_srinfo', 'ta_numero', 'apostilamento_numero', 'observacoes'],
        value_vars=anos,
        var_name='ano',
        value_name='valor'
    )

    # Criar coluna 'titulo_meta_ajustada' com base nos valores de 'titulo_meta'
    ajuste_titulo_meta = {
        'Projetos contratados': 'Nº de projetos contratados',
        'Contratação de projetos': 'Nº de projetos contratados',
        'Contratação de empresas': 'Nº de empresas que contrataram',
        'Empresas contratantes': 'Nº de empresas que contrataram',
        'Empresas prospectadas': 'Nº de empresas prospectadas',
        'Eventos com empresas': 'Nº de eventos com empresas',
        'Geração de novos produtos e processos': 'Geração de novos produtos e processos',
        'Geração de propriedade intelectual': 'Geração de propriedade intelectual',
        'Grau de satisfação das empresas': 'Nota de satisfação das empresas',
        'Inserção de recursos humanos em projetos de PD&I': 'Inserção de recursos humanos em projetos de PD&I',
        'Número de empresas contratadas': 'Nº de empresas que contrataram',
        'Número de empresas contratantes': 'Nº de empresas que contrataram',
        'Número de empresas prospectadas': 'Nº de empresas prospectadas',
        'Número de empresas técnicas': 'Nº de empresas prospectadas',
        'Número de projetos a contratar': 'Nº de projetos contratados',
        'Número de propostas técnicas': 'Nº de propostas técnicas',
        'Participação de alunos em projetos de PD&I': 'Participação de alunos em projetos de PD&I',
        'Participação de alunos(as) em projetos de PD&I': 'Participação de alunos em projetos de PD&I',
        'Participação de empresas em eventos': 'Participação de empresas em eventos',
        'Participação de empresas novas na carteira': 'Participação de empresas novas na carteira',
        'Participação de projetos de alta tecnologia em carteira': 'Participação de projetos de alta tecnologia em carteira',
        'Participação financeira das empresas no portfólio': 'Percentual de participação financeira das empresas nos projetos contratados',
        'Participação financeira das empresas nos projetos contratados': 'Percentual de participação financeira das empresas nos projetos contratados',
        'Pedidos de Propriedade intelectual': 'Nº de pedidos de PI',
        'Propostas técnicas': 'Nº de propostas técnicas',
        'Prospecção de empresas': 'Nº de empresas prospectadas',
        'Prospecção de empresas em eventos': 'Prospecção de empresas em eventos',
        'Prospecção de emrpresas': 'Nº de empresas prospectadas',
        'Satisfação das Empresas': 'Nota de satisfação das empresas',
        'Startups, micro e pequenas empresas contratantes': 'Nº de empresas startups, micro e pequenas contratantes',
        'Taxa de cumprimento dos prazos de execução': 'Taxa de cumprimento dos prazos de execução',
        'Taxa de licenciamento de tecnologias': 'Taxa de licenciamento de tecnologias',
        'Taxa de sucesso de projeto': 'Percentual de sucesso de projeto',
        'Taxa de sucesso de projetos': 'Percentual de sucesso de projeto',
        'Taxa de sucesso de propostas técnicas': 'Percentual de sucesso de propostas técnicas',
        'Tempo de retorno dos investimentos': 'Tempo de retorno dos investimentos'
    }
    plano_metas_long['titulo_meta_ajustada'] = plano_metas_long['titulo_meta'].map(ajuste_titulo_meta).fillna(plano_metas_long['titulo_meta'])

    # Remover linhas onde unidade_embrapii ou valor estão vazios
    plano_metas_long = plano_metas_long.dropna(subset=['unidade_embrapii', 'valor'])
    
    #Criar a coluna _considerar
        #1º passo: criar uma cópia do dataframe -> df_suporte
            #levar apenas os campos unidade_embrapii, titulo_meta_ajustada, ta_numero, ano
            #antes de levar, eu quero que sejam eliminadas as células de valor igual a zero
        #2º passo: ordenar ta_numero do maior para o menor e, depois, ordenar o ano do menor para o maior
        #3º passo: remover as duplicadas de unidade_embrapii, titulo_meta_ajustada e ano
            #Objetivo é manter apenas a linha com o maior TA, quando houver repetição
        
        #4º passo: fazer uma espécie de 'de para' entre o plano_metas_long e o df_suporte
            #você deve pesquisar a linha do plano_metas_long no df_suporte, com base em unidade_embrapii, titulo_meta_ajustada, ta_numero e ano
            #quando os valores se repetirem, o _considerar é 'Sim', quando não é 'Não'


    # Criar df_suporte aplicando o filtro de valor diferente de zero
    df_suporte = plano_metas_long[plano_metas_long['valor'] != 0][['unidade_embrapii', 'titulo_meta_ajustada', 'ta_numero', 'ano']].copy()



    # Ordenar ta_numero (decrescente) e ano (crescente)
    df_suporte = df_suporte.sort_values(by=['unidade_embrapii', 'titulo_meta_ajustada', 'ano', 'ta_numero'], 
                                        ascending=[True, True, True, False])

    # Remover duplicatas para manter o maior TA por unidade_embrapii, titulo_meta_ajustada e ano
    df_suporte = df_suporte.drop_duplicates(subset=['unidade_embrapii', 'titulo_meta_ajustada', 'ano'], keep='first')

    # Adicionar coluna _considerar no plano_metas_long com base no df_suporte
    plano_metas_long = plano_metas_long.merge(df_suporte.assign(_considerar='Sim'), 
                                              on=['unidade_embrapii', 'titulo_meta_ajustada', 'ta_numero', 'ano'], 
                                              how='left')

    # Preencher NaN em _considerar como 'Não' para as linhas que não estão no df_suporte
    plano_metas_long['_considerar'] = plano_metas_long['_considerar'].fillna('Não')

    # Reordenar colunas conforme solicitado
    plano_metas_long = plano_metas_long[[
        'unidade_embrapii', 'titulo_meta', 'titulo_meta_ajustada', 
        'responsavel_adicao_dados', 'responsavel_ue_srinfo', 'ta_numero', 
        'apostilamento_numero', 'observacoes', 'ano', 'valor', '_considerar'
    ]]

    # Salvar o resultado em um novo arquivo Excel
    plano_metas_long.to_excel(NOME_ARQUIVO_SAIDA, index=False)


