def negociacao_comparacao():
    """
    Função para apurar a presença de empresas que a GEREM interagiu na base de negociações do SRInfo,
    utilizando "nome_capital" com maior peso e filtrando por data_interacao e data_prospeccao.

    Retorno:
    Cria planilhas que indicam empresas que tiveram interações com a GEREM e aparecem na base do SRInfo.
    """
    try:
        # Lê os arquivos Excel
        df_gerem = pd.read_excel(PATH_GEREM_INTERACAO)
        df_negociacao = pd.read_excel(PATH_NEGOCIACOES_EMPRESAS_NOME)

        # Criar df_gerem_empresas com as colunas id_gerem, empresa, empresa_nome_capital e data_interacao
        df_gerem_empresas = df_gerem[['id_gerem', 'empresa', 'empresa_nome_capital', 'data_interacao']].copy()

        # Capitalizar os valores em "empresa", "empresa_nome_capital" e "nome_empresa"
        df_gerem_empresas['empresa'] = df_gerem_empresas['empresa'].str.upper()
        df_gerem_empresas['empresa_nome_capital'] = df_gerem_empresas['empresa_nome_capital'].str.upper()
        df_negociacao['razao_social'] = df_negociacao['razao_social'].str.upper()

        # Realizar comparação por verossimilhança
        comparacoes = []
        for _, row_gerem in df_gerem_empresas.iterrows():
            empresa_gerem = row_gerem['empresa']
            nome_capital_gerem = row_gerem['empresa_nome_capital']
            id_gerem = row_gerem['id_gerem']
            data_interacao = row_gerem['data_interacao']

            for _, row_negociacao in df_negociacao.iterrows():
                nome_empresa_negociacao = row_negociacao['razao_social']
                codigo_negociacao = row_negociacao['codigo_negociacao']
                data_inicio_negociacao = row_negociacao['data_prim_ver_prop_tec']

                # Filtro de data: prospecção deve ser posterior à interação
                if pd.to_datetime(data_inicio_negociacao) <= pd.to_datetime(data_interacao):
                    continue

                # Comparação
                grau_nome_capital = calcular_grau_verossimilhanca(nome_capital_gerem, nome_empresa_negociacao)
                grau_final = grau_nome_capital

                if grau_final > 50:  # Considerar apenas comparações acima de 50
                    comparacoes.append({
                        'id_gerem': id_gerem,
                        'gerem_empresa': empresa_gerem,
                        'nome_capital': nome_capital_gerem,
                        'data_interacao': data_interacao,
                        'codigo_negociacao': codigo_negociacao,
                        'negociacao_empresa': nome_empresa_negociacao,
                        'data_inicio_negociacao': data_inicio_negociacao,
                        'grau_verossimilhanca': round(grau_final)
                    })

        # Criar DataFrame com os resultados da comparação
        df_comparacao = pd.DataFrame(comparacoes)

        # Criar o id_unico
        df_comparacao['id_unico'] = df_comparacao.apply(
            lambda x: f"{x['id_gerem']}_{x['codigo_negociacao']}",
            axis=1
        )

        # Verificar se o arquivo já existe e apagar, se necessário
        if os.path.exists(PATH_NEGOCIACAO_COMPARACAO):
            os.remove(PATH_NEGOCIACAO_COMPARACAO)

        # Exportar DataFrames
        df_comparacao.to_excel(PATH_NEGOCIACAO_COMPARACAO, index=False)

        print("OK - " + inspect.currentframe().f_code.co_name)

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

def negociacao_validacao():
    """
    Função para validar a comparação de negociação, adicionando as colunas 'status_analise_humana' e 'data_analise_humana' no DataFrame de comparação,
    e repassando os valores não analisados para o DataFrame de validação como novas linhas.

    Retorno:
    Cria e salva planilhas atualizadas com os dados analisados e não analisados.
    """
    try:
        # Lê os arquivos Excel
        df_comparacao = pd.read_excel(PATH_NEGOCIACAO_COMPARACAO)
        df_validacao = pd.read_excel(PATH_NEGOCIACAO_VALIDACAO)


        # Criar as colunas 'status_analise_humana' e 'data_analise_humana' em df_comparacao
        def obter_status_e_data(row):
            id_unico = row['id_unico']
            match = df_validacao[df_validacao['id_unico'] == id_unico]
            if not match.empty:
                return (
                    match['status_analise_humana'].iloc[0],
                    match['_validacao_verossimilhanca'].iloc[0],
                    match['data_analise_humana'].iloc[0]
                )
            else:
                return 'Não analisado', None, None

        df_comparacao[['status_analise_humana', 'validacao_verossimilhanca', 'data_analise_humana']] = df_comparacao.apply(
            lambda row: pd.Series(obter_status_e_data(row)), axis=1
        )


        # Identificar valores não validados e adicioná-los ao df_validacao
        nao_analisados = df_comparacao[df_comparacao['status_analise_humana'] == 'Não analisado']
        novas_linhas = nao_analisados[[
            'id_unico', 'id_gerem', 'gerem_empresa', 'nome_capital', 
            'data_interacao', 'codigo_negociacao', 'negociacao_empresa', 
            'data_inicio_negociacao', 'grau_verossimilhanca'
        ]]

        # Concatenar os novos registros ao DataFrame de validação
        df_validacao = pd.concat([df_validacao, novas_linhas], ignore_index=True)

        # Garantir que os campos que não são "Analisado" fiquem com "Não analisado"
        df_validacao['status_analise_humana'] = df_validacao['status_analise_humana'].apply(
            lambda x: 'Analisado' if x == 'Analisado' else 'Não analisado'
        )

        # Remover duplicados não analisados, considerando apenas 'id_unico' e 'id_prospeccao'
        # Criar um DataFrame separado apenas com os "Não analisados"
        df_nao_analisados = df_validacao[df_validacao['status_analise_humana'] == 'Não analisado']

        # Remover duplicatas apenas entre os "Não analisados", mantendo a primeira ocorrência
        df_nao_analisados = df_nao_analisados.drop_duplicates(subset=['id_unico', 'codigo_negociacao'], keep='first')

        # Criar um DataFrame separado com os "Analisados" (mantém todas as ocorrências)
        df_analisados = df_validacao[df_validacao['status_analise_humana'] == 'Analisado']

        # Reunir os dois DataFrames novamente
        df_validacao = pd.concat([df_analisados, df_nao_analisados], ignore_index=True)

        # Ordenar validação
        df_validacao = df_validacao.sort_values(by="grau_verossimilhanca", ascending=False)
        df_validacao = df_validacao.sort_values(by="negociacao_empresa")
        df_validacao = df_validacao.sort_values(by="nome_capital")
        df_validacao = df_validacao.sort_values(by="status_analise_humana")

        # Salvar os DataFrames
        # df_comparacao.to_excel(caminho_analisado, index=False)
        df_comparacao.to_excel(PATH_NEGOCIACAO_ANALISADO, index=False)
        df_validacao.to_excel(PATH_NEGOCIACAO_VALIDACAO_UP, index=False)

        print("OK - " + inspect.currentframe().f_code.co_name)

    except Exception as e:
        print(f"Erro ao processar os dados: {e}")
