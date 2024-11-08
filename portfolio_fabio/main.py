from puxar_planilhas_sharepoint import puxar_planilhas
from manipulacoes import juntar_planilhas, ajustes, combinar_dados, processar_dados, empresas, gerar_planilha_unica
from levar_arquivos_sharepoint import levar_arquivos_sharepoint

def gerar_portfolio():

    # # puxar planilhas do sharepoint
    puxar_planilhas()

    # juntar planilhas
    planilhas_juntas = juntar_planilhas()

    # PORTFOLIO
    # obter a primeira planilha conjunta (portfólio)
    merged = planilhas_juntas[0]
    # fazer os ajustes necessários no portfólio
    merged2 = ajustes(merged)
    # combinar dados e obter o último mês do ipca
    ultimo_mes = combinar_dados(merged2)
    # processar dados para obter o portfólio finalizado
    processar_dados(ultimo_mes)

    # EMPRESAS
    # obter a segunda planilha conjunta (empresas)
    emp = planilhas_juntas[1]
    # processar dados das empresas e gerar a planilha final
    empresas(emp)

    # juntando as planilhas
    gerar_planilha_unica()

    # levar arquivos pro sharepoint
    levar_arquivos_sharepoint()


if __name__ == "__main__":
    gerar_portfolio()
