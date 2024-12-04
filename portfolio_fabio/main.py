from puxar_planilhas_sharepoint import puxar_planilhas
from manipulacoes import juntar_planilhas, ajustes, combinar_dados, processar_dados, empresas, gerar_planilha_unica
from levar_arquivos_sharepoint import levar_arquivos_sharepoint

def gerar_portfolio():

    # puxar planilhas do sharepoint
    print("Passo 1/7: Fazendo download das planilhas do SharePoint")
    puxar_planilhas()

    # juntar planilhas
    print("Passo 2/7: Juntando as planilhas")
    planilhas_juntas = juntar_planilhas()

    # PORTFOLIO
    # obter a primeira planilha conjunta (portfólio)
    merged = planilhas_juntas[0]
    # fazer os ajustes necessários no portfólio
    print("Passo 3/7: Fazendo as manipulações e ajustes necessários")
    merged2 = ajustes(merged)
    # combinar dados e obter o último mês do ipca
    ultimo_mes, extracao = combinar_dados(merged2)
    # processar dados para obter o portfólio finalizado
    print("Passo 4/7: Processando os dados do portfólio de projetos")
    processar_dados(ultimo_mes, extracao)

    # EMPRESAS
    # obter a segunda planilha conjunta (empresas)
    emp = planilhas_juntas[1]
    # processar dados das empresas e gerar a planilha final
    print("Passo 5/7: Gerando as informações das empresas")
    empresas(emp, extracao)

    # juntando as planilhas
    print("Passo 6/7: Gerando planilha única com 2 abas: portfólio e empresas")
    gerar_planilha_unica(extracao)

    # levar arquivos pro sharepoint
    print("Passo 7/7: Levando o arquivo para o SharePoint")
    levar_arquivos_sharepoint()


if __name__ == "__main__":
    gerar_portfolio()
