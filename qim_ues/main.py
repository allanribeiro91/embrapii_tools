from scripts_public.buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from scripts_public.webdriver import configurar_webdriver
from scripts_public.baixar_dados_srinfo import baixar_dados_srinfo
from scripts_public.manipulacoes import pa_qim, resultados
from scripts_public.levar_arquivos_sharepoint import levar_arquivos_sharepoint


def qim_ues(buscar = False, baixar = False, manipular = False, levar = False):
    # buscando os valores existentes no SharePoint
    if buscar:
        buscar_arquivos_sharepoint()
    # baixando os valores do SRInfo
    if baixar:
        driver = configurar_webdriver()
        baixar_dados_srinfo(driver)

    # manipulando os valores para obter planilhas finais
    if manipular:
        pa, today = pa_qim()
        resultados(pa, today)

    # levando planilhas para sharepoint
    if levar:
        levar_arquivos_sharepoint()

if __name__ == "__main__":
    qim_ues(buscar=True, baixar=True, manipular=True, levar=True)