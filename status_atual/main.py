from baixar_status_semanal import baixar_status_semanal
from mover_arquivo import mover_arquivo
from levar_arquivos_sharepoint import levar_arquivos_sharepoint
from enviar_email import enviar_email


def main_status_semanal():
    baixar_status_semanal()
    mover_arquivo()
    levar_arquivos_sharepoint()
    enviar_email()



if __name__ == "__main__":
    main_status_semanal()