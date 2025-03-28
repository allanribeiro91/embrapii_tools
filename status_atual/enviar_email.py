import os
import win32com.client as win32
from dotenv import load_dotenv
from datetime import datetime

# Carrega variáveis de ambiente, se necessário
load_dotenv()

# Define o caminho da pasta onde a imagem PNG está localizada
ROOT = os.getenv('ROOT')
imagem_diretorio = os.path.join(ROOT, 'arquivo_pdf')

# Função para encontrar a imagem PNG dinamicamente
def encontrar_imagem_png(diretorio):
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith('.png'):
            return os.path.join(diretorio, arquivo)
    raise FileNotFoundError("Nenhuma imagem PNG encontrada na pasta.")

def enviar_email():
    # Busca a imagem PNG na pasta de forma dinâmica
    imagem_path = encontrar_imagem_png(imagem_diretorio)

    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)

    # email.To = 'milena.goncalves@embrapii.org.br'
    email.To = 'allan.ribeiro@embrapii.org.br; milena.goncalves@embrapii.org.br'
    email.Subject = 'Embrapii - Status Semanal ' + datetime.today().strftime('%d/%m/%Y')

    # Anexa a imagem e obtém o CID (Content-ID) para referência no HTML
    anexo = email.Attachments.Add(imagem_path)
    anexo.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ImagemPNG"
    )

    email.HTMLBody = """
        <img src="cid:ImagemPNG" alt="Status Semanal" style="width:80vw;">
        
        <p>
            Mais informações:
            <a href='https://app.powerbi.com/groups/me/apps/ccbb1664-f0e2-439f-b607-12a98a3341e2/reports/7a015c58-3f3d-40cd-9f75-2f4081ff6a2c?redirectedFromSignup=1&experience=power-bi'>
            Embrapii Dados
            </a>
        </p>

        <p>Att,<br>
        Allan Ribeiro
        </p>

    """

    email.Send()
    print('Email enviado')