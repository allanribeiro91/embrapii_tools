import requests

# URL correta do Webhook (verifique se está completa)
WEBHOOK_URL = "https://prod-25.brazilsouth.logic.azure.com:443/workflows/c5858eb547ea417086721e241a56d3f0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4N9d4Aj4IDQw5OnxMp2eApPNVy34Jn4rRvO1VZZ3dG4"

# Mensagem dentro do campo "attachments"
payload = {
    "attachments": [
        {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "Só alegria/n, só o ouro✅!",
                        "weight": "Bolder",
                        "size": "Large"
                    }
                ]
            }
        }
    ]
}

# Enviar a requisição
response = requests.post(WEBHOOK_URL, json=payload)

# Verificar a resposta
if response.status_code == 200 or response.status_code == 202:
    print("Mensagem enviada com sucesso! (Aguardando processamento)")
else:
    print(f"Erro ao enviar mensagem: {response.status_code}, {response.text}")
