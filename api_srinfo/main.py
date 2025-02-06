import requests
import os
import pandas as pd
import time
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# Credenciais de login
USERNAME = os.getenv('usuario')
PASSWORD = os.getenv('senha')

# URLs da API
URL_TOKEN = "https://srinfo.embrapii.org.br/token/"
API_ENDPOINTS = {
    "Reserva de Recursos": "https://srinfo.embrapii.org.br/partnerships/api/fundsapprovals/",
    "Empresas Contatos Avaliacao Projetos": "https://srinfo.embrapii.org.br/analytics/api/reports/evaluationreferences/?start_date=2000-01-01&end_date=2030-01-01&format=json", 
    "Termo Cooperacao": "https://srinfo.embrapii.org.br/accreditation/api/cooperationterm/",
    # "Repasse recursos": "https://srinfo.embrapii.org.br/financial/fundstransfer_json/?draw=0&length=25",
}

def get_token(username, password):
    """Obtém o token de autenticação."""
    try:
        response = requests.post(URL_TOKEN, data={"username": username, "password": password})
        if response.status_code == 200:
            token_data = response.json()
            return token_data.get("access", None)
        else:
            print(f"Erro ao obter o token. Status code: {response.status_code}, Detalhes: {response.text}")
    except Exception as e:
        print(f"Erro na requisição do token: {e}")
    return None

def fetch_data(url, token):
    """Busca os dados paginados da API."""
    headers = {"Authorization": f"Bearer {token}"}
    data = []

    print(f"Buscando dados de: {url}...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        try:
            # Tenta converter para JSON
            json_data = response.json()
            
            # Verifica se é uma lista ou um dicionário
            if isinstance(json_data, list):
                # Resposta é uma lista direta
                data.extend(json_data)
            elif isinstance(json_data, dict) and "results" in json_data:
                # Resposta é um dicionário com paginação
                data.extend(json_data.get("results", []))
            else:
                print("Formato de resposta inesperado:", json_data)
        except Exception as e:
            print(f"Erro ao processar a resposta JSON: {e}")
    else:
        print(f"Erro ao buscar dados. Status code: {response.status_code}, Detalhes: {response.text}")
    
    return data

def criar_arquivo(data, caminho_saida):
    """
    Recebe os dados em formato JSON e salva em um arquivo Excel, adicionando a coluna "data_extracao".
    Aplica ajustes nos campos "id" e "status" para o arquivo "termo_cooperacao.xlsx".

    Args:
        data (list): Dados em formato JSON (lista de dicionários).
        caminho_saida (str): Caminho completo para salvar o arquivo Excel.
    """
    if not data:
        print("Nenhum dado para salvar.")
        return

    try:
        # Converte os dados para um DataFrame
        df = pd.DataFrame(data)

        # Ajustes específicos para "termo_cooperacao.xlsx"
        if "termo_cooperacao.xlsx" in caminho_saida.lower():
            # Ajuste para o campo "id"
            if "id" in df.columns:
                df["id"] = df["id"].apply(lambda x: x.get("id") if isinstance(x, dict) else x)

            # Ajuste para o campo "status"
            if "status" in df.columns:
                df["status"] = df["status"].apply(lambda x: x[1] if isinstance(x, list) and len(x) > 1 else x)

        # Adiciona a coluna "data_extracao" com a data atual
        data_atual = datetime.now().strftime("%d/%m/%Y")
        df["data_extracao"] = data_atual

        # Salva os dados no formato Excel
        df.to_excel(caminho_saida, index=False, sheet_name="Dados")
        print(f"Arquivo salvo com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")


def main():
    # Obtém o token de autenticação
    token = get_token(USERNAME, PASSWORD)
    if not token:
        print("Não foi possível obter o token. Verifique as credenciais.")
        return

    # Busca os dados de cada endpoint
    for name, url in API_ENDPOINTS.items():
        print(f"=======================================")
        print(f"Importando dados de: {name}")
        data = fetch_data(url, token)
        print(f"Dados recebidos de {name}: {len(data)} registros.")

        # Define o caminho de saída para cada endpoint
        caminho_saida = f"{name.lower().replace(' ', '_')}.xlsx"
        criar_arquivo(data, caminho_saida)

if __name__ == "__main__":
    main()



