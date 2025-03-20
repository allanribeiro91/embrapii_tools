import clickhouse_connect
import socket
import os
import sys
from dotenv import load_dotenv
import inspect
import pyautogui
import subprocess
import time

load_dotenv()

ROOT = os.getenv('ROOT')
USUARIO = os.getenv('usuario_vpn')
SENHA = os.getenv('senha_vpn')
FORTICLIENT_PATH = os.getenv('forticlient_path')
sys.path.append(ROOT)


def conectar_vpn():
    # Passo 1: Abrir o FortiClient
    # Ajuste o caminho conforme a instalação no seu PC
    subprocess.Popen(FORTICLIENT_PATH)
    print("Abrindo FortiClient...")
    time.sleep(10)  # Aguarde o programa abrir (ajuste o tempo se necessário)

    # Passo 2: Preencher credenciais
    username = USUARIO
    password = SENHA

    # Passo 3: Escolher a VPN
    pyautogui.click(x=714, y=416) # Clicar no campo de escolha do nome da VPN
    pyautogui.click(x=685, y=479) # Escolher a segunda VPN da lista - SSL

    # Passo 4: Inserir credenciais e conectar
    pyautogui.press('tab') # Vai para o campo de usuário
    pyautogui.write(username)
    pyautogui.press('tab')  # Vai para o campo de senha
    pyautogui.write(password)
    pyautogui.press('enter')  # Conectar

    time.sleep(10)  # Aguarde a conexão ser estabelecida

    print("Tentando conectar à VPN...")


def is_vpn_connected(host, port):
    """
    Função para verificar se a VPN está conectada
    host: str - IP do servidor ClickHouse
    port: int - Porta do servidor ClickHouse
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
        try:
            socket.create_connection((host, port), timeout=5)
            return True  # Conexão bem-sucedida
        except (socket.timeout, OSError):
            return False  # Sem acesso ao banco (VPN pode estar desligada)
    
    #Ações a serem realizadas pela função
        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")


def consulta_clickhouse(host, port, user, password, query, pasta, nome_arquivo):
    """
    Função para consultar ao clickhouse e salvar o resultado em um arquivo CSV
    host: str - IP do servidor ClickHouse
    port: int - Porta do servidor ClickHouse
    user: str - Usuário do ClickHouse
    password: str - Senha do ClickHouse
    query: str - Consulta SQL
    pasta: str - Pasta onde o arquivo será salvo
    nome_arquivo: str - Nome do arquivo CSV
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:

        if is_vpn_connected(host, port):
            print("VPN conectada. Rodando a consulta...")

            # Conectar ao ClickHouse
            client = clickhouse_connect.get_client(host=host, port=port, user=user, password=password)

            # Executa a consulta e obtém os dados como DataFrame
            result = client.query_df(query)

            # Salvar o resultado em um arquivo CSV
            result.to_csv(os.path.abspath(os.path.join(ROOT, pasta ,f"{nome_arquivo}.csv")),
                        index=False, encoding="utf-8")

        else:
            print("VPN NÃO conectada! Conecte-se à VPN e tente novamente.")

        print("🟢 " + inspect.currentframe().f_code.co_name)
    except Exception as e:
        print(f"🔴 Erro: {e}")