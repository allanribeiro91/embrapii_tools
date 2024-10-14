import pyautogui
import pygetwindow as gw  # Nova importação
import time

COORDENADAS_ALLAN = [(3331, -1185), (3189, -924), (2616, -852)]

def abrir_edge():
    # Pressiona a tecla 'Windows', digita 'Edge' e aperta Enter
    pyautogui.hotkey('win', 's')
    time.sleep(1)  # Espera para o menu iniciar abrir
    pyautogui.write('Edge')
    time.sleep(1)  # Espera o nome ser digitado
    pyautogui.press('enter')

def acessar_link():
    # Aguarda o navegador abrir
    time.sleep(5)
    # Pressiona Ctrl + L para focar na barra de navegação
    pyautogui.hotkey('ctrl', 'l')

    time.sleep(2)

    # Digita o link do relatório
    link = 'https://lookerstudio.google.com/u/0/reporting/95cf4522-1328-4bee-b3db-a79aa11459d6/page/p_0s3v6pwrld'
    pyautogui.write(link)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(10)


def clicar_nos_botoes():
    # Define as coordenadas com base na tela ativa
    coordenadas = COORDENADAS_ALLAN

    # Clica nas coordenadas definidas
    for x, y in coordenadas:
        pyautogui.click(x, y)
        time.sleep(2)  # Pausa entre cliques


def baixar_status_semanal():
    abrir_edge()
    acessar_link()
    clicar_nos_botoes()
    time.sleep(30)
    pyautogui.hotkey('alt', 'f4')




