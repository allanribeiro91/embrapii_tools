import pyautogui

# Este loop continua imprimindo as coordenadas do mouse até você interromper (Ctrl + C)
while True:
    x, y = pyautogui.position()  # Obtém a posição atual do cursor
    print(f"Posição atual do mouse: X={x}, Y={y}", end='\r')  # Imprime na mesma linha
