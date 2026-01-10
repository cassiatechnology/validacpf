import pyautogui
import time

print("Pressione Ctrl+C para interromper o programa.")

try:
    time.sleep(5)
    while True:     
        # Pega as coordenadas x e y do mouse
        x, y = pyautogui.position()
        
        # Cria uma string com as coordenadas
        posicao_str = f"X: {x} | Y: {y}"
        
        # O '\r' faz com que o texto seja sobrescrito na mesma linha do terminal
        # O 'end=""' evita que ele pule para a próxima linha
        print(posicao_str.ljust(20), end="\r")
        
        # Pequena pausa para não sobrecarregar o processador
        time.sleep(0.1)

except KeyboardInterrupt:
    print("\nPrograma encerrado.")