import pyautogui

# Obter coordenadas do canto superior esquerdo
print("Posicione o cursor no canto superior esquerdo da área de resultados e pressione Enter.")
input()
left, top = pyautogui.position()
print(f"Canto superior esquerdo: ({left}, {top})")

# Obter coordenadas do canto inferior direito

print("Posicione o cursor no canto inferior direito da área de resultados e pressione Enter.")
input()
right, bottom = pyautogui.position()
print(f"Canto inferior direito: ({right}, {bottom})")

# Calcular largura e altura
width = right - left
height = bottom - top
print(f"Largura: {width}, Altura: {height}")
                