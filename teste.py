import pandas as pd
import pyautogui
from time import sleep

postes = pyautogui.locateCenterOnScreen('entre_postes.png', confidence=0.6)
pyautogui.click(postes.x, postes.y)
    # Selecionar todo o texto existente e apagar
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

