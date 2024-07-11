import pytesseract
import cv2 #opencv
import pyautogui
from time import sleep, time

seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
pyautogui.click(seta_pista2_closed)
sleep(2)
pyautogui.scroll(100)

'''
#fechar todas as janelas
seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
pyautogui.click(seta_passeio1)
sleep(2)

seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.9)
pyautogui.click(seta_pista_rodagem2.x, seta_pista_rodagem2.y)

sleep(2)
seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
pyautogui.click(seta_pista_rodagem1.x, seta_pista_rodagem1.y)
sleep(2)
seta_passeio2 = pyautogui.locateCenterOnScreen('seta_passeio2.png', confidence=0.9)
pyautogui.click(seta_passeio2.x, seta_passeio2.y)
#abrir a janela necessária
seta_passeio1_closed = pyautogui.locateCenterOnScreen('seta_passeio1_closed.png', confidence=0.8)
pyautogui.click(seta_passeio1_closed)
sleep(2)
'''

