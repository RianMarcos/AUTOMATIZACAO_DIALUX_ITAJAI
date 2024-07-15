import pytesseract
import cv2 #opencv
import pyautogui
from time import sleep, time
'''
#----------------------------------------#MODIFICANDO CLASSIFIÇÃO Em e Uo-------------------------------------------#
#fechar todas as janelas 
#FAZER LOGICA QUE SÓ VAI FECHAR SEGUNDA PISTA DE HOUVER CANTEIRO CENTRAL
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
sleep(1)

#modificar em 
em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
pyautogui.click(em_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')


#modificar uo
uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
pyautogui.click(uo_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')


#fechar janela modificada e passar para próxima 
seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
pyautogui.click(seta_passeio1)
sleep(1)
pyautogui.scroll(-300)
sleep(1)
#passar para proxima
#preencher valroes para canteiro central

#abrir a janela necessária
seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
pyautogui.click(seta_pista2_closed)
sleep(2)

#modificar em 
em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
pyautogui.click(em_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

#modificar uo
uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
pyautogui.click(uo_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

pyautogui.scroll(-300)
sleep(1)

#fechar janela modificada e passar para próxima 
seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.96)
pyautogui.click(seta_pista_rodagem2)
sleep(1)
pyautogui.scroll(-300)
sleep(1)

#abrir a janela necessária
seta_pista1_closed = pyautogui.locateCenterOnScreen('seta_pista1_closed.png', confidence=0.9)
pyautogui.click(seta_pista1_closed)
sleep(1)

#modificar em 
em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
pyautogui.click(em_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')


#modificar uo
uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
pyautogui.click(uo_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

sleep(0.2)
pyautogui.scroll(-300)
sleep(2)

#fechar janela modificada e passar para próxima 
seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
pyautogui.click(seta_pista_rodagem1)
sleep(2)
pyautogui.scroll(-600)
sleep(2)

#-----------acima esta ok-----------------

#abrir a janela necessária
seta_passeio2_closed = pyautogui.locateCenterOnScreen('seta_passeio2_closed.png', confidence=0.9)
pyautogui.click(seta_passeio2_closed)
sleep(1)
pyautogui.scroll(-600)
sleep(1)
#modificar em 
em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
pyautogui.click(em_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

#modificar uo
uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
pyautogui.click(uo_parametro)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

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
valida_central =1
#-----abrindo todas as guias-----
seta_pista1_closed = pyautogui.locateCenterOnScreen('seta_pista1_closed.png', confidence=0.9)
pyautogui.click(seta_pista1_closed)
sleep(2)
#abrir a janela necessária
seta_passeio1_closed = pyautogui.locateCenterOnScreen('seta_passeio1_closed.png', confidence=0.8)
pyautogui.click(seta_passeio1_closed)
sleep(1)

if(valida_central == 1):
    #abrir a janela necessária
    seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
    pyautogui.click(seta_pista2_closed)
    sleep(2)

pyautogui.scroll(-1000)
    #abrir todas as janelas novamente para verificar os checks (manter aberta)
    #lembrar de usar rolagem scroll
#abrir todas as janelas novamente para verificar os checks (manter aberta)
#lembrar de usar rolagem scroll

