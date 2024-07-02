import pandas as pd
import pyautogui
from time import sleep, time

cont_geral = 0

# Carregar os dados da planilha
df = pd.read_excel('table_itajai_test.xlsx', sheet_name='RIAN - V4P4')
# Verificar as colunas para encontrar os nomes corretos
print(df.columns)

# Extrair dados das colunas "larg_passeio_opost", "largura_via" e "larg_passeio_adj"
larg_passeio_opost = df['larg_passeio_opost'].tolist()
largura_via = df['largura_via'].tolist()
larg_passeio_adj = df['larg_passeio_adj'].tolist()
entre_postes = df['entre_postes'].tolist()
altura_lum = df['altura_lum'].tolist()
angulo = df['angulo'].tolist()
poste_pista = df['poste_pista'].tolist()
comprimento_braco = df['comprimento_braco'].tolist()

#------------ABRINDO CENARIO PADRAO ITAJAI-------------
# 1 - ABRINDO ARQUIVO
#pyautogui.doubleClick(147, 423, duration=0.5)
#sleep(30)  # TEMPO ATÉ ABRIR E CARREGAR O DIALUX

# Abrindo guia planejamento
pyautogui.click(399, 82, duration=0.5)

def tab_interate(cont):
    i = 0
    while i<cont:
        pyautogui.press('tab')
        i = i +1
    i = 0

def scroll_to_position(target_y, steps=200):
    start_time = time()
    while True:
        current_y = pyautogui.position().y
        if current_y >= target_y:
            break
        pyautogui.scroll(-steps)
        sleep(0.5)
        
        # Verifique se 10 segundos se passaram
        if time() - start_time >= 10:
            # Move o mouse para a posição atual da barra de rolagem
            scrollbar_position = pyautogui.position()
            pyautogui.moveTo(scrollbar_position.x, scrollbar_position.y)
            break

# Iterar sobre os valores extraídos e digitar no campo correspondente
for larg_passeio_oposto, larg_via, larg_passeio_adjacente, entre_postes_x, altura_lum_x, angulo_x, poste_pista_x, comprimento_braco_x in zip(larg_passeio_opost, largura_via, larg_passeio_adj, entre_postes, altura_lum, angulo, poste_pista, comprimento_braco):
    cont_geral = cont_geral +1 #var para fazer a contagem de cenários 
    cont__str = str(cont_geral); #var para fazer conversão de int para string e passar como parametro no nome do cenário
    # Selecionando o passeio1
    pyautogui.doubleClick(102, 586, duration=0.5)
    sleep(0.6)
    # Clicando no campo largura passeio
    pyautogui.doubleClick(269, 754, duration=0.5)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_opost
    pyautogui.write(str(larg_passeio_oposto))
    sleep(0.5)

    ##--------------------- PARAMETROS RUA ---------------------
    # Clicando no campo largura via (ajustar coordenadas conforme necessário)
    pyautogui.doubleClick(118, 621, duration=0.5)  # clicando na rua
    sleep(0.5)
    pyautogui.doubleClick(262, 820, duration=0.5)  # clicando no campo de largura
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para largura_via
    pyautogui.write(str(larg_via))
    sleep(0.5)

    ##--------------------- PARAMETROS PASSEIO ADJACENTE ---------------------
    # Selecionando o passeio2
    pyautogui.doubleClick(112, 651, duration=0.5)
    sleep(0.5)
    # Clicando no campo largura passeio
    pyautogui.doubleClick(244, 754, duration=0.5)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_adjacente
    pyautogui.write(str(larg_passeio_adjacente))
    sleep(0.8)

    ##--------------------- PARAMETROS LUMINÁRIA ---------------------
    img = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.7)
    pyautogui.click(img.x, img.y)
    sleep(0.8)
    
    # Posicione o mouse sobre a scrollbar 
    #pyautogui.moveTo(492, 512)  # Ajuste as coordenadas conforme necessário
    #target = 917
    #scroll_to_position(target, 300)
    #sleep(1.5)

    
    #Distância entre postes entre_postes
    #postes = pyautogui.locateCenterOnScreen('entre_postes.png', confidence=0.6)
    #pyautogui.click(postes.x, postes.y)
  
    #Ditancia entr postes
    tab_interate(20)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_adjacente
    pyautogui.write(str(entre_postes_x))
    sleep(0.5)
    #Altura do ponto de luz
    tab_interate(3)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(altura_lum_x))

    #Angulo
    tab_interate(5)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(angulo_x))

    #distância poste-pista
    tab_interate(5)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(poste_pista_x))
    #comprimento do braço
    tab_interate(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(comprimento_braco_x))

    sleep(1)
    projecto = pyautogui.locateCenterOnScreen('guia_projecto.png', confidence=0.6)
    pyautogui.click(projecto.x, projecto.y)
    sleep(0.3)
    nome_projet = pyautogui.locateCenterOnScreen('nome_projeto.png', confidence=0.6)
    pyautogui.click(nome_projet.x, nome_projet.y)
    pyautogui.write(str("Itajai " + cont__str))
    

    #fazer a comparação de qual luminaria é a mais eficiente, para isso vamos tirar print dos resultados, extrair o texto das imagens 
    # e fazer uma comparação pra ver qual esta mais próximo do resultado. https://awari.com.br/ocr-em-python-aprenda-a-extrair-texto-de-imagens-com-facilidade/
    #ideia de comparação: extrair qual a classe da via e gerar um script que gera automaticamente uma planilha com os parametros para comparação