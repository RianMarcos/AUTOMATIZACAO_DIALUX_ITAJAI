import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os

cont_geral = 0
check_distri = 0

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

# Carregar os dados da planilha
df = pd.read_excel('table_itajai_test.xlsx', sheet_name='RIAN - V4P4')
# Verificar as colunas para encontrar os nomes corretos
print(df.columns)

# Adicionar colunas 'luminaria_escolhida' e 'angulo_escolhido' se não existirem
if 'luminaria_escolhida' not in df.columns:
    df['luminaria_escolhida'] = ""
if 'angulo_escolhido' not in df.columns:
    df['angulo_escolhido'] = ""
if 'cenario' not in df.columns:
    df['cenario'] = ""


# Garantir que a coluna 'luminaria_escolhida' é do tipo object
df['luminaria_escolhida'] = df['luminaria_escolhida'].astype(object)

# Extrair dados das colunas "larg_passeio_opost", "largura_via" e "larg_passeio_adj"
larg_passeio_opost = df['larg_passeio_opost'].tolist()
largura_via = df['largura_via'].tolist()
larg_passeio_adj = df['larg_passeio_adj'].tolist()
entre_postes = df['entre_postes'].tolist()
altura_lum = df['altura_lum'].tolist()
angulo = df['angulo'].tolist()
poste_pista = df['poste_pista'].tolist()
comprimento_braco = df['comprimento_braco'].tolist()
distribuicao = df['distribuicao'].str.lower().tolist()      
# Converter a coluna 'qtde_faixas' para inteiros
# Preencher valores ausentes com 0 e converter a coluna 'qtde_faixas' para inteiros
df['qtde_faixas'] = df['qtde_faixas'].fillna(0).astype(int)
qtde_faixas = df['qtde_faixas'].tolist()
qtde_ruas = df['qtde_ruas'].tolist()
larg_canteiro_central = df['larg_canteiro_central'].tolist()
pendor = df['pendor'].tolist()
classe_via = df['classe_via'].str.lower().tolist() 
classe_passeio = df['classe_passeio'].str.lower().tolist() 


#------------ABRINDO CENARIO PADRAO ITAJAI-------------
# 1 - ABRINDO ARQUIVO
#pyautogui.doubleClick(147, 423, duration=0.5)
#sleep(30)  # TEMPO ATÉ ABRIR E CARREGAR O DIALUX

def click_image(image_path, confidence=0.7, double_click=False):
    location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
    if location:
        if double_click:
            pyautogui.doubleClick(location.x, location.y)
        else:
            pyautogui.click(location.x, location.y)
        sleep(0.5)
    else:
        print(f"Imagem {image_path} não encontrada.")


def to_upper_safe(texto):
    if not isinstance(texto, str):
        raise TypeError("A variável fornecida não é uma string")
    
    # Remover espaços em branco e caracteres invisíveis
    texto = texto.strip()
    
    try:
        texto_maiusculo = texto.upper()
        return texto_maiusculo
    except Exception as e:
        print(f"Erro ao converter para maiúsculas: {e}")
        return None

def exclui_passeio(check_passeio_adjacente, check_passeio_oposto):
    sleep(1)
    #abrir guia das ruas
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.8)
    pyautogui.click(ruas)   
    sleep(1.5)

    print("Excluindo passeios necessários")
    if(check_passeio_oposto == 0): #será necessário excluir primeiro passeio
        try:
            passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.8)
            check_passeio1 = 1 if passeio1 is not None else 0  
        except pyautogui.ImageNotFoundException:
            check_passeio1 = 0
        if check_passeio1 == 1: 
            print("Passeio1 Encontrado")
            passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.8)
            pyautogui.click(passeio1)
            sleep(0.9)
            remover = pyautogui.locateCenterOnScreen('remover2.png', confidence=0.9)
            pyautogui.click(remover) 
            sleep(0.7)
        else:
            print("Passeio1 já foi excluido") 
            sleep(0.5)
    else:
        print("Manter primeiro passeio")
    sleep(1)
    if(check_passeio_adjacente == 0): #será necessário excluir segundo passeio
        try:
            sleep(0.5)
            passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.8)
            check_passeio2 = 1 if passeio2 is not None else 0  
        except pyautogui.ImageNotFoundException:
            check_passeio2 = 0
        if check_passeio2 == 1: 
            print("Passeio2 Encontrado")
            sleep(0.5)
            passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.8)
            pyautogui.click(passeio2)
            sleep(0.9)
            remover = pyautogui.locateCenterOnScreen('remover2.png', confidence=0.9)
            pyautogui.click(remover) 
            sleep(0.8)
        else:
            print("Passeio2 já foi excluido") 
            sleep(0.5)
    else:
        print("Manter segundo passeio")

def verifica_add_passeio():
    #entra todo começo de loop para adicionar passeio se ainda nao tem 
    #verifica se ja existe os dois passeios, se nao existir adiciona 
    try:
        passeio1 = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.9)
        check_passeio1 = 1 if passeio1 is not None else 0   
    except pyautogui.ImageNotFoundException:
        check_passeio1 = 0
    if check_passeio1 == 1: 
        print("Passeio1 Encontrado")
    else:
        add_passeio = pyautogui.locateCenterOnScreen('add_passeio.png', confidence=0.8)
        pyautogui.click(add_passeio)
        print("Passeio1 adicionado") 
        sleep(1)

        #clicar no primeiro passeio
        first_passeio = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.8)
        pyautogui.click(first_passeio)
        sleep(1)
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        name_passeio_1 = "Passeio 1 (C3)"
        pyautogui.write(str(name_passeio_1))
        tab_interate(4)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        name_passeio_1 = "Passeio 1 (C3)"
        pyautogui.write(str(name_passeio_1))
        sleep(0.3)
        tab_interate(1)
        pyautogui.press('left', presses=9)
        sleep(8)



    try:
        passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.9)
        check_passeio2 = 1 if passeio2 is not None else 0   # Verifica se a imagem 'central.png' foi encontrada
    except pyautogui.ImageNotFoundException:
        check_passeio2 = 0
    if check_passeio2 == 1: 
        print("Passeio2 Encontrado")
    else:
        add_passeio = pyautogui.locateCenterOnScreen('add_passeio.png', confidence=0.8)
        pyautogui.click(add_passeio)
        print("Passeio2 adicionado")  
        sleep(2) 

        #clicar no primeiro passeio
        first_passeio = pyautogui.locateCenterOnScreen('first_passeio.png', confidence=0.8)
        pyautogui.click(first_passeio)
        sleep(1)

        #moficar nome
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        nome_passeio_2 = "Passeio 2"
        pyautogui.write(str(nome_passeio_2))
        tab_interate(4)
        sleep(0.5)
        name_passeio = "PASSEIO"
        pyautogui.write(str(name_passeio))
        tab_interate(1)
        pyautogui.press('left', presses=9)
        sleep(8)

        #descer para último
        seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        pyautogui.click(seta_baixo.x, seta_baixo.y)
        sleep(5)


def classifica_vias_passeios():
    seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
    pyautogui.click(seta_passeio1)
    sleep(1)

    if(valida_central == 1):
        seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.9)
        pyautogui.click(seta_pista_rodagem2.x, seta_pista_rodagem2.y)

    sleep(1)
    seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
    pyautogui.click(seta_pista_rodagem1.x, seta_pista_rodagem1.y)
    sleep(1)
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
    pyautogui.write(str(classe_passeio_em))

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_uo))

    #fechar janela modificada e passar para próxima 
    seta_passeio1 = pyautogui.locateCenterOnScreen('seta_passeio1.png', confidence=0.8)
    pyautogui.click(seta_passeio1)
    sleep(1)
    pyautogui.scroll(-300)
    sleep(1)
    #passar para proxima
    if(valida_central == 1):
        #preencher valroes para canteiro central

        #abrir a janela necessária
        seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
        pyautogui.click(seta_pista2_closed)
        sleep(1)

        #modificar em 
        em_parametro = pyautogui.locateCenterOnScreen('em_parametro.png', confidence=0.8)
        pyautogui.click(em_parametro)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(classe_via_em))

        #modificar uo
        uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
        pyautogui.click(uo_parametro)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(classe_via_uo))
        pyautogui.scroll(-300)
        sleep(1)

        #fechar janela modificada e passar para próxima 
        seta_pista_rodagem2 = pyautogui.locateCenterOnScreen('seta_pista_rodagem2.png', confidence=0.9)
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
    pyautogui.write(str(classe_via_em))

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_via_uo))
    sleep(0.2)
    pyautogui.scroll(-300)
    sleep(1)

    #fechar janela modificada e passar para próxima 
    seta_pista_rodagem1 = pyautogui.locateCenterOnScreen('seta_pista_rodagem1.png', confidence=0.9)
    pyautogui.click(seta_pista_rodagem1)
    sleep(1)
    pyautogui.scroll(-600)
    sleep(1)
    
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
    pyautogui.write(str(classe_passeio_em))

    #modificar uo
    uo_parametro = pyautogui.locateCenterOnScreen('uo_parametro.png', confidence=0.8)
    pyautogui.click(uo_parametro)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(classe_passeio_uo))

    #-----abrindo todas as guias-----
    seta_pista1_closed = pyautogui.locateCenterOnScreen('seta_pista1_closed.png', confidence=0.9)
    pyautogui.click(seta_pista1_closed)
    sleep(1)
    #abrir a janela necessária
    seta_passeio1_closed = pyautogui.locateCenterOnScreen('seta_passeio1_closed.png', confidence=0.8)
    pyautogui.click(seta_passeio1_closed)
    sleep(1)

    if(valida_central == 1):
        #abrir a janela necessária
        seta_pista2_closed = pyautogui.locateCenterOnScreen('seta_pista2_closed.png', confidence=0.9)
        pyautogui.click(seta_pista2_closed)
        sleep(1)

    pyautogui.scroll(-1000)
    #abrir todas as janelas novamente para verificar os checks (manter aberta)
    #lembrar de usar rolagem scroll

#função para verificar se possui canteiro central e fazer devido deslocamento de posição via 'tab'
def teste_central(x_img, y_img, tabs):
    try:
        img_central = pyautogui.locateCenterOnScreen('central.png', confidence=0.8)
        auxiliar = 1 if img_central is not None else 0   # Verifica se a imagem 'central.png' foi encontrada
    except pyautogui.ImageNotFoundException:
        auxiliar = 0
    if auxiliar == 1: 
        pyautogui.click(x_img, y_img)
        tab_interate(tabs)
        print("Canteiro central encontrado")
    else:
        pyautogui.click(x_img, y_img)
        tab_interate(tabs - 1)
        print("Canteiro central não encontrado")  

def check_all(screenshot_path, validacao_central):
    if not os.path.exists(screenshot_path):
        print(f"File not found: {screenshot_path}")
        return False

    image = cv2.imread(screenshot_path)
    if image is None:
        print(f"Failed to load image: {screenshot_path}")
        return False

    # Convert image to HSV color space
    hsv_image = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

    # Define range for green color in HSV
    lower_green = np.array([35, 100, 100])
    upper_green = np.array([85, 255, 255])

    # Create a mask for green color
    mask = cv2.inRange(hsv_image, lower_green, upper_green)

    # Debugging: save the mask to visualize
    mask_path = screenshot_path.replace('.png', '_mask.png')
    cv2.imwrite(mask_path, mask)
    print(f"Mask saved to {mask_path}")

    # Find contours in the mask
    contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    # Count the number of contours
    num_checks = len(contours)
    print(f"Number of checks found: {num_checks}")

    # Check if the number of contours (checks)
    if(check_passeio_oposto == 1 and check_passeio_adjacente ==1):
        if(validacao_central == 1):
            if num_checks >= 8:
                return True
            return False
        else:
            if num_checks >= 6:
                return True
            return False
    elif(check_passeio_oposto == 0 and check_passeio_adjacente == 1):   
        if(validacao_central == 1):
            if num_checks >= 6:
                return True
            return False
        else:
            if num_checks >= 4:
                return True
            return False
        
    elif(check_passeio_oposto == 1 and check_passeio_adjacente == 0):   
        if(validacao_central == 1):
            if num_checks >= 6:
                return True
            return False
        else:
            if num_checks >= 4:
                return True
            return False
        
    elif(check_passeio_oposto == 0 and check_passeio_adjacente == 0):   
        if(validacao_central == 1):
            if num_checks >= 4:
                return True
            return False
        else:
            if num_checks >= 2:
                return True
            return False  


def tab_interate(cont):
    i = 0
    while i < cont:
        pyautogui.press('tab')
        i += 1
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
for idx, (larg_passeio_oposto, larg_via, larg_passeio_adjacente, entre_postes_x, altura_lum_x, angulo_x, poste_pista_x, comprimento_braco_x, qtde_faixas_x, larg_canteiro_central_x, pendor_x, classe_via_x, classe_passeio_x) in enumerate(zip(larg_passeio_opost, largura_via, larg_passeio_adj, entre_postes, altura_lum, angulo, poste_pista, comprimento_braco, qtde_faixas, larg_canteiro_central, pendor, classe_via, classe_passeio)):

    if(classe_via_x == "v1" or classe_via_x == "V1"):
        classe_via_em = 30
        classe_via_uo = 0.4
        print(classe_via_em)
        print(classe_via_uo)
    elif(classe_via_x == "v2"):
        classe_via_em = 20
        classe_via_uo = 0.3
    elif(classe_via_x == "v3"):
        classe_via_em = 15
        classe_via_uo = 0.2
    elif(classe_via_x == "v4"):
        classe_via_em = 10
        classe_via_uo = 0.2
    elif(classe_via_x == "v5"):
        classe_via_em = 5
        classe_via_uo = 0.2

    if(classe_passeio_x == "p1"):
        classe_passeio_em = 20
        classe_passeio_uo = 0.3
    elif(classe_passeio_x == "p2"):
        classe_passeio_em = 10
        classe_passeio_uo = 0.25
    elif(classe_passeio_x == "p3"):
        classe_passeio_em = 5
        classe_passeio_uo = 0.2
    elif(classe_passeio_x == "p4"):
        classe_passeio_em = 3
        classe_passeio_uo = 0.2

    #Verificar se será necessário excluir ou adicionar um passeio
    if(larg_passeio_oposto == 0):
        check_passeio_oposto = 0
    else: 
        check_passeio_oposto = 1

    if(larg_passeio_adjacente == 0):
        check_passeio_adjacente = 0
    else: 
        check_passeio_adjacente = 1

    print("Distribuição: "+ distribuicao[idx])

    cont_geral += 1  # var para fazer a contagem de cenários 
    cont__str = str(cont_geral)  # var para fazer conversão de int para string e passar como parametro no nome do cenário
    
    # Abrindo guia planejamento
    pyautogui.click(399, 82, duration=0.5)
    sleep(1)
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.6)
    pyautogui.click(ruas.x, ruas.y)
    sleep(1)

    verifica_add_passeio()

    auxiliar_1 = 0
    if distribuicao[cont_geral-1] == 'central' or distribuicao[cont_geral-1]== 'canteiro central' or distribuicao[cont_geral-1]==  'canteiro_central' or distribuicao[cont_geral-1]== 'central':
        #validar se ja existe canteiro central, se nao existir adicionar 
        try:
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            auxiliar_1 = 1 if faixa_central_1 is not None else 0  
        except pyautogui.ImageNotFoundException:
            auxiliar_1 = 0
            print('Imagem da faixa central não encontrada')
        if auxiliar_1 == 1: 
            print("Faixa ja adicionada")
        else:
            adicionar_faixa_central = pyautogui.locateCenterOnScreen('adicionar_faixa_central.png', confidence=0.8)
            pyautogui.click(adicionar_faixa_central.x, adicionar_faixa_central.y)
            sleep(1)
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            pyautogui.click(faixa_central_1.x, faixa_central_1.y)
            sleep(2)
            seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
            pyautogui.click(seta_baixo.x, seta_baixo.y)
            sleep(1)
            adicionar_pista = pyautogui.locateCenterOnScreen('adicionar_pista.png', confidence=0.9)
            pyautogui.click(adicionar_pista.x, adicionar_pista.y)
            sleep(2)
            pista_de_rodagem2 = pyautogui.locateCenterOnScreen('pista_de_rodagem2.png', confidence=0.9)
            pyautogui.click(pista_de_rodagem2.x, pista_de_rodagem2.y)
            sleep(1.5)
            tab_interate(10)
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            pyautogui.press('left')
            sleep(8)
            seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
            pyautogui.click(seta_baixo.x, seta_baixo.y)
            sleep(1)

        #restante dos passos para inserir valores nos campos correspondentes 
        #PASSEIO1
        sleep(1.5)
        passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.7)
        pyautogui.doubleClick(passeio1.x, passeio1.y)
        sleep(1.5)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_oposto))
        sleep(1.5)

        #PISTA DE RODAGEM2
        tab_interate(11)
        sleep(1)
        pyautogui.press('Down')
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
        # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
        # pyautogui.doubleClick(pista1.x, pista1.y)
        # sleep(1)
        tab_interate(6)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1)

        tab_interate(1)
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        sleep(0.3)
        # Digitando o novo valor para
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1.5)
   
        #CANTEIRO CENTRAL
        tab_interate(12) #validar se esta certo
        pyautogui.press('Down')
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(larg_canteiro_central_x))

        #PISTA DE RODAGEM1
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('Down')
        sleep(0.5)
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
        # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
        # pyautogui.doubleClick(pista1.x, pista1.y)
        # sleep(1)
        tab_interate(6)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1)
        tab_interate(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1)
        
        #PASSEIO2
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('Down')
        passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.7)
        pyautogui.doubleClick(passeio2.x, passeio2.y)
        sleep(1)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_adjacente))
        sleep(1.5)      

    else:
        try: 
            #antes de remover o passeio será necessário mover a distribuicao do canteiro central p/ qqlr outra
            luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.7)
            pyautogui.doubleClick(luminaria.x, luminaria.y)
            sleep(0.5)

            tab_interate(16)
            bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence=0.8)
            pyautogui.doubleClick(bilateral.x, bilateral.y)
            sleep(0.5)  
            
            ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.8)
            pyautogui.doubleClick(ruas.x, ruas.y)
            sleep(0.5)  

            #removendo canteiro central e segunda via adicionada
            faixa_central_1 = pyautogui.locateCenterOnScreen('faixa_central_1.png', confidence=0.8)
            pyautogui.click(faixa_central_1.x, faixa_central_1.y)
            sleep(2)
            remover2 = pyautogui.locateCenterOnScreen('remover2.png', confidence=0.9)
            pyautogui.click(remover2.x, remover2.y)
            sleep(2)

            pista_de_rodagem2 = pyautogui.locateCenterOnScreen('pista_de_rodagem2.png', confidence=0.9)
            pyautogui.click(pista_de_rodagem2.x, pista_de_rodagem2.y)
            sleep(0.8)
            remover = pyautogui.locateCenterOnScreen('remover.png', confidence=0.9)
            pyautogui.click(remover.x, remover.y)
            sleep(2)
        except pyautogui.ImageNotFoundException:
            print("Imagem do canteiro central nao encontrada")
    
        # Selecionando o passeio1
        #tab_interate(8)
        sleep(1.5)
        passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.7)
        pyautogui.doubleClick(passeio1.x, passeio1.y)
        sleep(1)
        tab_interate(3)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para larg_passeio_opost
        pyautogui.write(str(larg_passeio_oposto))
        sleep(1.5)

        ##--------------------- PARAMETROS RUA ---------------------
        tab_interate(11)
        sleep(1)
        pyautogui.press('Down')
        # Clicando no campo largura via (ajustar coordenadas conforme necessário)
    # pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.7)
    # pyautogui.doubleClick(pista1.x, pista1.y)
    # sleep(1)
        tab_interate(6)
        # Selecionar todo o texto existente e apagar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        # Digitando o novo valor para largura_via
        pyautogui.write(str(larg_via))
        sleep(1)

        tab_interate(1)
        print("A quantidade de faixas é: "+ str(qtde_faixas_x))
        sleep(0.5)
        # Digitando qtde de faixas
        pyautogui.write(str(qtde_faixas_x))
        sleep(1.5)
    
        ##--------------------- PARAMETROS PASSEIO ADJACENTE ---------------------
        tab_interate(12)
        sleep(1)
        pyautogui.press('Down')
        # Selecionando o passeio2
    # passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.7)
        #pyautogui.doubleClick(passeio2.x, passeio2.y)
        #sleep(1)
        tab_interate(3)
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
    tab_interate(16)

    #clicar no bilateral para atualizar distruibuições e liberar canteiro central
    img_bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence =0.7)
    pyautogui.click(img_bilateral.x, img_bilateral.y)
    sleep(2.5)
    img_uni = pyautogui.locateCenterOnScreen('unilateral_inferior.png', confidence =0.7)
    pyautogui.click(img_uni.x, img_uni.y)
    sleep(2.5)

    # Posicione o mouse sobre a scrollbar 
    #pyautogui.moveTo(492, 512)  # Ajuste as coordenadas conforme necessário
    #target = 917
    #scroll_to_position(target, 300)
    #sleep(1.5)

    #Distância entre postes entre_postes
    #postes = pyautogui.locateCenterOnScreen('entre_postes.png', confidence=0.6)
    #pyautogui.click(postes.x, postes.y)
    #Selecionando tipo de distruibuicao dos postes

    valida_central = 0
    if distribuicao[cont_geral-1] == 'unilateral' or distribuicao[cont_geral-1] == 'unilateral inferior' or distribuicao[cont_geral-1] == 'unilateral_inferior' :
        img_uni = pyautogui.locateCenterOnScreen('unilateral_inferior.png', confidence =0.7)
        pyautogui.click(img_uni.x, img_uni.y)
        print("ENTROU NO UNILATERAL")
        sleep(1)
        x_img = img_uni.x #posicao da distruibuicao
        y_img = img_uni.y
        tabs = 6 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)
  
    elif distribuicao[cont_geral-1] == 'bilateral' or distribuicao[cont_geral-1] == 'bilateral frontal':
        img_bilateral = pyautogui.locateCenterOnScreen('bilateral.png', confidence =0.7)
        pyautogui.click(img_bilateral.x, img_bilateral.y)
        print("ENTROU NO bilateral")
        sleep(1)
        x_img = img_bilateral.x #posicao da distruibuicao
        y_img = img_bilateral.y
        tabs = 4 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)

    elif distribuicao[cont_geral-1]== 'bilateral_alternada' or distribuicao[cont_geral-1] == 'bilateral alternada':
        img_bilateral_alternada = pyautogui.locateCenterOnScreen('bilateral_alternada.png', confidence =0.7)
        pyautogui.click(img_bilateral_alternada.x, img_bilateral_alternada.y)
        sleep(1)
        x_img = img_bilateral_alternada.x #posicao da distruibuicao
        y_img = img_bilateral_alternada.y
        tabs = 3 #qtde de tabs para passar para o proximo
        teste_central(x_img, y_img, tabs)
        sleep(0.5)

    elif distribuicao[cont_geral-1] == 'central' or distribuicao[cont_geral-1]== 'canteiro central' or distribuicao[cont_geral-1]==  'canteiro_central' or distribuicao[cont_geral-1]== 'central':
        valida_central = 1
        img_central = pyautogui.locateCenterOnScreen('central.png', confidence =0.8)
        pyautogui.click(img_central.x, img_central.y)
        print("Entrou na distri canteiro central")
        sleep(1)     
        tab_interate(2)
    

    sleep(0.5)
    #Distancia entre postes
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

    #se for canteiro central inserir duas luminarias por poste
    if(valida_central == 1):
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(2))
    else:
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(1))
    sleep(0.6)
    #Angulo
    tab_interate(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str(angulo_x))
    
    if(valida_central == 0): #trava o pendor se nao for distri_central e insere valores para poste_pista e braco
        print("Entrou na validação")
        tab_interate(2)
        pyautogui.press('space')
        sleep(0.5)
        #distância poste-pista
        tab_interate(3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(poste_pista_x))
        #comprimento do braço
        tab_interate(2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(comprimento_braco_x))
    else: #se a distribuicao estiver no canteiro central, será necessário informar valor do pendor e até mesmo do deslocamento longitudinal
        tab_interate(3)     
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(pendor_x))
        #ADICIONAR AQUI DESLOCAMENTO LONGITUDINAL

    classifica_vias_passeios()
                                              
    if(check_passeio_adjacente == 0 or check_passeio_oposto == 0):
        print("entrou no IF que chama a funcao de exclusao")
        exclui_passeio(check_passeio_adjacente, check_passeio_oposto)

    #------------------------------------------#CHOOSE LUM-------------------------------------------#
    check_lum = []
    lum = ["AGN7026D4", "AGN7030D4", "AGN7040D4", "AGN7050D4", "AGN7060D4", "AGN7070D4", "AGN7080D4", "AGN7090D4", "AGN7100D4", "AGN7110D4", "AGN7120D4", "AGN7130D4", "AGN7150D4", "AGN7160D4", "AGN7170D4", "AGN7180D4", "AGN7200D4", "AGN7220D4", "AGN7240D4"]
    tamanho_lista_luminarias = len(lum)
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.7) #ir para ruas e voltar para luminarias para resetar tabs
    pyautogui.click(ruas.x, ruas.y)
    sleep(0.4)
    luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.6)
    pyautogui.click(luminaria.x, luminaria.y)
    sleep(0.4)
    primeira_luminaria = pyautogui.locateCenterOnScreen('primeira_luminaria.png', confidence=0.6)
    pyautogui.click(primeira_luminaria.x, primeira_luminaria.y)
    sleep(0.4)

    cont= 0
    agnes = 20 
    continua = False

    left = 1075
    top = 445
    width = 29
    height = 362
    # Capturar tela da área de resultados
    screenshot = pyautogui.screenshot(region=(left, top, width, height)) #captura checks
    # Especificar o caminho completo para salvar a captura de tela
    screenshotchecks = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/checks{cont}.png"  
    screenshot.save(screenshotchecks)    
    #screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/results_{cont}.png"
    cont = -1    
    cont_choose = 0
    while continua == False and cont_choose < tamanho_lista_luminarias:
        cont += 1   
        left = 1052
        top = 450
        width = 31
        height = 357
        sleep(0.5)
        # Capturar tela da área de resultados
        screenshot = pyautogui.screenshot(region=(left, top, width, height)) #captura checks
        # Especificar o caminho completo para salvar a captura de tela
        screenshotchecks = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/checks{cont}.png"  
        screenshot.save(screenshotchecks)    
        #screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/results_{cont}.png"
        if check_all(screenshotchecks, valida_central) ==True:
            continua = True
            check_lum.append(lum[cont])
        else:
            continua = False
            pyautogui.press('Down') #ir para proxima luminária 
            sleep(1)
        cont_choose +=1
        print("CONTADOR WHILE")
        print(cont_choose)
        print("tamanho lista lums")
        print(tamanho_lista_luminarias)
        atende = 1
        sleep(0.5)
        if(cont_choose == tamanho_lista_luminarias):
            print("nenhuma luminaria atende o cenario")
            check_lum = "NAO_ATENDE"
            atende = 0
            #adicionar aqui logica para salvar cenario sem luminarias

    #--------------------------------------------------------------------------------------------------#

    #modificando nome do projeto
    sleep(1.5)
    projecto = pyautogui.locateCenterOnScreen('guia_projecto.png', confidence=0.6)
    pyautogui.click(projecto.x, projecto.y)
    sleep(0.4)
    nome_projet = pyautogui.locateCenterOnScreen('nome_projeto.png', confidence=0.6)
    pyautogui.click(nome_projet.x, nome_projet.y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')

    if(atende == 1):
        luminaria_escolhida = str(check_lum[0])
    else:
        luminaria_escolhida = "NAO ATENDE"

    modify_name = "Itajai " + cont__str + " - " + luminaria_escolhida
    #modify_name.upper() 
     
    to_upper_safe(modify_name)
    pyautogui.write(modify_name.upper()) 

    #pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida))
    sleep(2.3)

    #Gerando relat - guia_documentacao
    guia_doc = pyautogui.locateCenterOnScreen('guia_documentacao.png', confidence=0.6)
    pyautogui.click(guia_doc.x, guia_doc.y)
    sleep(1.8)
    exibir = pyautogui.locateCenterOnScreen('exibir_doc.png', confidence=0.6)
    pyautogui.click(exibir.x, exibir.y)
    sleep(17)

    #salvando pdf relatório
    guardarpdf = pyautogui.locateCenterOnScreen('guardar_como.png', confidence=0.6)
    pyautogui.click(guardarpdf.x, guardarpdf.y)
    sleep(0.5)
    pdf_save = pyautogui.locateCenterOnScreen('pdf.png', confidence=0.6)
    pyautogui.click(pdf_save.x, pdf_save.y)
    sleep(0.8)
    ok_pdf = pyautogui.locateCenterOnScreen('ok_pdf.png', confidence=0.6)
    pyautogui.click(ok_pdf.x, ok_pdf.y)
    sleep(2)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.9)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.8)
    teste = pyautogui.locateCenterOnScreen('teste_pasta.png', confidence=0.6)
    pyautogui.doubleClick(teste.x, teste.y)
    sleep(1)
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(5.5)

    #salvando arquivo editável
    ficheiro = pyautogui.locateCenterOnScreen('ficheiro.png', confidence=0.7)
    pyautogui.click(ficheiro.x, ficheiro.y)
    sleep(0.5)
    guardar_project = pyautogui.locateCenterOnScreen('guardar_project.png', confidence=0.8)
    pyautogui.click(guardar_project.x, guardar_project.y)
    sleep(1.5)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.7)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.3)
    teste = pyautogui.locateCenterOnScreen('teste_pasta.png', confidence=0.6)
    pyautogui.doubleClick(teste.x, teste.y)
    sleep(0.5)
    nome_editavel = pyautogui.locateCenterOnScreen('nome_editavel.png', confidence=0.7)
    pyautogui.doubleClick(nome_editavel.x, nome_editavel.y)
    sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    project_name = "Itajai " + cont__str + " - " + luminaria_escolhida
    pyautogui.write(project_name.upper() + ".evo")
   # pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida + ".evo")) #nome arquivo
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(4)
    # Atualizar a planilha com a luminária escolhida e o ângulo
    df.at[idx, 'luminaria_escolhida'] = luminaria_escolhida
    df.at[idx, 'angulo_escolhido'] = angulo_x

    # Garantir que a coluna 'cenario' é do tipo object
    df['cenario'] = df['cenario'].astype(object)

    df.at[idx, 'cenario'] = modify_name

    # Identificar colunas a serem removidas
    colunas_remover = [col for col in df.columns if 'Material' in col]

    # Remover as colunas indesejadas
    df_limpo = df.drop(columns=colunas_remover)

    # Salvar a planilha atualizada  
    df_limpo.to_excel('table_itajai_test_atualizada.xlsx', sheet_name='RIAN - V4P4', index=False)
