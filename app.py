import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os


cont_geral = 0
empasseio = 3.0
uopasseio = 0.20
check_distri = 0

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

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
distribuicao = df['distribuicao'].str.lower().tolist()      
# Converter a coluna 'qtde_faixas' para inteiros
# Preencher valores ausentes com 0 e converter a coluna 'qtde_faixas' para inteiros
df['qtde_faixas'] = df['qtde_faixas'].fillna(0).astype(int)
qtde_faixas = df['qtde_faixas'].tolist()
qtde_ruas = df['qtde_ruas'].tolist()
larg_canteiro_central = df['larg_canteiro_central'].tolist()
pendor = df['pendor'].tolist()


#------------ABRINDO CENARIO PADRAO ITAJAI-------------
# 1 - ABRINDO ARQUIVO
#pyautogui.doubleClick(147, 423, duration=0.5)
#sleep(30)  # TEMPO ATÉ ABRIR E CARREGAR O DIALUX

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
    if(validacao_central == 1):
        if num_checks >= 8:
            return True
        return False
    else:
        if num_checks >= 6:
            return True
        return False

# Função para verificar se os valores obtidos são suficientes
def check_results_passeio1_em():
    left = 990
    top = 552
    width = 51
    height = 22
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/resultsEm_{cont_geral}.png"
    screenshot.save(screenshot_path)
    results_text = pytesseract.image_to_string(screenshot_path, config='--psm 7').strip()
    print(f"Resultado antes da conversão para float: {results_text}")
    try:
        float_result = float(results_text)
        print(f"Resultado depois da conversão para float {float_result}")
    except ValueError:
        print(f"Erro ao converter Em result: '{results_text}'")
        float_result = float('inf')  # Usa um valor muito alto para evitar que este resultado seja selecionado
    return float_result

def check_results_passeio1_uo():
    left = 999
    top = 582
    width = 39
    height = 18
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/resultsUo_{cont_geral}.png"
    screenshot.save(screenshot_path)
    results_text = pytesseract.image_to_string(screenshot_path, config='--psm 7').strip()
    
    try:
        float_result = float(results_text)
    except ValueError:
        print(f"Erro ao converter Uo result: '{results_text}'")
        float_result = float('inf')  # Usa um valor muito alto para evitar que este resultado seja selecionado
    return float_result

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
for larg_passeio_oposto, larg_via, larg_passeio_adjacente, entre_postes_x, altura_lum_x, angulo_x, poste_pista_x, comprimento_braco_x, qtde_faixas_x, larg_canteiro_central_x, pendor_x in zip(larg_passeio_opost, largura_via, larg_passeio_adj, entre_postes, altura_lum, angulo, poste_pista, comprimento_braco, qtde_faixas, larg_canteiro_central, pendor):

    cont_geral += 1  # var para fazer a contagem de cenários 
    cont__str = str(cont_geral)  # var para fazer conversão de int para string e passar como parametro no nome do cenário
    
    # Abrindo guia planejamento
    pyautogui.click(399, 82, duration=0.5)
    sleep(1.5)
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.6)
    pyautogui.click(ruas.x, ruas.y)
    sleep(1)
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
            sleep(8)
            seta_baixo = pyautogui.locateCenterOnScreen('seta_baixo.png', confidence=0.9)
            pyautogui.click(seta_baixo.x, seta_baixo.y)
            sleep(1)

        #restante dos passos para inserir valores nos campos correspondentes 
        #PASSEIO1
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
        sleep(1.5)
        
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
    sleep(3)
    img_uni = pyautogui.locateCenterOnScreen('unilateral_inferior.png', confidence =0.7)
    pyautogui.click(img_uni.x, img_uni.y)
    sleep(3)

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
  
    elif distribuicao[cont_geral-1] == 'bilateral':
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
    

    sleep(3)
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

    #Angulo
    tab_interate(5)
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
    #------------------------------------------#CHOOSE LUM-------------------------------------------#
    check_lum = []
    lum = ["AGN7026D4", "AGN7030D4", "AGN7040D4", "AGN7050D4", "AGN7060D4", "AGN7070D4", "AGN7080D4", "AGN7090D4", "AGN7100D4", "AGN7110D4", "AGN7120D4", "AGN7130D4", "AGN7150D4", "AGN7160D4", "AGN7170D4", "AGN7180D4", "AGN7200D4", "AGN7220D4", "AGN7240D4"]
    tamanho_lista_luminarias = len(lum)
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.6) #ir para ruas e voltar para luminarias para resetar tabs
    pyautogui.click(ruas.x, ruas.y)
    sleep(0.4)
    luminaria = pyautogui.locateCenterOnScreen('luminaria.png', confidence=0.6)
    pyautogui.click(luminaria.x, luminaria.y)
    sleep(0.4)
    primeira_luminaria = pyautogui.locateCenterOnScreen('primeira_luminaria.png', confidence=0.6)
    pyautogui.click(primeira_luminaria.x, primeira_luminaria.y)
    sleep(0.4)
    #tab_interate(9)
    #pyautogui.press('Down') #ir para proxima luminária
    # Verificar resultados
    #if check_results_passeio1_em() == True and check_results_passeio1_uo() == True:
     #   print(f"Luminária selecionada no cenário {cont__str} atende aos requisitos.")
    #else:
     #   print(f"Luminária selecionada no cenário {cont__str} não atende aos requisitos.")

    cont= 0
    agnes = 20 
    continua = False

    left = 1041
    top = 550
    width = 25
    height = 255
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
        if(cont_choose == tamanho_lista_luminarias):
            print("nenhuma luminaria atende o cenario")
            check_lum = "[NAO_ATENDE]"
            #adicionar aqui logica para salvar cenario sem luminarias

    '''
    while cont < agnes:
        left = 1041
        top = 550
        width = 25
        height = 255
        # Capturar tela da área de resultados
        screenshot = pyautogui.screenshot(region=(left, top, width, height)) #captura checks
        # Especificar o caminho completo para salvar a captura de tela
        screenshotchecks = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/checks{cont}.png"  
        screenshot.save(screenshotchecks)    
        #screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/results_{cont}.png"
        
        if check_all(screenshotchecks) ==True:
            check_lum.append(lum[cont]) #coloca na lista as luminárias que atenderam os 6 checks
            break
        print(check_lum)
        pyautogui.press('Down') #ir para proxima luminária 
        sleep(1)
        cont += 1 
        
   

    if best_lum:
        print(f"A luminária mais eficiente é: {best_lum} com Em: {best_em_result} e Uo: {best_uo_result}")
    else:
        print("Nenhuma luminária atende aos cenários.")
    ''' 
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
    luminaria_escolhida = str(check_lum)
    modify_name = "Itajai " + cont__str + " " + luminaria_escolhida
    pyautogui.write(modify_name.upper())  
    #pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida))
    sleep(2.3)

    #Gerando relat - guia_documentacao
    guia_doc = pyautogui.locateCenterOnScreen('guia_documentacao.png', confidence=0.6)
    pyautogui.click(guia_doc.x, guia_doc.y)
    sleep(1.5)
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
    sleep(1)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.5)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.8)
    teste = pyautogui.locateCenterOnScreen('teste_pasta.png', confidence=0.6)
    pyautogui.doubleClick(teste.x, teste.y)
    sleep(0.5)
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(4)

    #salvando arquivo editável
    ficheiro = pyautogui.locateCenterOnScreen('ficheiro.png', confidence=0.5)
    pyautogui.click(ficheiro.x, ficheiro.y)
    sleep(0.5)
    guardar_project = pyautogui.locateCenterOnScreen('guardar_project.png', confidence=0.8)
    pyautogui.click(guardar_project.x, guardar_project.y)
    sleep(0.5)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.6)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.3)
    teste = pyautogui.locateCenterOnScreen('teste_pasta.png', confidence=0.6)
    pyautogui.doubleClick(teste.x, teste.y)
    sleep(0.5)
    nome_editavel = pyautogui.locateCenterOnScreen('nome_editavel.png', confidence=0.4)
    pyautogui.doubleClick(nome_editavel.x, nome_editavel.y)
    sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    project_name = "Itajai " + cont__str + " " + luminaria_escolhida
    pyautogui.write(project_name.upper() + ".evo")
   # pyautogui.write(str("Itajai " + cont__str + " " + luminaria_escolhida + ".evo")) #nome arquivo
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(4)


   
    #fazer a comparação de qual luminaria é a mais eficiente, para isso vamos tirar print dos resultados, extrair o texto das imagens 
    # e fazer uma comparação pra ver qual esta mais próximo do resultado. https://awari.com.br/ocr-em-python-aprenda-a-extrair-texto-de-imagens-com-facilidade/
    #ideia de comparação: extrair qual a classe da via e gerar um script que gera automaticamente uma planilha com os parametros para comparação
