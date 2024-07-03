import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np


cont_geral = 0

empasseio = 3.0
uopasseio = 0.20

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

#------------ABRINDO CENARIO PADRAO ITAJAI-------------
# 1 - ABRINDO ARQUIVO
#pyautogui.doubleClick(147, 423, duration=0.5)
#sleep(30)  # TEMPO ATÉ ABRIR E CARREGAR O DIALUX

def check_all(screenshot_path):
    image = cv2.imread(screenshot_path)
    hsv_image = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    
    # Define range for green color in HSV
    lower_green = np.array([40, 40, 40])
    upper_green = np.array([80, 255, 255])
    
    mask = cv2.inRange(hsv_image, lower_green, upper_green)
    green_pixels = cv2.countNonZero(mask)
    
    # Assumindo que cada "check" verde ocupa um número significativo de pixels verdes
    if green_pixels >= 6 * 100:  # Ajuste 100 conforme necessário para corresponder ao tamanho dos "checks"
        return True
    return False


# Função para verificar se os valores obtidos são suficientes
def check_results_passeio1_em():
    left = 992
    top = 477
    width= 69
    height =29
    # Capturar tela da área de resultados
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    
    # Especificar o caminho completo para salvar a captura de tela
    screenshot_path = "C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/resultsEm.png"  
    screenshot.save(screenshot_path)
    # Usar OCR para extrair texto da imagem
    results_text = pytesseract.image_to_string(screenshot_path)
    float_result = float(results_text)
    # Verificar se os valores atendem aos requisitos
    if float_result >= empasseio:
        return True
    return False

def check_results_passeio1_uo():
    left = 999
    top = 508
    width= 64
    height =23
    # Capturar tela da área de resultados
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    
    # Especificar o caminho completo para salvar a captura de tela
    screenshot_path = "C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/resultsUo.png"  
    screenshot.save(screenshot_path)
    
    # Usar OCR para extrair texto da imagem
    results_text = pytesseract.image_to_string(screenshot_path)
    float_result = float(results_text)
    # Verificar se os valores atendem aos requisitos
    if float_result >= uopasseio:
        return True
    return False

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
    
    # Abrindo guia planejamento
    pyautogui.click(399, 82, duration=0.5)
    sleep(1.5)
    
    ruas = pyautogui.locateCenterOnScreen('ruas.png', confidence=0.6)
    pyautogui.click(ruas.x, ruas.y)
    sleep(0.5)

    # Selecionando o passeio1
    passeio1 = pyautogui.locateCenterOnScreen('passeio1.png', confidence=0.8)
    pyautogui.doubleClick(passeio1.x, passeio1.y)
    sleep(1)
    tab_interate(3)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_opost
    pyautogui.write(str(larg_passeio_oposto))
    sleep(0.8)

    ##--------------------- PARAMETROS RUA ---------------------
    # Clicando no campo largura via (ajustar coordenadas conforme necessário)
    pista1 = pyautogui.locateCenterOnScreen('pista1.png', confidence=0.8)
    pyautogui.doubleClick(pista1.x, pista1.y)
    sleep(1)
    tab_interate(6)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para largura_via
    pyautogui.write(str(larg_via))
    sleep(0.8)

    ##--------------------- PARAMETROS PASSEIO ADJACENTE ---------------------
    # Selecionando o passeio2
    passeio2 = pyautogui.locateCenterOnScreen('passeio2.png', confidence=0.8)
    pyautogui.doubleClick(passeio2.x, passeio2.y)
    sleep(1)
    tab_interate(3)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_adjacente
    pyautogui.write(str(larg_passeio_adjacente))
    sleep(0.8)

    '''
    # Selecionando o passeio1
    pyautogui.doubleClick(102, 586, duration=0.8)
    sleep(0.8)
    # Clicando no campo largura passeio
    pyautogui.doubleClick(269, 754, duration=0.8)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_opost
    pyautogui.write(str(larg_passeio_oposto))
    sleep(0.8)

    ##--------------------- PARAMETROS RUA ---------------------
    # Clicando no campo largura via (ajustar coordenadas conforme necessário)
    pyautogui.doubleClick(118, 621, duration=0.8)  # clicando na rua
    sleep(0.8)
    pyautogui.doubleClick(262, 820, duration=0.8)  # clicando no campo de largura
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para largura_via
    pyautogui.write(str(larg_via))
    sleep(0.8)

    ##--------------------- PARAMETROS PASSEIO ADJACENTE ---------------------
    # Selecionando o passeio2
    pyautogui.doubleClick(112, 651, duration=0.8)
    sleep(0.8)
    # Clicando no campo largura passeio
    pyautogui.doubleClick(244, 754, duration=0.8)
    # Selecionar todo o texto existente e apagar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    # Digitando o novo valor para larg_passeio_adjacente
    pyautogui.write(str(larg_passeio_adjacente))
    sleep(0.8)
    '''
    
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


    #------------------------------------------#CHOOSE LUM-------------------------------------------#
    check_lum = []
    lum = ["lum1", "lum2", "lum3", "lum4", "lum5"]
    
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
    pyautogui.press('Down') #ir para proxima luminária
    # Verificar resultados
    if check_results_passeio1_em() == True and check_results_passeio1_uo() == True:
        print(f"Luminária selecionada no cenário {cont__str} atende aos requisitos.")
    else:
        print(f"Luminária selecionada no cenário {cont__str} não atende aos requisitos.")

        cont= 0
        agnes = 20 
        while cont < agnes:
            screenshot_path = f"C:/Users/AdminDell/Desktop/SCREENSHOTS_RESULTS/results_{cont}.png"
            if check_all(screenshot_path):
                check_lum.append(lum[cont])
            cont += 1

        if check_lum:
            best_lum = min(check_lum, key=lambda x: float(x.split()[-1]))
            print(f"A luminária mais eficiente é: {best_lum}")
        else:
            print("Nenhuma luminária atende aos cenários.")



    #--------------------------------------------------------------------------------------------------#

    #modificando nome do projeto
    sleep(1)
    projecto = pyautogui.locateCenterOnScreen('guia_projecto.png', confidence=0.6)
    pyautogui.click(projecto.x, projecto.y)
    sleep(0.4)
    nome_projet = pyautogui.locateCenterOnScreen('nome_projeto.png', confidence=0.6)
    pyautogui.click(nome_projet.x, nome_projet.y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(str("Itajai " + cont__str))
    sleep(0.4)

    #Gerando relat - guia_documentacao
    guia_doc = pyautogui.locateCenterOnScreen('guia_documentacao.png', confidence=0.6)
    pyautogui.click(guia_doc.x, guia_doc.y)
    sleep(1.5)
    exibir = pyautogui.locateCenterOnScreen('exibir_doc.png', confidence=0.6)
    pyautogui.click(exibir.x, exibir.y)
    sleep(11)

    #salvando pdf relatório
    guardarpdf = pyautogui.locateCenterOnScreen('guardar_como.png', confidence=0.6)
    pyautogui.click(guardarpdf.x, guardarpdf.y)
    sleep(0.3)
    pdf_save = pyautogui.locateCenterOnScreen('pdf.png', confidence=0.6)
    pyautogui.click(pdf_save.x, pdf_save.y)
    sleep(0.5)
    ok_pdf = pyautogui.locateCenterOnScreen('ok_pdf.png', confidence=0.6)
    pyautogui.click(ok_pdf.x, ok_pdf.y)
    sleep(0.3)
    documentos = pyautogui.locateCenterOnScreen('documentos_w11.png', confidence=0.6)
    pyautogui.click(documentos.x, documentos.y)
    sleep(0.5)
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
    pyautogui.write(str("Itajai " + cont__str + ".evo")) #nome arquivo
    salvar_pasta = pyautogui.locateCenterOnScreen('salvar_pasta.png', confidence=0.6)
    pyautogui.click(salvar_pasta.x, salvar_pasta.y)
    sleep(4)


   
    #fazer a comparação de qual luminaria é a mais eficiente, para isso vamos tirar print dos resultados, extrair o texto das imagens 
    # e fazer uma comparação pra ver qual esta mais próximo do resultado. https://awari.com.br/ocr-em-python-aprenda-a-extrair-texto-de-imagens-com-facilidade/
    #ideia de comparação: extrair qual a classe da via e gerar um script que gera automaticamente uma planilha com os parametros para comparação
    #no dialux criar guias como foram feitos no projeto de cachoeiro, fazer comparação tipo aquela
    #na funcao de salvar as imagens, precisa concatenar um numero para que as imagens nao fiquem com o mesmo nome e acabem dando problma