import pytesseract
import cv2 #opencv

# links uteis:
# corrigir instalação windows: https://stackoverflow.com/questions/50951955/pytesseract-tesseractnotfound-error-tesseract-is-not-installed-or-its-not-i
# instalar outra língua: https://github.com/tesseract-ocr/tessdata
# pegar linguas: print(pytesseract.get_languages())

#ler imagem
imagem = cv2.imread("pista1.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

#pedir para o tesseract extrair o texto da imagem
texto = pytesseract.image_to_string(imagem)

print(texto)

