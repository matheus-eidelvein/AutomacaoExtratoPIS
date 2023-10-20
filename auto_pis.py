import pyautogui
import openpyxl
from time import sleep


planilha_grf = openpyxl.load_workbook(
    r"Caminho/Para/Seu/Arquivo/Excel.xlsx") # Caminho indicando o nome da Planilha que deseja exportar o dado requesitado
sheet_pis = planilha_grf["EXTRATO"] # Nome da página dentro da planilha
for linha in sheet_pis.iter_rows(min_row=2, max_row=102): # Mínimo e máximo de linhas, do primeiro dado até o último
    linha_pis = linha[3].value # N° relacionado a coluna onde se localiza os dados(PIS) 
    pyautogui.click(515, 611, duration=1)
    pyautogui.click(511, 711, duration=1)
    pyautogui.click(546, 687, duration=1)
    pyautogui.write(str(linha_pis)) # Reescreve o dado que está na planilha 
    pyautogui.click(300, 829, duration=1)
    pyautogui.click(293, 786, duration=1)

# A definição das coordenadas pode ser alterada de acordo com a interface do site.
