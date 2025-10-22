import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep
import os

workbook = openpyxl.load_workbook('Contato.xlsx')
Pagina_cliente = workbook['Planilha4']

for linha in Pagina_cliente.iter_rows(min_row=2):
    #nome, responsavel, numero
    nome = linha[0].value
    responsavel = linha[1].value
    numero = linha[2].value
    
    messagem = f'Olá {responsavel}.\nDesculpa o incomodo, Eu me chamo Felipe, sou fotográfo de formatura.\nO motivo do meu ctt é que tirei fotos do aluno(a) {nome}, no Cea Office - Escola de Informática Prossionalizante, fotos padrão formatura, como beca, capelo, canudo e todos os paramentos, se tiver interesse me retorne e passarei os detalhes! Grato e bom dia' 
    
    try:
      link_message_whats = f'https://web.whatsapp.com/send?phone={numero}&text={quote(messagem)}'
      webbrowser.open(link_message_whats)
      sleep(10)
      pyautogui.press('enter')
      sleep(9)
      
      pyautogui.hotkey('ctrl','w')
      sleep(7)
