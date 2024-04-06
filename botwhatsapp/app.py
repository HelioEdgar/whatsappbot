import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/')
sleep(20)

# Ler planilha e guardar informações sobre nome e telefone
workbook = openpyxl.load_workbook('nomes.xlsx')
nomes_placa = workbook['Folha1']

for i in range(1, 20):
    for linha in nomes_placa.iter_rows(min_row=2):
        #nome, telefone
        nome = linha[0].value
        telefone = linha[1].value
        mensagem = f'{nome}, Why are you sad?'

        # Criar links personalizados do whatsapp e enviar mensagens para cada nome com base nos dados da planilha
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl','w')
            sleep(5)
        except:
            print(f'Não foi possivel enviar mensagem ao {nome}')
            with open('erros.csv', 'a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone},{os.linesep}')
    