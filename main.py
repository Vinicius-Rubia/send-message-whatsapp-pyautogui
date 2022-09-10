import os
import pandas as pd
import pyautogui
import time
import numpy as np
from math import floor
from unidecode import unidecode

pyautogui.press('winleft')
time.sleep(1)
pyautogui.write('chrome')
time.sleep(1)
pyautogui.press('enter')
time.sleep(5)
pyautogui.write('gmail.com')
time.sleep(1)
pyautogui.press('enter')
time.sleep(10)
pyautogui.press('enter')
time.sleep(10)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(30)

relatorio = pd.read_excel(r"SALDO.xlsx")
relatorio = relatorio.drop(['SALDO', 'Unnamed: 2','Unnamed: 3','Unnamed: 4', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9'], axis=1)
relatorio = relatorio.drop(0)
relatorio = relatorio.drop(1)
relatorio.columns = relatorio.columns.str.lower().str.replace('unnamed: 1', 'conta')
relatorio.columns = relatorio.columns.str.lower().str.replace('unnamed: 5', 'saldo')
relatorio.columns = relatorio.columns.str.lower().str.replace('unnamed: 6', 'diaria')
relatorio = relatorio.replace({'--': np.nan})
relatorio = relatorio.replace({'R$': ''})
relatorio['diaria'] = pd.to_numeric(relatorio['diaria'], errors='coerce')
relatorio = relatorio.dropna(how='any', axis=0)

relatorio.to_excel('nova.xlsx')

tabela = pd.read_excel("nova.xlsx")

contact = 'contato'

pyautogui.press('winleft')
time.sleep(1)
pyautogui.write('chrome')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.write('https://web.whatsapp.com/')
time.sleep(1)
pyautogui.press('enter')
time.sleep(7)
pyautogui.hotkey('ctrl', 'alt', '/')
time.sleep(1)
pyautogui.write(contact)
time.sleep(1)
pyautogui.press('enter')
time.sleep(2)
pyautogui.write('Segue o Relatorio de saldo')
time.sleep(1)
pyautogui.press('enter')
time.sleep(2)

for i, conta in enumerate(tabela['conta'].str.upper()):
    saldo = tabela.loc[i, 'saldo']
    diaria = tabela.loc[i, 'diaria']
    custo_semanal = saldo/diaria
    new_custo_mensal = floor(custo_semanal)

    if custo_semanal < 7:
        message1 = f'{unidecode(conta)}'
        message2 = 'Saldo insuficiente para 7 dias'
        message3 = f'Valido para: {new_custo_mensal} dias'
        message4 = f'Saldo Restante R${saldo} reais'
        message5 = f'Custo Diario R${diaria} reais'

        pyautogui.write(message1)
        pyautogui.hotkey('shift', 'enter')
        time.sleep(1)
        pyautogui.write(message2)
        pyautogui.hotkey('shift', 'enter')
        time.sleep(1)
        pyautogui.write(message3)
        pyautogui.hotkey('shift', 'enter')
        time.sleep(1)
        pyautogui.write(message4)
        pyautogui.hotkey('shift', 'enter')
        time.sleep(1)
        pyautogui.write(message5)
        pyautogui.press('enter')
        time.sleep(5)
        
os.remove('nova.xlsx')
os.remove(r"SALDO.xlsx")
time.sleep(3)
pyautogui.hotkey('alt', 'f4')
time.sleep(1)
pyautogui.hotkey('alt', 'f4')