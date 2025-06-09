import os
import pandas as pd
import pyautogui as pa
from time import sleep


file_path = r'C:\Users\caixa.patos\Downloads\alterar.xlsx'

pa.PAUSE = 1.5

def alterar(nb):
    pa.write(str(nb))
    pa.press('tab')
    sleep(1.5)
    pa.write(str(2))
    sleep(1.5)
    pa.press('tab')
    sleep(0.8)
    pa.hotkey('alt', 's')
    sleep(1.5)


df = pd.read_excel(file_path, sheet_name="alterar", engine='openpyxl', dtype=str)

cont = len(df['NB'])

pa.hotkey('alt', 'tab')

for nb in df['NB'].dropna():
    print(cont)
    cont -= 1
    alterar(nb)