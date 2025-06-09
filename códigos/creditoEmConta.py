import pandas as pd
import pyautogui as pa
from time import sleep
time = 1.2 


def alterar(mapa, nota, forma):
    sleep(1.8)
    pa.write(mapa)
    pa.press('tab')
    pa.write(nota)    
    pa.press('tab')
    pa.write('003') 
    pa.press('tab')
    sleep(time)
    pa.hotkey('ctrl', 'a')
    pa.write(forma)
    pa.press('tab') 
    sleep(time)
    pa.press('ENTER')
    sleep(time)
    pa.press('ENTER')

    
file_path = r'C:\Users\caixa.patos\Desktop\Utilidades.xlsm'

df = pd.read_excel(file_path, sheet_name="CEMC", engine='openpyxl', header = 0, dtype=str)

pa.FAILSAFE = False
 
# df = pd.read_excel(file_path, sheet_name="NotasDoCliente", engine='openpyxl', header = 1, dtype=str)

#DEFINIR FORMA DE PAGAMENTO
forma = str(10)

pa.hotkey('alt', 'tab') 
sleep(1)
cont = len(df['NOTA'])


for nota, mapa in zip(df['NOTA'], df['MAPA']):
    print(f"faltam {cont}")
    cont -= 1
    alterar(mapa, nota, forma)