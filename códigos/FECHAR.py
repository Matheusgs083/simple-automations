import pandas as pd
import pyautogui as pa
from time import sleep
time = 2.0


def digitar(cx, uni):
    pa.write(cx)
    pa.press('tab')
    pa.write(uni)
    pa.press('tab')


def digitarDev(cx, uni):
    pa.press('tab')
    pa.press('tab')
    pa.write(cx)
    pa.press('tab')
    pa.write(uni)
    pa.press('tab')


file_path = r'C:\Users\caixa.patos\Desktop\Utilidades.xlsm'

df= pd.read_excel(file_path, sheet_name="TROCAS", engine='openpyxl', header = 1, dtype=str)
df.raw = pd.read_excel(file_path, sheet_name="TROCAS", engine='openpyxl', header=None)

pa.FAILSAFE = False
 

pa.hotkey('alt', 'tab') 
sleep(1)

dev = df.raw.iloc[10, 7]

if dev == "ok":
    for cx, uni in zip(df['cx'], df['uni']):
        digitar(cx, uni)

else:
    for cx, uni in zip(df['cx'], df['uni']):
        digitarDev(cx, uni)