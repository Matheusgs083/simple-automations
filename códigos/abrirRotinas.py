import pyautogui as pa
from time import sleep

def abrirRotina(codigo):
    pa.click(1213, 179, duration= 1) 
    pa.write(codigo)
    pa.press('ENTER')
    pa.click(764,81, duration = 0.5)

def abrirRotinaPersonalizado(codigo, inicial, final):
    pa.click(1213, 179, duration= 1) 
    pa.write(codigo)
    pa.press('ENTER')
    pa.click(inicial, final,  duration = 0.5)

def clicar_arrastar(ponto_inicial, ponto_final):
    pa.click(ponto_inicial[0], ponto_inicial[1], duration = 0.5)  # Clique no ponto inicial
    pa.mouseDown()                                             # Pressiona o botão do mouse
    pa.moveTo(ponto_final[0], ponto_final[1], duration= 1) # Move a1té o ponto final
    pa.mouseUp()      

pa.PAUSE = 0.7

codigos = ["120701", "03030701", "030303", "030330", "030332", "03030702", "030302"]
#codigos das rotinas, para adicionar, basta colocar na lista

for i in codigos:
    if i == "120701": # ajuste na 120701, por conta do tamanho diferente da aba
        abrirRotinaPersonalizado(i, 981,80)
    elif i =="030302": # caso especial, pois precisa mover a janela
        abrirRotina(i)
        clicar_arrastar((379, 13), (970, 137))
        clicar_arrastar((568, 120), (513, 91))
    else:
        abrirRotina(i)
    pa.hotkey('alt' , 'tab')




