import pyautogui as pa
from time import sleep

pa.hotkey('alt', 'tab') 

for i in range(50):
    sleep(1.0)    
    pa.press('down')            
    sleep(0.8)  
    pa.press('tab')
    sleep(0.8) 
    pa.press('ENTER')   
