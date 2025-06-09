import pyautogui as pa
from time import sleep
import tkinter as tk
from tkinter import simpledialog, messagebox

def abrirPromax():
    pa.click(463, 747) # abrir promax na barra de tarefas
    sleep(7)
    pa.click(1335, 88) # fechar coisa irritante
    sleep(2)
    pa.click(726,253) # login
    sleep(2)
    pa.click(715, 297) # usuario e senha
    sleep(1.5)
    pa.click(681, 317)
    sleep(2)
    pa.click(697,309)

def inputIntervaloDeMapas():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    inicial = simpledialog.askstring("Input", "Digite o mapa inicial:")
    final = simpledialog.askstring("Input", "Digite o mapa final:")
    return inicial, final

def imprimirViaCega():
    pa.PAUSE = 1

    inicial, final = inputIntervaloDeMapas() 

    pa.click(1213, 179, duration = 1)  # digitar rotina
    pa.write("030232")
    pa.press('ENTER')
    sleep(1)
    pa.press('down')
    pa.click(625, 250, duration = 1) # selecionar intervalo de mapas
    pa.write(inicial)
    pa.press('tab')
    pa.write(final)
    pa.click(726,542, duration = 1) # visualizar mapa
    pa.click(765,78, duration = 1) # fecha coisa irritante
    pa.click(724,15, duration = 1) # deixar em tela cheia 
    pa.hotkey('alt', 'i') # imprimir
    pa.hotkey('alt', 'tab')  # voltar ao menu inicial
    pa.click(1213, 179, duration= 1)
    pa.write("03031902") # entra na verificação de troca
    pa.press('ENTER')
    pa.click(764,81, duration = 0.5)

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

def perguntar_ao_usuario():
    root = tk.Tk()
    root.title("Escolha a Ação")
    root.geometry("300x200")
    
    # Variável para armazenar a escolha
    escolha = tk.StringVar(value="0")

    # Função para finalizar e retornar o valor selecionado
    def finalizar():
        root.quit()

    # Adicionando os botões de escolha
    tk.Label(root, text="O que você gostaria de fazer?", pady=20).pack()

    tk.Radiobutton(root, text="Abrir as rotinas", variable=escolha, value="abrir").pack(anchor="w")
    tk.Radiobutton(root, text="Imprimir as vias cegas", variable=escolha, value="imprimir").pack(anchor="w")
    tk.Radiobutton(root, text="Fazer ambos", variable=escolha, value="ambos").pack(anchor="w")

    # Botão para finalizar a escolha
    tk.Button(root, text="Confirmar", command=finalizar).pack(pady=20)
    
    root.mainloop()
    
    return escolha.get()

def perguntar_se_abrir_promax():
    resposta = messagebox.askquestion("Abrir ProMax?", "Você deseja abrir o ProMax?")
    return resposta


def executar_rotinas():
    pa.PAUSE = 0.7

    codigos = ["120701", "030345", "03030701", "030303", "030330", "030332", "03030702", "0303020   1"]
    # Códigos das rotinas, para adicionar, basta colocar na lista

    for i in codigos:
        if i == "120701":  # Ajuste para o código 120701 devido a diferenças na interface
            abrirRotinaPersonalizado(i, 981, 80)
        elif i == "030302":  # Caso especial, onde é necessário mover a janela
            abrirRotina(i)
            clicar_arrastar((379, 13), (970, 137))  # Mover a janela
            clicar_arrastar((568, 120), (513, 91))  # Outro movimento da janela
        else:
            abrirRotina(i)  # Para outros códigos, apenas abre a rotina
        pa.hotkey('alt', 'tab')  # Alterna entre as janelas

def main():
    
    resposta_promax = perguntar_se_abrir_promax()

    if resposta_promax == "yes":  # Se o usuário escolher "Sim"
        abrirPromax()
        sleep(2)

    resposta = perguntar_ao_usuario()
    
    if resposta == 'ambos':  # Caso o usuário queira abrir as rotinas e imprimir
        executar_rotinas()
        imprimirViaCega()
    elif resposta == 'imprimir':  # Caso o usuário só queira imprimir as vias cegas
        imprimirViaCega()
    elif resposta == 'abrir':  # Caso o usuário só queira abrir as rotinas
        executar_rotinas()


sleep(2)
main()
