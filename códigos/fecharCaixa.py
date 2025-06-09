import pyautogui as pa
from time import sleep
import tkinter as tk
from tkinter import simpledialog, messagebox
pa.PAUSE = 1.2


def transferencias():
    sleep(0.5)   # COLAR TRANSFERENCIAS
    pa.click(946,351)
    sleep(2)
    pa.press('ENTER')
    pa.click(459,472)
    pa.press('alt')
    pa.press('c')
    for i in range(2):
        pa.press('v')
def despesas():
    pa.click(511,384) #click no botao dos recibos
    pa.click(638,294) # click para copoar
    sleep(1.5)
    pa.press('ENTER')
    pa.click(747,270) # colar
    pa.press('alt')
    pa.press('c')
    for i in range(2):
        pa.press('v')
    pa.click(638,332) # imprimir
    pa.press('ENTER')
    pa.click(632,259) # vai pro resumo
    pa.click(628,474) # volta po caixa

# Pergunta ao usuário se deseja executar as funções
root = tk.Tk()
root.withdraw()  # Esconde a janela principal

resposta = messagebox.askyesno(
    "Fechar o Caixa",
    "Deseja fechar o caixa? Lembre-se que para fechar o caixa, você deve estar na planilha do caixa."
)

if resposta:  # Se o usuário clicar em "Sim"
    transferencias()
    despesas()
    messagebox.showinfo("Finalizado", "O caixa foi fechado com sucesso.")
else:
    messagebox.showinfo("Operação Cancelada", "O fechamento do caixa foi cancelado.")