import os
import pandas as pd
import pyautogui as pa
from time import sleep
import tkinter as tk
from tkinter import messagebox, filedialog

# Definição de pausa padrão
pa.PAUSE = 2.0

def alterar(mapa, nota, forma): 
    pa.write(mapa)
    pa.press('tab')
    pa.write(nota)    
    pa.press('tab')
    pa.write('003')
    pa.press('tab')
    pa.write(forma)
    pa.press('tab')
    pa.press('ENTER')
    pa.press('ENTER')

def iniciar_processo():
    
    try:
        velocidade = float(entry_velocidade.get())
    except ValueError:
        velocidade = 0.5
    
    time = velocidade
    pa.hotkey('alt', 'tab')
    sleep(2)
    
    try:
        
        df = pd.read_excel(file_path.get(), engine='openpyxl', header=0, dtype=str)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo Excel: {e}")
        return

    forma = str(10)
    
    for nota, mapa in zip(df['NOTA'], df['MAPA']):
        alterar(mapa, nota, forma)

def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xls;*.xlsx;*.xlsm")],
        initialdir=os.path.expanduser("~")
    )
    if caminho_arquivo:
        file_path.set(caminho_arquivo)
        lbl_caminho_arquivo.config(text=f"Arquivo selecionado: {caminho_arquivo}")


root = tk.Tk()
root.title("Financeiro Patos")

window_width = 800
window_height = 250

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x_position = (screen_width // 2) - (window_width // 2)
y_position = (screen_height // 2) - (window_height // 2)

root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

file_path = tk.StringVar() 

# Mensagem inicial
msg = tk.Label(
    root, 
    text="Certifique-se de que a rotina '03030701' esteja aberta!\n"
         "E seja o primeiro no ALT+TAB antes de continuar, caso contrário o programa não funcionará corretamente.\n"
         "Qualquer dúvida, fale com o Financeiro Patos.",
    font=("Arial", 12), 
    justify="center", 
    padx=10, 
    pady=10
)
msg.pack()

# Botão para selecionar o arquivo
btn_selecionar = tk.Button(
    root, 
    text="Selecionar Planilha", 
    command=selecionar_arquivo, 
    font=("Arial", 12), 
    padx=10, 
    pady=5
)
btn_selecionar.pack()

# Label para mostrar o caminho do arquivo selecionado
lbl_caminho_arquivo = tk.Label(
    root, 
    text="Nenhum arquivo selecionado", 
    font=("Arial", 10), 
    wraplength=600
)
lbl_caminho_arquivo.pack()

# Entrada para velocidade
lbl_velocidade = tk.Label(
    root, 
    text="Digite a velocidade entre os comandos (para dias com o promax mais lento) (em segundos ex: 0.8, 1, 2):", 
    font=("Arial", 12)
)
lbl_velocidade.pack()

entry_velocidade = tk.Entry(root, font=("Arial", 12))
entry_velocidade.pack()
entry_velocidade.insert(0, "0.5")

btn_iniciar = tk.Button(
    root, 
    text="OK", 
    command=iniciar_processo, 
    font=("Arial", 12), 
    padx=10, 
    pady=5
)
btn_iniciar.pack()
    
root.mainloop()

#MATHEUS CAIXA NOTURNO