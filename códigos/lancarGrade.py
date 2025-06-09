import pyautogui
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para carregar a planilha
def carregar_planilha():
    global grade, codigos, quantidades  # Garantir que estamos alterando as variáveis globais
    arquivo = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        try:
            # Carregar a planilha e exibir algumas informações
            grade = pd.read_excel(arquivo)

            # Verificar se as colunas necessárias estão presentes
            if "Produto" not in grade.columns or "Qtde" not in grade.columns:
                messagebox.showerror("Erro", "A planilha não contém as colunas 'Produto' e 'Qtde'.")
                return

            codigos = grade.loc[grade["Produto"] > 0, "Produto"]
            quantidades = grade.loc[grade["Qtde"] > 0, "Qtde"]

            # Exibir informações na interface
            label_informacoes.config(text=f"Planilha carregada com {len(codigos)} produtos.")
            botao_iniciar.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
    else:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
    return

# Função para iniciar o processo de automação
def iniciar_automacao():
    if grade is None or codigos is None or quantidades is None:
        messagebox.showwarning("Aviso", "Carregue uma planilha primeiro!")
        return

    # Avisar o usuário para garantir que o Promax esteja ativo
    mensagem_promax = messagebox.askyesno("Aviso", "Por favor, verifique se o Promax está no primeiro plano (pressione ALT+TAB para isso). Deseja continuar?")
    if not mensagem_promax:
        return

    try:
        # Abrir a rotina no promax
        pyautogui.hotkey("alt", "tab")
        time.sleep(2)
        pyautogui.write('020301')
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.write('1')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

        # Loop para lançar os produtos na grade
        for i in range(len(codigos)):
            pyautogui.write(str(codigos[i]))
            pyautogui.press("tab")
            time.sleep(1.5)
            pyautogui.write(str(quantidades[i]))
            pyautogui.press("tab")
            time.sleep(1.5)
            pyautogui.hotkey("alt", "s")
            time.sleep(3)
        
        messagebox.showinfo("Sucesso", "Automação concluída com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante a automação: {e}")

# Configurar a interface gráfica
root = tk.Tk()
root.title("Automação de Produtos")
root.geometry("400x350")

# Variáveis globais
grade, codigos, quantidades = None, None, None

# Label para mostrar informações da planilha
label_informacoes = tk.Label(root, text="Nenhuma planilha carregada.")
label_informacoes.pack(pady=20)

# Aviso sobre o Promax
aviso_promax = tk.Label(root, text="** Garanta que o Promax esteja no primeiro programa (ALT+TAB) antes de iniciar a automação. **", fg="red", wraplength=350)
aviso_promax.pack(pady=10)

# Botão para carregar a planilha
botao_carregar = tk.Button(root, text="Carregar Planilha", command=lambda: carregar_planilha())
botao_carregar.pack(pady=10)

# Botão para iniciar a automação (inicialmente desabilitado)
botao_iniciar = tk.Button(root, text="Iniciar Automação", command=iniciar_automacao, state=tk.DISABLED)
botao_iniciar.pack(pady=10)

# Rodar a interface
root.mainloop()
