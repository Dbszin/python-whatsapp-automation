import pandas as pd
import pywhatkit as kit
import time
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from threading import Thread
import os

# Configurações globais
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1
dados = None
progresso = 0
running = False
imagem_path = None

# Função para carregar o arquivo Excel
def carregar_arquivo():
    global dados, progresso
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if arquivo:
        try:
            dados = pd.read_excel(arquivo, header=0)
            dados.columns = dados.columns.str.strip().str.lower()
            
            # Identificar colunas dinamicamente
            coluna_nome = encontrar_coluna(dados, ['nome', 'name'])
            coluna_telefone = encontrar_coluna(dados, ['telefone', 'phone', 'celular', 'número'])
            
            if coluna_nome is None or coluna_telefone is None:
                messagebox.showerror("Erro", "O arquivo deve ter colunas 'nome' e 'telefone'.")
                return

            dados.rename(columns={coluna_nome: 'nome', coluna_telefone: 'telefone'}, inplace=True)
            dados = dados.dropna(subset=["telefone", "nome"])

            # Recuperar progresso salvo
            if os.path.exists("progresso.txt"):
                with open("progresso.txt", "r") as file:
                    progresso = int(file.read().strip())
            else:
                progresso = 0

            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo: {e}")

# Função para encontrar colunas dinamicamente
def encontrar_coluna(df, palavras_chave):
    for coluna in df.columns:
        if any(palavra.lower() in coluna.lower() for palavra in palavras_chave):
            return coluna
    return None

# Função para enviar mensagens
def enviar_mensagens():
    global progresso, running, imagem_path
    running = True
    mensagem_personalizada = txt_mensagem.get("1.0", tk.END).strip()
    quantidade_mensagens = int(entry_quantidade.get())  # Obter a quantidade de mensagens

    for index, row in dados.iloc[progresso:].iterrows():
        if not running or (index - progresso) >= quantidade_mensagens:  # Verificar limite de mensagens
            break

        numero = str(row["telefone"]).strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        if not numero.startswith("+55"):
            numero = f"+55{numero}"

        primeiro_nome = str(row["nome"]).strip().split()[0]

        mensagem = mensagem_personalizada.replace("{nome}", primeiro_nome)

        try:
            print(f"Enviando mensagem para {numero}...")

            if imagem_path and os.path.exists(imagem_path):
                kit.sendwhats_image(numero, imagem_path, mensagem, wait_time=20)
            else:
                kit.sendwhatmsg_instantly(numero, mensagem, wait_time=10)

            enviar_mensagem_automaticamente()

            progresso = index + 1
            with open("progresso.txt", "w") as file:
                file.write(str(progresso))

        except Exception as e:
            print(f"Erro ao enviar para {numero}: {e}")
            continue

    running = False
    messagebox.showinfo("Finalizado", "Mensagens enviadas.")

# Função para automatizar o envio e fechar guia
def enviar_mensagem_automaticamente():
    time.sleep(8) 
    pyautogui.press('enter')
    time.sleep(6)
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(2)

# Função para carregar imagem
def carregar_imagem():
    global imagem_path
    imagem_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
    if imagem_path:
        lbl_imagem.config(text=f"Imagem: {os.path.basename(imagem_path)}")

# Função para excluir imagem
def excluir_imagem():
    global imagem_path
    imagem_path = None
    lbl_imagem.config(text="Nenhuma imagem carregada.")

# Controles do sistema
def iniciar():
    if dados is None:
        messagebox.showwarning("Aviso", "Carregue um arquivo primeiro.")
        return
    Thread(target=enviar_mensagens).start()

def finalizar():
    global running
    running = False

# Interface gráfica
root = tk.Tk()
root.title("Envio Automático")
root.geometry("600x500")

frame = tk.Frame(root)
frame.pack(pady=20)

btn_carregar = tk.Button(frame, text="Carregar Excel", command=carregar_arquivo)
btn_carregar.pack(side=tk.LEFT, padx=10)

btn_carregar_imagem = tk.Button(frame, text="Carregar Imagem", command=carregar_imagem)
btn_carregar_imagem.pack(side=tk.LEFT, padx=10)

btn_excluir_imagem = tk.Button(frame, text="Excluir Imagem", command=excluir_imagem)
btn_excluir_imagem.pack(side=tk.LEFT, padx=10)

btn_iniciar = tk.Button(frame, text="Iniciar", command=iniciar)
btn_iniciar.pack(side=tk.LEFT, padx=10)

btn_finalizar = tk.Button(frame, text="Finalizar", command=finalizar)
btn_finalizar.pack(side=tk.LEFT, padx=10)

lbl_mensagem = tk.Label(root, text="Mensagem (use {nome} para personalizar):")
lbl_mensagem.pack(pady=10)

txt_mensagem = scrolledtext.ScrolledText(root, width=70, height=5)
txt_mensagem.pack(pady=10)

lbl_quantidade = tk.Label(root, text="Quantidade de mensagens a enviar:")
lbl_quantidade.pack(pady=10)

entry_quantidade = tk.Entry(root)
entry_quantidade.pack(pady=10)
entry_quantidade.insert(0, "10")  # Valor padrão

lbl_imagem = tk.Label(root, text="Nenhuma imagem carregada.")
lbl_imagem.pack(pady=10)

root.mainloop()