import pandas as pd
import pywhatkit as kit
import time
import keyboard
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

# Configurações de segurança do PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

# Variáveis globais
dados = None
contador = 0
progresso = 0
max_mensagens = 500
paused = False
running = False

# Função para carregar o arquivo Excel
def carregar_arquivo():
    global dados, progresso
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if arquivo:
        try:
            dados = pd.read_excel(arquivo, header=0)  # Lendo a primeira linha como cabeçalho
            
            # Remover espaços extras e padronizar nomes das colunas
            dados.columns = dados.columns.str.strip().str.lower()
            
            if "telefone" not in dados.columns or "nome" not in dados.columns:
                messagebox.showerror("Erro", "O arquivo Excel deve conter as colunas 'telefone' e 'nome'.")
                return

            dados = dados.dropna(subset=["telefone", "nome"])  # Remover linhas sem telefone ou nome
            progresso = 0
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo Excel: {e}")

# Função para enviar mensagens
def enviar_mensagens():
    global contador, progresso, running
    running = True
    for index, row in dados.iloc[progresso:].iterrows():
        if not running:  # Verifica se o processo deve ser interrompido
            break

        numero = str(row["telefone"]).strip()
        numero = numero.replace("(", "").replace(")", "").replace("-", "").replace(" ", "")
        
        if not numero.startswith("+55"):
            numero = f"+55{numero}"

        nome_completo = str(row["nome"]).strip()
        primeiro_nome = nome_completo.split()[0] if nome_completo else ""

        mensagem = f" Olá {primeiro_nome}, temos uma promoção especial para você:\n💥 Plano Dopamina Edos Fit 💥\n⚡ Por apenas R$89,90/mês! ⚡\n\n🎯 Tudo que você precisa para transformar seu estilo de vida:\n✅ Musculação, Funcional, Fitdance, Big ass\n✅ App de treino personalizado na palma da mão\n✅ Cadeira de massagem relaxante incluída\n\n🚨 Vagas limitadas!\n🔥 Exclusivo para novos alunos e upgrades de plano!\n\n⏳ Não perca tempo! A sua mudança começa agora! Aproveite!"

        try:
            print(f"\nEnviando mensagem para {numero}...")

            kit.sendwhatmsg_instantly(numero, mensagem, wait_time=15)
            enviar_mensagem_automaticamente()

            contador += 1

            with open("progresso.txt", "w") as file:
                file.write(str(index+1))

        except Exception as e:
            print(f"Erro ao enviar mensagem para {numero}: {e}")
            continue

    running = False
    messagebox.showinfo("Finalizado", "Processo finalizado.")

def finalizar():
    global running
    running = False
    print("Processo finalizado pelo usuário.")


# Função para automatizar o envio e fechar a guia
def enviar_mensagem_automaticamente():
    time.sleep(10)
    pyautogui.press('enter')
    time.sleep(7)
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(2)

# Funções para controlar o sistema
def iniciar():
    if dados is None:
        messagebox.showwarning("Aviso", "Por favor, carregue um arquivo Excel primeiro.")
        return
    Thread(target=enviar_mensagens).start()

def pausar():
    global paused
    paused = not paused

def finalizar():
    global running
    running = False

# Interface gráfica
root = tk.Tk()
root.title("Envio de Mensagens Automático")

frame = tk.Frame(root)
frame.pack(pady=20)

btn_carregar = tk.Button(frame, text="Carregar Arquivo Excel", command=carregar_arquivo)
btn_carregar.pack(side=tk.LEFT, padx=10)

btn_iniciar = tk.Button(frame, text="Iniciar", command=iniciar)
btn_iniciar.pack(side=tk.LEFT, padx=10)

btn_pausar = tk.Button(frame, text="Pausar", command=pausar)
btn_pausar.pack(side=tk.LEFT, padx=10)

btn_finalizar = tk.Button(frame, text="Finalizar", command=finalizar)
btn_finalizar.pack(side=tk.LEFT, padx=10)

root.mainloop()