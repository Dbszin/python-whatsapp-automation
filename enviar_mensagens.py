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
running = False
imagem_path = None
ultimo_contato = None  # Armazena os dados do último contato enviado

# Função para carregar o arquivo Excel
def carregar_arquivo():
    global dados
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not arquivo:
        return

    try:
        # Ler o arquivo Excel e garantir que os telefones sejam tratados como string
        dados = pd.read_excel(arquivo, dtype=str)  
        dados.columns = dados.columns.str.strip().str.lower()

        coluna_nome = encontrar_coluna(dados, ['nome', 'name'])
        coluna_telefone = encontrar_coluna(dados, ['telefone', 'phone', 'celular', 'número'])

        if coluna_nome is None or coluna_telefone is None:
            messagebox.showerror("Erro", "O arquivo deve ter colunas 'nome' e 'telefone'.")
            return

        # Padronizar colunas
        dados.rename(columns={coluna_nome: 'nome', coluna_telefone: 'telefone'}, inplace=True)
        dados = dados.dropna(subset=["telefone", "nome"])

        # Limpar números de telefone (remover espaços, caracteres não numéricos e formatar corretamente)
        dados["telefone"] = (
            dados["telefone"]
            .astype(str)
            .str.replace(r"[^\d]", "", regex=True)  # Remove tudo que não for número
            .str.lstrip("0")  # Remove zeros no início
        )

        # Adicionar +55 se o número não começar com ele
        dados["telefone"] = dados["telefone"].apply(lambda x: f"+55{x}" if not x.startswith("+55") else x)

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
    global running, ultimo_contato
    running = True
    mensagem_base = txt_mensagem.get("1.0", tk.END).strip()

    try:
        linha_inicial = int(entry_linha_inicial.get()) - 1
        quantidade_mensagens = int(entry_quantidade.get())
    except ValueError:
        messagebox.showerror("Erro", "Insira números válidos para linha inicial e quantidade.")
        return

    df_mensagens = dados.iloc[linha_inicial : linha_inicial + quantidade_mensagens]

    for index, row in df_mensagens.iterrows():
        if not running:
            break

        primeiro_nome = str(row["nome"]).strip().split()[0]
        numero = row["telefone"]
        mensagem = mensagem_base.replace("{nome}", primeiro_nome)

        try:
            print(f"Enviando mensagem para {numero}...")

            if imagem_path and os.path.exists(imagem_path):
                kit.sendwhats_image(numero, imagem_path, mensagem, wait_time=20)
            else:
                kit.sendwhatmsg_instantly(numero, mensagem, wait_time=15)

            # Salvar o último contato enviado
            ultimo_contato = {
                "nome": row["nome"],
                "telefone": numero,
                "linha": index + 2  # +2 porque pandas começa do 0 e no Excel começa de 1
            }

            # Enviar mensagem e fechar guia de forma não bloqueante
            Thread(target=enviar_mensagem_automaticamente).start()

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
    lbl_imagem.config(text=f"Imagem: {os.path.basename(imagem_path)}" if imagem_path else "Nenhuma imagem carregada.")

# Função para excluir imagem
def excluir_imagem():
    global imagem_path
    imagem_path = None
    lbl_imagem.config(text="Nenhuma imagem carregada.")

# Função para exibir o último contato ao finalizar
def finalizar():
    global running
    running = False

    if ultimo_contato:
        messagebox.showinfo(
            "Última Mensagem Enviada",
            f"📌 Último contato enviado:\n"
            f"👤 Nome: {ultimo_contato['nome']}\n"
            f"📞 Telefone: {ultimo_contato['telefone']}\n"
            f"📊 Linha no Excel: {ultimo_contato['linha']}"
        )
    else:
        messagebox.showinfo("Finalizado", "Nenhuma mensagem foi enviada ainda.")

# Interface gráfica
root = tk.Tk()
root.title("Envio Automático")
root.geometry("600x550")

frame = tk.Frame(root)
frame.pack(pady=20)

btn_carregar = tk.Button(frame, text="Carregar Excel", command=carregar_arquivo)
btn_carregar.pack(side=tk.LEFT, padx=10)

btn_carregar_imagem = tk.Button(frame, text="Carregar Imagem", command=carregar_imagem)
btn_carregar_imagem.pack(side=tk.LEFT, padx=10)

btn_excluir_imagem = tk.Button(frame, text="Excluir Imagem", command=excluir_imagem)
btn_excluir_imagem.pack(side=tk.LEFT, padx=10)

btn_iniciar = tk.Button(frame, text="Iniciar", command=lambda: Thread(target=enviar_mensagens, daemon=True).start())
btn_iniciar.pack(side=tk.LEFT, padx=10)

btn_finalizar = tk.Button(frame, text="Finalizar", command=finalizar)
btn_finalizar.pack(side=tk.LEFT, padx=10)

lbl_mensagem = tk.Label(root, text="Mensagem (use {nome} para personalizar):")
lbl_mensagem.pack(pady=10)

txt_mensagem = scrolledtext.ScrolledText(root, width=70, height=5)
txt_mensagem.pack(pady=10)

lbl_quantidade = tk.Label(root, text="Quantidade de mensagens a enviar:")
lbl_quantidade.pack(pady=5)

entry_quantidade = tk.Entry(root)
entry_quantidade.pack(pady=5)
entry_quantidade.insert(0, "10")

lbl_linha_inicial = tk.Label(root, text="Número da linha inicial para começar:")
lbl_linha_inicial.pack(pady=5)

entry_linha_inicial = tk.Entry(root)
entry_linha_inicial.pack(pady=5)
entry_linha_inicial.insert(0, "1")

lbl_imagem = tk.Label(root, text="Nenhuma imagem carregada.")
lbl_imagem.pack(pady=10)

root.mainloop()
