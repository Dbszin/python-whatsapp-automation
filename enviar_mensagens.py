import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
from whatsapp_sender import WhatsAppSender
import os

class WhatsAppGUI:
    """
    Classe responsável pela interface gráfica e interação com o usuário.
    """
    def __init__(self, root):
        self.root = root
        self.sender = WhatsAppSender()
        self.running = False
        self._setup_ui()

    def _setup_ui(self):
        self.root.title("Envio Automático de WhatsApp")
        self.root.geometry("650x600")

        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        btn_carregar = tk.Button(frame, text="Carregar Excel", command=self.carregar_arquivo)
        btn_carregar.pack(side=tk.LEFT, padx=10)

        btn_carregar_imagem = tk.Button(frame, text="Carregar Imagem", command=self.carregar_imagem)
        btn_carregar_imagem.pack(side=tk.LEFT, padx=10)

        btn_excluir_imagem = tk.Button(frame, text="Excluir Imagem", command=self.excluir_imagem)
        btn_excluir_imagem.pack(side=tk.LEFT, padx=10)

        btn_iniciar = tk.Button(frame, text="Iniciar", command=self.iniciar_envio)
        btn_iniciar.pack(side=tk.LEFT, padx=10)

        btn_finalizar = tk.Button(frame, text="Cancelar", command=self.cancelar_envio)
        btn_finalizar.pack(side=tk.LEFT, padx=10)

        lbl_mensagem = tk.Label(self.root, text="Mensagem (use {nome}, {telefone}, etc):")
        lbl_mensagem.pack(pady=10)

        self.txt_mensagem = scrolledtext.ScrolledText(self.root, width=70, height=5)
        self.txt_mensagem.pack(pady=10)

        lbl_quantidade = tk.Label(self.root, text="Quantidade de mensagens a enviar:")
        lbl_quantidade.pack(pady=5)

        self.entry_quantidade = tk.Entry(self.root)
        self.entry_quantidade.pack(pady=5)
        self.entry_quantidade.insert(0, "10")

        lbl_linha_inicial = tk.Label(self.root, text="Número da linha inicial para começar:")
        lbl_linha_inicial.pack(pady=5)

        self.entry_linha_inicial = tk.Entry(self.root)
        self.entry_linha_inicial.pack(pady=5)
        self.entry_linha_inicial.insert(0, "1")

        self.lbl_imagem = tk.Label(self.root, text="Nenhuma imagem carregada.")
        self.lbl_imagem.pack(pady=10)

        # Barra de progresso
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=10)

        # Área de status/log visual
        self.txt_status = scrolledtext.ScrolledText(self.root, width=70, height=8, state='disabled')
        self.txt_status.pack(pady=10)

    def carregar_arquivo(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not arquivo:
            return
        if self.sender.carregar_arquivo(arquivo):
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
            self.log_status(f"Arquivo carregado: {arquivo}")
        else:
            messagebox.showerror("Erro", "Erro ao abrir o arquivo. Verifique o log para detalhes.")
            self.log_status("Erro ao carregar arquivo.")

    def carregar_imagem(self):
        imagem_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
        self.sender.set_imagem(imagem_path)
        self.lbl_imagem.config(text=f"Imagem: {os.path.basename(imagem_path)}" if imagem_path else "Nenhuma imagem carregada.")
        self.log_status(f"Imagem carregada: {imagem_path}" if imagem_path else "Imagem removida.")

    def excluir_imagem(self):
        self.sender.set_imagem(None)
        self.lbl_imagem.config(text="Nenhuma imagem carregada.")
        self.log_status("Imagem removida.")

    def iniciar_envio(self):
        if self.running:
            messagebox.showinfo("Aviso", "O envio já está em andamento.")
            return
        self.running = True
        self.progress['value'] = 0
        self.txt_status.config(state='normal')
        self.txt_status.delete('1.0', tk.END)
        self.txt_status.config(state='disabled')
        thread = threading.Thread(target=self.enviar_mensagens, daemon=True)
        thread.start()

    def cancelar_envio(self):
        if self.running:
            self.running = False
            self.log_status("Envio cancelado pelo usuário.")
        else:
            messagebox.showinfo("Aviso", "Nenhum envio em andamento.")

    def enviar_mensagens(self):
        mensagem_base = self.txt_mensagem.get("1.0", tk.END).strip()
        try:
            linha_inicial = int(self.entry_linha_inicial.get()) - 1
            quantidade_mensagens = int(self.entry_quantidade.get())
        except ValueError:
            messagebox.showerror("Erro", "Insira números válidos para linha inicial e quantidade.")
            self.running = False
            return
        total = quantidade_mensagens
        self.progress['maximum'] = total
        enviados = 0
        def update_callback(ultimo_contato):
            nonlocal enviados
            enviados += 1
            self.progress['value'] = enviados
            self.log_status(f"Enviado para {ultimo_contato['nome']} ({ultimo_contato['telefone']}) - Linha {ultimo_contato['linha']}")
            self.root.update_idletasks()
        def stop_flag():
            return not self.running
        try:
            self.sender.enviar_mensagens(mensagem_base, linha_inicial, quantidade_mensagens, update_callback, stop_flag)
            if self.running:
                messagebox.showinfo("Finalizado", "Mensagens enviadas.")
                self.log_status("Envio finalizado com sucesso.")
            else:
                self.log_status("Envio interrompido.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar mensagens: {e}")
            self.log_status(f"Erro: {e}")
        self.running = False

    def log_status(self, msg):
        self.txt_status.config(state='normal')
        self.txt_status.insert(tk.END, msg + '\n')
        self.txt_status.see(tk.END)
        self.txt_status.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppGUI(root)
    root.mainloop()
