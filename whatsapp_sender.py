import pandas as pd
import pywhatkit as kit
import os
import time
import pyautogui
import logging

class WhatsAppSender:
    """
    Classe responsável por carregar contatos, validar dados e enviar mensagens via WhatsApp.
    """
    def __init__(self, imagem_path=None):
        self.dados = None
        self.imagem_path = imagem_path
        self.ultimo_contato = None
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 1
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def carregar_arquivo(self, arquivo):
        """
        Carrega e valida o arquivo Excel de contatos.
        Retorna True em caso de sucesso, False caso contrário.
        """
        try:
            self.dados = pd.read_excel(arquivo, dtype=str)
            self.dados.columns = self.dados.columns.str.strip().str.lower()
            coluna_nome = self._encontrar_coluna(['nome', 'name'])
            coluna_telefone = self._encontrar_coluna(['telefone', 'phone', 'celular', 'número'])
            if coluna_nome is None or coluna_telefone is None:
                raise ValueError("O arquivo deve ter colunas 'nome' e 'telefone'.")
            self.dados.rename(columns={coluna_nome: 'nome', coluna_telefone: 'telefone'}, inplace=True)
            self.dados = self.dados.dropna(subset=["telefone", "nome"])
            self.dados["telefone"] = (
                self.dados["telefone"].astype(str)
                .str.replace(r"[^\d]", "", regex=True)
                .str.lstrip("0")
            )
            self.dados["telefone"] = self.dados["telefone"].apply(lambda x: f"+55{x}" if not x.startswith("+55") else x)
            logging.info(f"Arquivo {arquivo} carregado com sucesso.")
            return True
        except Exception as e:
            logging.error(f"Erro ao carregar arquivo: {e}")
            return False

    def _encontrar_coluna(self, palavras_chave):
        """
        Encontra a primeira coluna que contenha uma das palavras-chave.
        """
        for coluna in self.dados.columns:
            if any(palavra.lower() in coluna.lower() for palavra in palavras_chave):
                return coluna
        return None

    def set_imagem(self, imagem_path):
        """
        Define o caminho da imagem a ser enviada. Se None, remove a imagem.
        """
        if imagem_path and os.path.exists(imagem_path):
            self.imagem_path = imagem_path
            logging.info(f"Imagem definida: {imagem_path}")
        else:
            self.imagem_path = None
            logging.info("Imagem removida ou não encontrada.")

    def enviar_mensagens(self, mensagem_base, linha_inicial, quantidade_mensagens, update_callback=None, stop_flag=None):
        """
        Envia mensagens para os contatos carregados.
        update_callback: função chamada após cada envio (para atualizar UI, logs, etc)
        stop_flag: função que retorna True se o envio deve ser interrompido
        Suporta múltiplos campos de personalização no formato {coluna}.
        """
        if self.dados is None:
            raise ValueError("Nenhum dado carregado.")
        df_mensagens = self.dados.iloc[linha_inicial:linha_inicial+quantidade_mensagens]
        for index, row in df_mensagens.iterrows():
            if stop_flag and stop_flag():
                break
            mensagem = mensagem_base
            # Substitui todos os campos {coluna} pelo valor correspondente
            for col in self.dados.columns:
                valor = str(row[col]) if pd.notnull(row[col]) else ""
                mensagem = mensagem.replace(f"{{{col}}}", valor)
            try:
                numero = row["telefone"]
                logging.info(f"Enviando mensagem para {numero}...")
                if self.imagem_path:
                    kit.sendwhats_image(numero, self.imagem_path, mensagem, wait_time=20)
                else:
                    kit.sendwhatmsg_instantly(numero, mensagem, wait_time=20)
                # Aguarda o tempo mínimo necessário para garantir o envio
                time.sleep(8)
                pyautogui.press('enter')
                time.sleep(5)
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(1)
                self.ultimo_contato = {"nome": row["nome"], "telefone": numero, "linha": index + 2}
                if update_callback:
                    update_callback(self.ultimo_contato)
            except Exception as e:
                logging.error(f"Erro ao enviar para {numero}: {e}")
                if update_callback:
                    update_callback({"nome": row["nome"], "telefone": numero, "linha": index + 2, "erro": str(e)})
                continue

    def get_ultimo_contato(self):
        """
        Retorna o último contato para o qual a mensagem foi enviada.
        """
        return self.ultimo_contato 