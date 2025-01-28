import pandas as pd
import pywhatkit as kit
import time
import keyboard
import pyautogui

# Configurações de segurança do PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

# 1. Ler o arquivo Excel
try:
    dados = pd.read_excel("arquivo excell", header=1)
    dados = dados.dropna(axis=1, how='all')
    dados.columns = dados.columns.str.strip()
    
    print("Colunas disponíveis:", list(dados.columns))
    print("\nPressione 'q' a qualquer momento para parar o programa de forma segura.")
    
except Exception as e:
    print("Erro ao abrir o arquivo Excel. Verifique o nome e o formato.")
    print(e)
    exit()

max_mensagens = 10
contador = 0

# 2. Carregar o progresso
try:
    with open("progresso.txt", "r") as file:
        progresso = int(file.read())
except FileNotFoundError:
    progresso = 0

def deve_parar():
    return keyboard.is_pressed('q')

def enviar_mensagem_automaticamente():
    # Espera a página do WhatsApp Web carregar
    time.sleep(10)
    # Pressiona ENTER para enviar
    pyautogui.press('enter')
    # Espera a mensagem ser enviada
    time.sleep(7)
    # Fecha a guia atual (Ctrl + W)
    pyautogui.hotkey('ctrl', 'w')
    # Espera a guia fechar
    time.sleep(2)

try:
    for index, row in dados.iloc[progresso:].iterrows():
        if contador >= max_mensagens:
            print("Limite de mensagens atingido. Finalizando...")
            break

        if deve_parar():
            print("\nPrograma interrompido pelo usuário!")
            break

        numero = str(row["telefone"]).strip()
        numero = numero.replace("(", "").replace(")", "").replace("-", "").replace(" ", "")
        
        if not numero.startswith("+55"):
            numero = f"+55{numero}"

        nome_completo = str(row["nome"]).strip()
        primeiro_nome = nome_completo.split()[0] if nome_completo else ""

        mensagem = f" Olá {primeiro_nome}, Mensagem a ser enviada"

        try:
            print(f"\nEnviando mensagem para {numero}...")
            print(f"Nome: {primeiro_nome}")
            
            # Envia a mensagem
            kit.sendwhatmsg_instantly(numero, mensagem, wait_time=15)
            
            # Automatiza o envio e fecha a guia
            enviar_mensagem_automaticamente()
            
            contador += 1

            # Salva o progresso
            with open("progresso.txt", "w") as file:
                file.write(str(index+1))

        except Exception as e:
            print(f"Erro ao enviar mensagem para {numero}: {e}")
            continue

except KeyboardInterrupt:
    print("\nPrograma interrompido pelo usuário!")

finally:
    print("\nProgresso salvo. Total de mensagens enviadas:", contador)
    print("Posição atual salva no arquivo progresso.txt")
    print("Programa finalizado.")