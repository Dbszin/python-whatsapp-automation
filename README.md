# Python WhatsApp Automation

Automatize o envio de mensagens personalizadas pelo WhatsApp usando Python, Excel e uma interface gráfica simples.

## Funcionalidades
- Envio automático de mensagens para contatos de uma planilha Excel
- Personalização de mensagens com qualquer campo do Excel (ex: {nome}, {telefone}, {status})
- Suporte ao envio de imagens
- Barra de progresso e log visual
- Cancelamento do envio a qualquer momento
- Interface gráfica intuitiva

## Instalação

1. **Clone o repositório:**
   ```bash
   git clone <url-do-repositorio>
   cd python-whatsapp-automation-1
   ```
2. **Crie um ambiente virtual (opcional, mas recomendado):**
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```
3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

## Como usar

1. Prepare um arquivo Excel (.xlsx) com pelo menos as colunas `nome` e `telefone` (pode ter outras colunas para personalização).
2. Execute o sistema:
   ```bash
   python enviar_mensagens.py
   ```
3. Na interface:
   - Clique em "Carregar Excel" e selecione sua planilha.
   - (Opcional) Carregue uma imagem para enviar junto.
   - Escreva sua mensagem, usando chaves para personalizar (ex: `Olá {nome}, seu status é {status}`)
   - Defina a linha inicial e a quantidade de mensagens.
   - Clique em "Iniciar" para começar o envio.
   - Acompanhe o progresso e o log visual.
   - Cancele a qualquer momento se necessário.

## Observações
- O WhatsApp Web será aberto automaticamente para cada envio.
- O tempo de espera entre mensagens é ajustado para evitar bloqueios.
- Não compartilhe dados sensíveis em sua planilha.

## Dependências principais
- pandas
- pywhatkit
- pyautogui
- tkinter

---

**Projeto refatorado para máxima produtividade, robustez e facilidade de uso!**

