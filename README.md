# **Envio Automático de Mensagens no WhatsApp**

## **Descrição do Projeto**

Este projeto automatiza o envio de mensagens personalizadas para contatos via **WhatsApp Web**. Ele utiliza um arquivo **Excel** contendo informações de contato (números de telefone e nomes) e permite configurar o envio de mensagens de forma intuitiva por meio de uma interface gráfica.

Além disso, o sistema permite o envio de imagens e armazena o progresso para evitar mensagens duplicadas, tornando a automação mais eficiente e confiável.

---

## **Funcionalidades**

- Lê um arquivo **Excel** com os dados dos contatos.
- Envia mensagens no **WhatsApp Web** utilizando a biblioteca `pywhatkit`.
- Possui interface gráfica para facilitar o carregamento de arquivos e a configuração do envio.
- Suporte para envio de imagens junto com a mensagem.
- Registra o progresso para evitar mensagens duplicadas.
- Permite interromper o processo a qualquer momento.
- Ajuste dinâmico do tempo de envio para evitar bloqueios pelo WhatsApp.
- Exibição de status detalhado durante o processo de envio.

---

## **Requisitos do Sistema**

### **Linguagem e Bibliotecas**

- **Python 3.8+**
- Bibliotecas necessárias:
  ```bash
  pip install pandas pywhatkit pyautogui tk openpyxl
  ```

### **Configurações do Sistema**

- O **WhatsApp Web** deve estar logado no navegador padrão.
- O arquivo Excel deve conter pelo menos duas colunas: `Nome` e `Telefone`.

---

## **Instalação e Execução**

### **Passo 1: Baixar e Configurar o Projeto**

1. **Clone o repositório ou copie os arquivos do projeto:**
   ```bash
   git clone https://github.com/seu-repositorio.git
   cd seu-repositorio
   ```
2. **Instale as dependências conforme listado acima.**

### **Passo 2: Como Usar**

1. **Execute o script principal do projeto:**
   ```bash
   python main.py
   ```
2. **Carregue a planilha Excel:**
   - Clique no botão **"Carregar Excel"** para selecionar o arquivo.
3. **(Opcional) Carregue uma imagem:**
   - Clique em **"Carregar Imagem"** para anexar uma imagem ao envio.
4. **Digite a mensagem desejada:**
   - Utilize `{nome}` para personalizar a mensagem com o nome do destinatário.
5. **Defina a quantidade de mensagens a serem enviadas.**
6. **Clique em "Iniciar" para começar o envio.**
7. **Para interromper o envio, clique em "Finalizar".**

---

## **Formato da Planilha Excel**

A planilha deve conter as seguintes colunas:

- `Nome` - Nome do contato (utilizado para personalização da mensagem).
- `Telefone` - Número do contato no formato internacional (**+55** para Brasil).

**Exemplo:**

| Nome  | Telefone       |
| ----- | -------------- |
| João  | +0000000000000 |
| Maria | +0000000000000 |

---

## **Avisos e Recomendações**

- O envio de mensagens em massa pode estar sujeito a **bloqueios pelo WhatsApp**. Evite enviar muitas mensagens em curtos espaços de tempo.
- O script utiliza **pyautogui**, que controla o mouse e o teclado. Evite mover o cursor ou alterar as janelas durante a execução.
- Antes de enviar mensagens para uma grande lista de contatos, **teste o funcionamento com um pequeno grupo.**
- O tempo de envio entre mensagens pode ser ajustado no código para evitar detecção como spam.

---

## **Possíveis Erros e Soluções**

| Erro                  | Causa                               | Solução                                     |
| --------------------- | ----------------------------------- | ------------------------------------------- |
| `FileNotFoundError`   | Arquivo Excel não encontrado        | Verifique o nome e local do arquivo         |
| `ModuleNotFoundError` | Biblioteca não instalada            | Execute `pip install -r requirements.txt`   |
| WhatsApp Web não abre | Não está logado no navegador padrão | Faça login no WhatsApp Web antes de iniciar |

---

## **Contribuições**

Contribuições são bem-vindas! Se encontrar problemas ou tiver sugestões, abra uma **issue** ou envie um **pull request**.

---

## **Licença**

Este projeto está licenciado sob a **MIT License**. Você é livre para utilizá-lo e modificá-lo conforme necessidade.

