

## **Descrição do Projeto**
Este script automatiza o envio de mensagens personalizadas para contatos via WhatsApp Web. Ele utiliza um arquivo Excel contendo informações de contato (números de telefone e nomes) e realiza o envio de mensagens em massa, garantindo a continuidade do processo mesmo após interrupções, utilizando um sistema de progresso salvo.

---

## **Funcionalidades**
- Lê um arquivo Excel com os dados de contatos.
- Envia mensagens no **WhatsApp Web** utilizando a biblioteca `pywhatkit`.
- Salva o progresso em um arquivo `progresso.txt` para retomar o envio em execuções futuras.
- Possui limite configurável para o número de mensagens enviadas por execução.
- Permite interromper o processo pressionando a tecla `q`.

---

## **Requisitos do Sistema**
### **Linguagem e Bibliotecas**
- **Python 3.8+**
- Bibliotecas:
  - `pandas`
  - `pywhatkit`
  - `pyautogui`
  - `keyboard`

### **Configurações do Sistema**
- É necessário que o **WhatsApp Web** esteja configurado no navegador.
- O arquivo Excel deve conter pelo menos duas colunas: `nome` e `telefone`.

---

## **Instalação**
1. **Clone o repositório ou copie o script.**
2. **Instale as dependências do Python:**
   ```bash
   pip install pandas pywhatkit pyautogui keyboard
   ```

3. **Prepare o arquivo Excel:**
   - Certifique-se de que ele contém as colunas `nome` e `telefone`.
   - Salve o arquivo com o nome `arquivo excell` no mesmo diretório do script.

---

## **Como Usar**
1. **Certifique-se de que o arquivo Excel está configurado corretamente.**
2. **Execute o script:**
   ```bash
   python script.py
   ```

3. **Durante a execução:**
   - Pressione `q` para interromper o programa de forma segura.
   - O script exibirá as colunas detectadas no Excel e iniciará o envio das mensagens.

4. **Progresso Salvo:**
   - O progresso é salvo automaticamente no arquivo `progresso.txt`. 
   - Caso o script seja interrompido, ele retomará do último contato enviado.

---

## **Configurações Importantes**
- **Limite de Mensagens por Execução:**
  O limite está configurado para 10 mensagens por execução. Você pode alterar a variável `max_mensagens` para ajustar esse valor:
  ```python
  max_mensagens = 10
  ```

- **Formato do Número de Telefone:**
  - O script adiciona o código do Brasil `+55` automaticamente se ele não estiver presente.

- **Mensagem Padrão:**
  Modifique a mensagem no seguinte trecho:
  ```python
  mensagem = f"Olá {primeiro_nome}, Mensagem a ser enviada"
  ```

---

## **Dicas**
1. **Teste antes:**
   - Execute o script com um pequeno número de contatos para verificar se está funcionando corretamente.
   
2. **Evite Bloqueios:**
   - O script utiliza um tempo de espera para evitar bloqueios no WhatsApp. Não reduza os tempos de espera (`time.sleep` e `wait_time`) para evitar ser sinalizado.

3. **Fechar o WhatsApp Web:**
   - O script fecha automaticamente a aba do WhatsApp Web após o envio de cada mensagem.

---

## **Erros Comuns**
- **Erro ao abrir o arquivo Excel:**
  - Verifique o nome do arquivo e se ele está no mesmo diretório do script.
  - Certifique-se de que o arquivo está no formato `.xlsx`.

- **WhatsApp Web não abre:**
  - Certifique-se de que o navegador está configurado corretamente para abrir o WhatsApp Web.

---

## **Segurança**
- O script utiliza `pyautogui`, que controla o mouse e o teclado. Não mova o cursor ou altere as janelas durante a execução.

---

## **Contribuições**
Contribuições são bem-vindas! Para reportar problemas ou sugerir melhorias, abra uma issue ou envie um pull request.

---

## **Licença**
Este projeto está licenciado sob a licença MIT. Sinta-se livre para utilizá-lo e modificá-lo conforme necessário.
