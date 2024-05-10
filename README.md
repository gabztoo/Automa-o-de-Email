![imagem_2024-05-10_195535708](https://github.com/gabztoo/Automa-o-de-Email/assets/162667498/5f3e22e4-d769-4a93-9c24-0dfcc02f27b3)
# Envio de E-mails em Massa utilizando Tkinter e Pandas

Este é um script Python que utiliza a biblioteca Tkinter para criar uma interface gráfica simples e permitir o envio de e-mails em massa. O script carrega dados de um arquivo Excel, preenche os campos do e-mail com informações personalizadas e permite anexar arquivos antes de enviar os e-mails.

## Funcionalidades:

Seleção de Arquivo Excel: Permite selecionar o arquivo Excel contendo os destinatários e outras informações necessárias.
Autenticação do Remetente: Solicita o e-mail e a senha do remetente para autenticação no servidor SMTP. Recomenda-se utilizar uma senha de aplicativo do Google para garantir maior segurança.
Personalização de E-mails: Permite personalizar o assunto e o corpo do e-mail com base nos dados do arquivo Excel.
Anexo de Arquivos: Oferece a opção de anexar arquivos aos e-mails a serem enviados.
Feedback de Envio: Fornece feedback em tempo real sobre para quem cada e-mail foi enviado.
Pré-requisitos:

## Antes de executar o script, certifique-se de ter instalado todas as dependências necessárias. Você pode instalar as dependências listadas no arquivo requirements.txt usando o comando:


`pip install -r requirements.txt`


## Utilização:

Execute o script envio_de_emails.py.
Preencha os campos necessários na interface gráfica:
Selecione o arquivo Excel contendo os destinatários.
Insira o e-mail do remetente e utilize a senha de aplicativo do Google gerada para este fim.
Escreva o assunto e o corpo do e-mail.
Se desejar, anexe um arquivo.
Clique no botão "Enviar E-mails" para iniciar o envio.
Aguarde até que todos os e-mails sejam enviados. O progresso será exibido na interface.
Observações:

* Na planilha excel, tem de haver as colunas **"Email"** e **"Nome"** (segue exemplo)
![imagem_2024-05-10_195617670](https://github.com/gabztoo/Automa-o-de-Email/assets/162667498/078b51b6-e4db-4f52-9430-bc1c51c13a6f)
.

* Certifique-se de que a conta de e-mail do remetente tenha permissão para enviar e-mails via SMTP.

* Utilize uma senha de aplicativo do Google para autenticação, garantindo maior segurança e evitando problemas de acesso.

* É recomendável verificar as configurações de segurança da sua conta de e-mail para evitar bloqueios por parte do provedor de e-mail.

* Seguindo estes passos, você poderá enviar e-mails em massa de forma eficiente e segura utilizando o script Python com Tkinter e Pandas.


![imagem_2024-05-10_163006438](https://github.com/gabztoo/Automa-o-de-Email/assets/162667498/683be1b8-eea4-405a-b10f-2611d73ebeae)
