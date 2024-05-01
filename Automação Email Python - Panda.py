import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import time

Empresas = pd.read_excel('./empresaAE.xlsx')

for index, empresa in Empresas.iterrows():
    msg = MIMEMultipart()
    msg['From'] = 'anam7615@gmail.com'
    msg['To'] = empresa['Email']
    msg['Subject'] = 'Possibilidade de Estágio'
    message = f"""\
Boa Tarde/Dia, {empresa['Empresas']}, tudo bem?

\nMe chamo Gabriel Gadelha, sou estudante de Análise e Desenvolvimento de Sistemas, especializando em Analista de Teste mas com aprendizado em Lógica de Programação.

\nTerminei curso de lógica de programação em C e aprendendo Python, e já com curso para o backend que estou fazendo juntamente com Analista de Teste (QA).

\nEstou enviando este email via automatização com Python utilizando Pandas, gostaria de saber se poderia ter a oportunidade de conversar sobre a possibilidade de estágio.

\nEstou enviando o GitHub do código para análise caso haja curiosidade: https://github.com/gabztoo/Automa-o-de-Email/blob/main/Automação%20Email%20Python%20-%20Panda.py

\nObrigado pela atenção!
"""
    msg.attach(MIMEText(message, 'plain'))

    # Anexando o arquivo PDF
    filename = 'Gabriel Gadelha - TI.pdf'  # Substitua 'arquivo.pdf' pelo nome do seu arquivo PDF
    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {filename}',
    )
    msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', port=587)
        server.starttls()
        server.login('anam7615@gmail.com', 'twoizwnwxfqzugtx')
        server.send_message(msg)
        server.quit()
        print(f"Email sent to {empresa['Email']}")
    except Exception as e:
        print(f"Error sending email to {empresa['Email']}: {str(e)}")
    
    # Adiciona 2 segundos de Delay pra não cair no anti-spam
    time.sleep(1)  # 2 segundos de Delay
