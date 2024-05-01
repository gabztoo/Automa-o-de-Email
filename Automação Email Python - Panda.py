import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

Empresas = pd.read_excel('./empresas.xlsx')

for index, empresa in Empresas.iterrows():
    msg = MIMEMultipart()
    msg['From'] = 'remetente@email.com'
    msg['To'] = 'destinatario@gmail.com'  # Use o e-mail específico da empresa
    msg['Subject'] = 'assunto'

    message = f"""\
    Boa Tarde/Dia, Gabriel, tudo bem?
    Me chamo Gabriel Gadelha, sou estudante de Análise e Desenvolvimento de Sistemas, especializando em Analista de Teste mas com aprendizado em Lógica de Programação.
    \nTerminei curso de lógica de programação em C e aprendendo Python, e já com curso para o backend que estou fazendo juntamente com Analista de Teste (QA).
    \nEstou enviando este email via automatização com Python utilizando Panda, gostaria de saber se poderia ter a oportunidade de conversar sobre a possibilidade de estágio.
    \nObrigado pela atenção!
    """
    msg.attach(MIMEText(message, 'plain'))

    # Anexando o arquivo PDF
    filename = 'Gabriel Gadelha - TI.pdf'  # Altere para o nome do seu arquivo PDF
    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {filename}',
    )
    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', port=587)
    server.starttls()
    server.login('remetente@email.com', 'senhaapp')
    text = msg.as_string()
    server.sendmail(msg['From'], msg['To'], text)
    server.quit()
