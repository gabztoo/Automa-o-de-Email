import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import time

# Função para enviar e-mails
def enviar_emails():
    try:
        # Carregar dados do arquivo Excel
        excel = pd.read_excel(filepath.get())

        for index, excel_row in excel.iterrows():
            msg = MIMEMultipart()
            msg['From'] = remetente.get()
            msg['To'] = excel_row['email']
            msg['Subject'] = assunto_text.get()  # Usar o assunto fornecido pelo usuário

            # Adicionar saudação personalizada
            saudacao = f"Boa tarde, {excel_row['pessoa']}!\n\n"  # Supondo que você tenha uma coluna 'pessoa' no Excel
            message = saudacao + mensagem_text.get("1.0", tk.END)

            # Corpo do e-mail
            msg.attach(MIMEText(message, 'plain'))

            # Anexar o arquivo PDF
            filename = arquivo_filepath.get()
            if filename:
                with open(filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {filename.split("/")[-1]}',
                )
                msg.attach(part)

            # Enviar e-mail
            server = smtplib.SMTP('smtp.gmail.com', port=587)
            server.starttls()
            server.login(remetente.get(), senha.get())
            server.send_message(msg)
            server.quit()

            # Exibir mensagem de envio bem-sucedido na interface gráfica
            enviado_para_label.config(text=f"Enviado para {excel_row['email']}")

            # Adicionar um pequeno delay para evitar o bloqueio de anti-spam
            time.sleep(2)  # 2 segundos de atraso

        messagebox.showinfo("Sucesso", "E-mails enviados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao enviar os e-mails: {str(e)}")

# Função para selecionar o arquivo Excel
def selecionar_arquivo():
    filename = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    filepath.set(filename)

# Função para selecionar e anexar um arquivo
def anexar_arquivo():
    filename = filedialog.askopenfilename()
    arquivo_filepath.set(filename)

# Função para mostrar ou ocultar a senha
def mostrar_senha():
    if senha_entry.cget("show") == "":
        senha_entry.config(show="*")
        mostrar_senha_button.config(text="Mostrar Senha")
    else:
        senha_entry.config(show="")
        mostrar_senha_button.config(text="Ocultar Senha")

# Configurações da janela principal
root = tk.Tk()
root.title("Envio de E-mails em Massa")

# Variáveis para armazenar o caminho do arquivo, mensagem, assunto e as credenciais do remetente
filepath = tk.StringVar()
arquivo_filepath = tk.StringVar()
remetente = tk.StringVar()
senha = tk.StringVar()

# Rótulos e campos de entrada
tk.Label(root, text="Selecione o arquivo Excel:").grid(row=0, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=filepath, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_arquivo).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="E-mail do Remetente:").grid(row=1, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=remetente, width=50).grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Senha do Remetente:").grid(row=2, column=0, padx=5, pady=5)
senha_entry = tk.Entry(root, textvariable=senha, show="*", width=50)
senha_entry.grid(row=2, column=1, padx=5, pady=5)

# Botão para mostrar ou ocultar a senha
mostrar_senha_button = tk.Button(root, text="Mostrar Senha", command=mostrar_senha)
mostrar_senha_button.grid(row=2, column=2, padx=5, pady=5)

# Campo para digitar o assunto do e-mail
tk.Label(root, text="Assunto do E-mail:").grid(row=3, column=0, padx=5, pady=5)
assunto_text = tk.Entry(root, width=50)
assunto_text.grid(row=3, column=1, padx=5, pady=5)

# Campo para escrever a mensagem do e-mail
tk.Label(root, text="Mensagem do E-mail:").grid(row=4, column=0, padx=5, pady=5)
mensagem_text = tk.Text(root, height=10, width=50)
mensagem_text.grid(row=4, column=1, padx=5, pady=5)

# Botão para anexar arquivo
tk.Label(root, text="Anexar Arquivo:").grid(row=5, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=arquivo_filepath, width=50).grid(row=5, column=1, padx=5, pady=5)
tk.Button(root, text="Anexar", command=anexar_arquivo).grid(row=5, column=2, padx=5, pady=5)

# Botão para enviar e-mails
tk.Button(root, text="Enviar E-mails", command=enviar_emails).grid(row=6, column=1, padx=5, pady=5)

# Adicionar um rótulo para mostrar "Enviado para"
enviado_para_label = tk.Label(root, text="")
enviado_para_label.grid(row=7, column=1, padx=5, pady=5)

root.mainloop()
