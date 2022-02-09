import pandas as pd
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

dados = pd.read_excel("emails.xlsx")

for i, j in dados.iterrows():
    try:
        remetente = "E=mail"
        destinatario = j['Endereço']
        senha = 'senha'
        msg = MIMEMultipart()

        msg['From'] = remetente 
        msg['To'] = destinatario
        msg['Subject'] = "Teste"

        body = "\nTeste com imagem"

        msg.attach(MIMEText(body, 'plain'))

        filename = 'arquivo.pdf'

        attachment = open('arquivo.pdf','rb')


        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        attachment.close()

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        text = msg.as_string()
        server.sendmail(remetente, destinatario, text)
        server.quit()
        print('\nEmail enviado com sucesso!')
    except:
        print("\nErro ao enviar email")