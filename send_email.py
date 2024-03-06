import os
import smtplib
from email.message import EmailMessage

email = 'leonardobaptista.devcursos@gmail.com'

with open('auth.txt') as f:
    senha = f.readlines()

    f.close()

senha_do_email = senha[0]

msg = EmailMessage()
msg['Subject'] = "teste email com python"
msg['From'] = 'leonardobaptista.devcursos@gmail.com'
msg['To'] = 'leonardobaptista.dev@gmail.com'
msg.set_content("Segue o relatório diário")


with open ('dados.xlsx', 'rb') as content_file:
    content = content_file.read()
    msg.add_attachment(content, maintype='application', subtype='xlsx', filename='dados.xlsx')

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:

    smtp.login(email, senha_do_email)
    smtp.send_message(msg)


