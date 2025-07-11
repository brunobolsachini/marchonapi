import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

remetente = 'bruno@compreoculos.com.br'
destinatario = 'bruno@compreoculos.com.br'
senha = os.getenv('SENHA_API')

if not senha:
    print('❌ SENHA_API não está definida no ambiente!')
    exit(1)

msg = MIMEMultipart()
msg['From'] = remetente
msg['To'] = destinatario
msg['Subject'] = 'Teste de envio de e-mail via GitHub Actions'
mensagem = '✅ Este é um teste de envio de e-mail via GitHub Actions.'
msg.attach(MIMEText(mensagem, 'plain'))

try:
    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()
    servidor.login(remetente, senha)
    servidor.sendmail(remetente, destinatario, msg.as_string())
    servidor.quit()
    print('✅ E-mail enviado com sucesso!')
except Exception as e:
    print(f'❌ Erro ao enviar e-mail: {e}')
    exit(1)
