import os
import smtplib
from email.message import EmailMessage

EMAIL_ADDRESS = os.environ.get('EMAIL_USER')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')  # Email Password for python

RECEIVER = 'acompanamientobecas@patronatobcp.org'

msg = EmailMessage()
msg['Subject'] = 'Recibo de conformidad'
msg['From'] = EMAIL_ADDRESS
msg['To'] = RECEIVER
msg.set_content('Hola, adjunto el recibo de conformidad del correspondiente mes')

with open('recibo_de_conformidad.pdf', 'rb') as f:
    file_data = f.read()
    file_name = f.name

msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:

    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)
