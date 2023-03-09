import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from dotenv import load_dotenv
load_dotenv()

server = smtplib.SMTP("smtp.rambler.ru", 2525)

def sendEmail(attachFile):
    user = os.getenv('USER_MAIL')
    password = os.getenv('PASSWORD')
    catcher = os.getenv('CATCHER')

    try:
        server.login(user, password)
        msg = MIMEMultipart()

        with open(f'email/{attachFile}', 'rb') as f:
            file = MIMEApplication(f.read(), 'docx')
        #print(file)
        msg['Subject'] = 'Тест'
        msg.attach(MIMEText('Привет'))
        msg.attach(file)
        file.add_header('content-disposition', 'attachment', filename=f'{attachFile}')
        server.sendmail(user, catcher, msg.as_string())
        print(f'send Email whit file: {attachFile}')
    except Exception as Err:
        print(f'{Err}')

