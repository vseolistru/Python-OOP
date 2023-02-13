import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

server = smtplib.SMTP("smtp.rambler.ru", 2525)

def sendEmail(attachFile):
    user = "v_se0@ro.ru"
    password = "pass"
    catcher = 'v_seo@list.ru'

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