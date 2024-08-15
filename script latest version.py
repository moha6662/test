import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import getpass
import time

#Helper FUnctions


def convert_csv_to_string_commas_seperated(emails_list):
    text = ''
    for email in emails_list:
        text += email + ','
    text = text[:len(text)-1]

    return text


def get_data():
    print('----Add the needed data for the email Server----')
    smtp_server = '10.240.48.36'
    smtp_port = 587
    smtp_user = input('Enter the user here:')
    smtp_password = getpass.getpass("Enter your password: ")


    # Source files:
    body_src_file = r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Body.txt'           #r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Body.txt'
    with open(body_src_file, "r") as fileBody:
        body = fileBody.read()

    subject_src_file =  r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Subject.txt'     #r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Subject.txt'
    with open(subject_src_file, "r") as fileSubject:
        subject = fileSubject.read()

    Source_file = r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Sources.csv'           #r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Sources.csv'
    destinations = pd.read_csv(Source_file)
    
    file_attachement = r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Recipients.csv'


    employee_name = destinations['Name'].tolist()
    employee_email = destinations['Email'].tolist()


    CCs = []

    for row in destinations['CCs'].tolist():
        new_row = row.split('/')
        CCs.append(new_row)

    data = {
        'server': smtp_server,
        'port': smtp_port,
        'user': smtp_user,
        'password': smtp_password,
        'body': body,
        'subject': subject,
        'names':employee_name,
        'destinations': employee_email,
        'attach':file_attachement,
        'CC':CCs,
    }
    print(CCs)
    return data

#get_data()

def sender():


    ccc = "EG.ADIB.InfoSecMonitoring@adib.eg,Eg_SecurityIncident@adib.eg"
   
    data = get_data()
    employee_name = data['names']
    employee_email = data['destinations']
    managers = data['CC']
    filePath = data['attach']

    body = data['body']

    # print(employee_name)
    # print(employee_email)
    #print(managers)

    # Connect to the server and send the email
    try:
        server = smtplib.SMTP(data['server'], data['port'])
        server.starttls()
        server.login(data['user'], data['password'])

        for counter in range(len(employee_name)):
            ccs = ",".join(managers[counter])
            recipients = managers[counter] + [employee_email[counter]]
            print(recipients)
            print('logged in')
            # Create the email message
            msg = MIMEMultipart('alternative')
            msg['From'] = 'EG.ADIB.InfoSecMonitoring@adib.eg'
            msg['To'] = employee_email[counter]
            msg['Cc'] = ccs
            msg['Subject'] = data['subject']
            dearSentence = 'Dear ' + employee_name[counter] + ',<br><br>'
            bodySend = dearSentence + body
            msg.attach(MIMEText(bodySend, 'html'))
            
            with open(filePath, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition','attachment;filename=Recipients.csv')
                msg.attach(part)
                
            server.sendmail(data['user'], recipients, msg.as_string())

            print(f'Email to {employee_email[counter]} sent successfully!')


        server.quit()


    except Exception as e:
        print(f'Failed to send email: {e}')


#get_data()
sender()