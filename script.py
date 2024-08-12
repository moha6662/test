import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def get_data():
    print('----Add the needed data for the email Server----')
    smtp_server = '10.240.48.36'
    smtp_port = 587
    smtp_user = input('Enter the user here:')
    smtp_password = input('Enter the password here:')

    # Source files:
    body_src_file =r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Body.txt'
    with open(body_src_file, "r") as fileBody:
        body = fileBody.read()

    subject_src_file = r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Subject.txt'
    with open(subject_src_file, "r") as fileSubject:
        subject = fileSubject.read()

    destination_email_src_file = r'C:\Users\T1Egsnoc3\Downloads\Script\Resources\Recipients.csv'
    destinations = pd.read_csv(destination_email_src_file)
    destinations = destinations['emails'].tolist()

    data = {
        'server': smtp_server,
        'port': smtp_port,
        'user': smtp_user,
        'password': smtp_password,
        'body': body,
        'subject': subject,
        'destinations': destinations
    }

    #print(data)
    return data

#get_data()

def sender():
    data = get_data()
    destination_emails = data['destinations']

    # Connect to the server and send the email
    try:
        server = smtplib.SMTP(data['server'], data['port'])
        server.starttls()
        server.login(data['user'], data['password'])


        for email_ in destination_emails:
            print('longed in')
            destination_email = email_ 
            
            # Create the email message
            msg = MIMEMultipart()
            #msg['From'] = data['user']
            msg['From'] = "EG.ADIB.InfoSecMonitoring@adib.eg"
            msg['CC'] = "karim.hosny@adib.eg"
            msg['To'] = destination_email
            msg['Subject'] = data['subject']
            msg.attach(MIMEText(data['body'], 'plain'))

            server.sendmail(data['user'], destination_email, msg.as_string())
            print(f'Email to {destination_email} sent successfully!')

        server.quit()


    except Exception as e:
        print(f'Failed to send email: {e}')



sender()