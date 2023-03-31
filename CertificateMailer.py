from asyncio.windows_events import NULL
from datetime import datetime
from getpass import getpass
import logging
import logging.config
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from stdiomask import getpass

logging.config.fileConfig('logconfig.ini')

smtp_addr = 'smtp.cdac.in' #os.environ['CDAC_SMTP'] #smtp.cdac.in
smtp_port = 25 #int(os.environ['CDAC_SMTP_PORT']) #25
sender_id = input('Enter sender user id(Eg: spriyanka): ')#'madhurin' #os.environ['CDAC_ID'] #spriyanka
sender_pass = getpass('Enter Sender Password : ')#'' #os.environ['CDAC_PASS'] #password
sender_email = input('Enter Sender Complete mail id(Eg: spriyanka@cdac.in): ')#'madhirin@cdac.in' #os.environ['CDAC_EMAIL'] #spriyanka@cdac.in

#mail header setup
mailsubject = input('Enter mail Subject: ')
mailcc = input('Enter list of mail ids to be kept in CC: ')
csvname = input('Enter the complete name of CSV file to be selected (Eg: AttachmentListTemp.csv): ')
messagefilename = input('Enter the complete name of text file containing mail message (Eg: MailMessage.txt): ')
#open csv file contaning attachment details
try:
    ids = pd.read_csv(csvname)
    logging.info(str(datetime.now()) + ' INFO ' + ' FILE READ ' + csvname)
except Exception as e:
    print(e)
    logging.error(str(datetime.now()) + ' ERROR ' + str(e))
    ids = NULL

#read Mail Message from text file
try:
    file_ptr = open(messagefilename, 'r')
    mailmessage = file_ptr.readlines()
    mailmessage = ''.join(mailmessage)
    file_ptr.close()
    logging.info(str(datetime.now()) + ' INFO ' + ' FILE READ ' + messagefilename)
except Exception as e:
    print(e)
    logging.error(str(datetime.now()) + ' ERROR ' + str(e))
    mailmessage = NULL

errormailids = []

if not ids.empty and mailmessage:
    for index, id in ids.iterrows():
        participantname = id['Name']
        participantemail = id['Email']

        mailbody = mailmessage
        mailbody = mailbody.replace('<<ParticipantName>>', participantname)
        
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = participantemail
        message['Subject'] = mailsubject
        message['Cc'] = mailcc
        message.attach(MIMEText(mailbody, 'plain'))
        
        pdfname = id['AttachmentPath']
        try:
            # open the file in bynary
            binary_pdf = open(pdfname, 'rb')

            payload = MIMEBase('application', 'octate-stream', Name=pdfname)
            # payload = MIMEBase('application', 'pdf', Name=pdfname)
            payload.set_payload((binary_pdf).read())

            # enconding the binary into base64
            encoders.encode_base64(payload)

            # add header with pdf name
            payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
            message.attach(payload)
            logging.info(str(datetime.now()) + ' INFO ' + ' FILE ATTACHED ' + pdfname)

            #create smtp session
            session = smtplib.SMTP(host=smtp_addr, port=smtp_port)
            session.connect(host=smtp_addr, port=smtp_port)
            #enable security
            #session.starttls()

            #login with mail_id and password
            session.login(sender_id, sender_pass)

            #read message body and send mail
            text = message.as_string()
            session.sendmail(sender_email, participantemail.split(',')+mailcc.split(', '), text)
            session.quit()
            print('Mail Sent to: ', participantemail)
            logging.info(str(datetime.now()) + ' INFO ' + ' MAIL SENT ' + participantemail)  
        except Exception as e:
            print(e)
            logging.error(str(datetime.now()) + ' ERROR ' + str(e))
            errormailids.append(participantemail)     
    
    #write error mail ids to text file
    file_ptr = open('ErrorIDLog.txt', 'w')
    file_ptr.write(', '.join(errormailids))
    file_ptr.close()   

logging.info(str(datetime.now()) + ' INFO ' + ' Task Complete')
print('-'*50)    
print('Task Completed. Check ErrorIDLog.txt for error mail ids. Check MailLog for Details.')