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
import configparser

logging.config.fileConfig('logconfig.ini')
config = configparser.ConfigParser()
config.read('Config.ini')

smtp_addr = config.get('CREDENTIALS','SMTPADDR',fallback='SMTPADDR not defined') #'smtp.cdac.in'
smtp_port = int(config.get('CREDENTIALS','SMTPPORT',fallback='SMTPPORT not defined'))
sender_id = config.get('CREDENTIALS','USERNAME',fallback='USERNAME not defined')
sender_pass = config.get('CREDENTIALS','PASS',fallback='PASS not defined') #getpass('Enter the mail password: ')#os.environ['CDAC_PASS'] #password
sender_email = config.get('CREDENTIALS','EMAILID',fallback='EMAILID not defined') #input('Enter the full email address (Eg: iotcourses@cdac.in): ')#os.environ['CDAC_EMAIL'] #spriyanka@cdac.in

#mail header setup
mailsubject = config.get('MAILCONFIG','SUBJECT',fallback='SUBJECT not defined')
mailcc = config.get('MAILCONFIG','MAILCC',fallback='MAILCC not defined')
idfile = config.get('MAILCONFIG','IDFILE',fallback='IDFILE not defined')
messagefilename = config.get('MAILCONFIG','MAILTEXTFILE',fallback='MAILTEXTFILE not defined')
#open csv file contaning attachment details
try:
    ids = pd.read_csv(idfile)
    logging.info(str(datetime.now()) + ' INFO ' + ' FILE READ ' + idfile)
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
        try:

            participantname = id['Name']
            participantemail = id['Email']

            mailbody = mailmessage
            mailbody = mailbody.replace('~ParticipantName~', participantname)
            
            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = participantemail
            message['Subject'] = mailsubject
            message['Cc'] = mailcc
            
            mailtexttype = config.get('MAILCONFIG','MAILTEXTTYPE',fallback='MAILTEXTTYPE not defined')
            if mailtexttype =='html':
                message.attach(MIMEText(mailbody, 'html'))#.set_content(MIMEText(mailbody, 'html'))
            else:
                message.attach(MIMEText(mailbody, 'plain'))
            
            
            pdfname = id['AttachmentPath']
        
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