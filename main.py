#!/usr/bin/env python

import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

excel_filepath = r'C:\Users\Name\Documents\emails.xlsx'    # Enter filepath of excel containing emails
column = 'Emails'                                          # Enter title of column containing emails
excel_emails = pd.read_excel(excel_filepath)
emails = excel_emails[column].values
session = smtplib.SMTP('smtp.gmail.com:587')

# Email Parameters
sender = 'sender@email.com'          
password = input('Enter password')   
subject = 'subject'                 
content = 'message'                  
image_filepath = r'C:\Users\Name\Pictures\Saved Pictures\coconut.jpg'   

# Create Email
msg = MIMEMultipart('related')
msg['From'] = sender
msg['Subject'] = subject
msg_alt = MIMEMultipart('alternative')
msg.attach(msg_alt)
msg_text = MIMEText(content, 'plain', 'utf-8')
msg_alt.attach(msg_text)
msg_text = MIMEText(content + '<br><img src="cid:image1"><br>', 'html', 'utf-8')   
msg_alt.attach(msg_text)
messages = []

# Embed Image into Email
if "none" not in image_filepath:
    filepath = open(image_filepath, 'rb')
    msg_image = MIMEImage(filepath.read())
    filepath.close()
    msg_image.add_header('Content-ID', '<image1>')
    msg.attach(msg_image)

# Create Mail Server
if 'gmail' in sender:
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
elif 'outlook' in sender:
    server = smtplib.SMTP(host='smtp.office365.com', port=587)
elif 'yahoo' in sender:
    server = smtplib.SMTP(host='smtp.mail.yahoo.com', port=465)
elif 'hotmail' in sender:
    server = smtplib.SMTP(host='mail.outlook.com', port=587)
else:
    print('Unsupported email address')

server.starttls()
server.login(sender, password)

for email in emails:
    msg['To'] = email
    server.sendmail(sender, email, msg.as_string())

for message in messages:
    receiver = message[0]
    headers = ["from: " + sender,
            "subject: " + subject,
            "to: " + receiver,
            "mime-version: 1.0",
            "content-type: text/html"]
    headers = "\r\n".join(headers)
    body = message[1]
    session.sendmail(sender, receiver, headers + "\r\n\r\n" + body)

server.close()
print('Emails sent')
