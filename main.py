import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

excel_filepath = r'C:\Users\Name\Documents\emails.xlsx'    # Enter filepath of excel containing emails
column = 'Emails'                                          # Enter title of column containing emails
excel_emails = pd.read_excel(excel_filepath)
emails = excel_emails[column].values

# Email Parameters
sender = 'sender@email.com'          # Enter sender email
password = input('Enter password')   # Enter password when running program
subject = 'subject'                  # Enter email subject
content = 'message'                  # Enter email message
image_filepath = r'C:\Users\Name\Pictures\Saved Pictures\coconut.jpg'   # Enter filepath of image to be embedded (type 'none' if no image)

# Create Email
msg = MIMEMultipart('related')
msg['From'] = sender
msg['Subject'] = subject
msg_alt = MIMEMultipart('alternative')
msg.attach(msg_alt)
msg_text = MIMEText(content, 'plain', 'utf-8')
msg_alt.attach(msg_text)
msg_text = MIMEText(content + '<br><img src="cid:image1"><br>', 'html', 'utf-8')   # Use HTML to change content settings (bold, italicize)
msg_alt.attach(msg_text)

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

# Send Emails
for email in emails:
    msg['To'] = email
    server.sendmail(sender, email, msg.as_string())

# Close Mail Server
server.close()
print('Emails sent')
