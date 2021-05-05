# Importing Libraries. We need imaplib to work with O365 email, email for comprehending the email components into 
# readable format, and pandas to create dataframe out of the email components
import imaplib
import sqlalchemy
from sqlalchemy.sql import text
from sqlalchemy import create_engine
import email
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# DB Credentials
user = ''
password = ''
host =  ''
port = 3306
database = ''

# Setting up DB connection
engine = create_engine('mysql+pymysql://' + user + ':' + password + '@' + host + '/' + database)
con = engine.connect()

# constructing blank dictionary to push registration data as fetched from email
msg_dict = {"email_id":[],"first_name":[],"last_name":[],"pincode":[],"status":[]}

# email credentials
imap_host = 'imap.gmail.com'
imap_user = 'email@email.com'
imap_pass = 'PASS'

# connect to host using SSL
imap = imaplib.IMAP4_SSL(imap_host)

mailserver = smtplib.SMTP('smtp.gmail.com',587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login('email@email.com', 'PASS')

## login to server
imap.login(imap_user, imap_pass)

imap.select('Inbox')

# Iterating through all unread messaged with a subject line containing COWIN to collect registration data
tmp, data = imap.search(None, '(SUBJECT "COWIN")',"UNSEEN")
for num in data[0].split():
    typ, data = imap.fetch(num, '(RFC822)' )
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            email_subject = msg['subject']
            email_from = msg['from']
            
            # Fetching correct details and de-registering the user
            if len(email_subject.split(" ")) > 2 and email_subject.split(" ")[1].upper() == "STOP" and isinstance(int(email_subject.split(" ")[2]),int) == True:
                
                sql_query = text("delete from user_regn where email_id = '"+ email_from.split('<')[1].split('>')[0].lower() +"'")
                print(sql_query)
                con.execute(sql_query)
                
                msg_dict["email_id"].append(email_from.split('<')[1].split('>')[0].lower())
                msg_dict["first_name"].append(email_from.split('<')[0].split(" ")[0].replace('"',''))
                msg_dict["last_name"].append(email_from.split('<')[0].split(" ")[1].replace('"',''))
                msg_dict["pincode"].append(email_subject.split(" ")[2])
                msg_dict["status"].append(1)
                msg = MIMEText('Dear '+email_from.split('<')[0].split(" ")[0].replace('"','')+',<br><br>Your details have been deleted for pin code ' + email_subject.split(" ")[2] + '. Thanks for using our service. Please drop your feedback to email@email.com.<br><br>Thanks,<br>Vaccine Saathi', 'html')
                msg['Subject'] = 'KMS:De-Registration Confirmation'
                msg['From'] = "Vaccine Saathi<email@email.com>"
                msg['To'] = email_from.split('<')[1].split('>')[0].lower()
                mailserver.sendmail('email@email.com',email_from,msg.as_string())
                imap.store(num, "+FLAGS", "\\Deleted")
                
# Clean-up tasks        
imap.expunge()
mailserver.quit()
imap.close()
