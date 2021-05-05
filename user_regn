# Importing Libraries. We need imaplib to work with O365 email, email for comprehending the email components into 
# readable format, and pandas to create dataframe out of the email components
import imaplib
from sqlalchemy import create_engine
import email
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# DB Credentials
user = 'USER'
password = 'PASSWORD'
host =  'HOST'
port = 3306
database = 'DATABASE'

# Setting up DB connection
engine = create_engine('mysql+pymysql://' + user + ':' + password + '@' + host + '/' + database)

# constructing blank dictionary to push registration data as fetched from email
msg_dict = {"email_id":[],"first_name":[],"last_name":[],"pincode":[],"status":[]}

# email credentials
imap_host = 'imap.gmail.com'
imap_user = 'youremail@domain.com'
imap_pass = 'PASSWORD'

# connect to host using SSL
imap = imaplib.IMAP4_SSL(imap_host)

mailserver = smtplib.SMTP('smtp.gmail.com',587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login('youremail@domain.com', 'PASSWORD')

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
           
            # Handling wrong details supllied over email
            if len(email_subject.split(" ")) < 2 or (email_subject.split(" ")[1].upper() == "START" and isinstance(int(email_subject.split(" ")[2]),int) == False) or any(re.findall(r'START|STOP', email_subject.split(" ")[1].upper())) == False:
                msg = MIMEText('Dear '+email_from.split('<')[0].split(" ")[0].replace('"','')+',<br><br>You have supplied invalid registration details. Kindly send a new email to tn.protiviti@timesgroup.com with a subject line COWIN&ltspace&gtSTART&ltspace&gtPINCODE.<br><br>Thanks,<br>Vaccine Saathi', 'html')
                msg['Subject'] = 'KMS:Invalid Registration Details'
                msg['From'] = "Vaccine Saathi<kyouremail@domain.com"
                msg['To'] = email_from.split('<')[1].split('>')[0].lower()
                mailserver.sendmail('youremail@domain.com',email_from,msg.as_string())
                imap.store(num, "+FLAGS", "\\Deleted")
            
            # Fetching correct details and registering the user
            elif len(email_subject.split(" ")) > 2 and email_subject.split(" ")[1].upper() == "START" and email_subject.split(" ")[1].upper() != "STOP" and isinstance(int(email_subject.split(" ")[2]),int) == True:
                
                if len(engine.execute("SELECT * FROM user_regn where email_id = '" + email_from.split('<')[1].split('>')[0].lower() + "' and pincode = '" + email_subject.split(" ")[2] + "'").fetchall()) == 0:
                    chk_email = ''
                else:
                    chk_email = engine.execute("SELECT * FROM user_regn where email_id = '" + email_from.split('<')[1].split('>')[0].lower() + "' and pincode = '" + email_subject.split(" ")[2] + "'").fetchall()[0][0]
                
                if email_from.split('<')[1].split('>')[0].lower() == chk_email:
                    msg = MIMEText('Dear '+email_from.split('<')[0].split(" ")[0].replace('"','')+',<br><br>We have your details already registered. We will notify you with the available slots once it is available in your area.<br><br>Thanks,<br>Vaccine Saathi', 'html')
                    msg['Subject'] = 'KMS:Duplicate Registration'
                    msg['From'] = "Vaccine Saathi<youremail@domain.com>"
                    msg['To'] = email_from.split('<')[1].split('>')[0].lower()
                    mailserver.sendmail('youremail@domain.com',email_from,msg.as_string())
                    imap.store(num, "+FLAGS", "\\Deleted")
                else:
                    msg_dict["email_id"].append(email_from.split('<')[1].split('>')[0].lower())
                    msg_dict["first_name"].append(email_from.split('<')[0].split(" ")[0].replace('"',''))
                    msg_dict["last_name"].append(email_from.split('<')[0].split(" ")[1].replace('"',''))
                    msg_dict["pincode"].append(email_subject.split(" ")[2])
                    msg_dict["status"].append(1)
                    msg = MIMEText('Dear '+email_from.split('<')[0].split(" ")[0].replace('"','')+',<br><br>Your details have been registered. We will notify you with the available slots once it is available in your area.<br><br>Thanks,<br>Vaccine Saathi', 'html')
                    msg['Subject'] = 'KMS:Registration Confirmation'
                    msg['From'] = "Vaccine Saathi<youremail@domain.com>"
                    msg['To'] = email_from.split('<')[1].split('>')[0].lower()
                    mailserver.sendmail('youremail@domain.com',email_from,msg.as_string())
                    imap.store(num, "+FLAGS", "\\Deleted")
                
# Clean-up tasks        
imap.expunge()
mailserver.quit()
imap.close()

# Creating the dataframe with the details of users
df = pd.DataFrame(msg_dict) 

# Inserting the dataframe in MySQL database
df.to_sql(name='user_regn', con=engine, if_exists = 'append', index=False)

