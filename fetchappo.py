# Importing Libraries. We need imaplib to work with O365 email, email for comprehending the email components into 
# readable format, and pandas to create dataframe out of the email components
import imaplib
from sqlalchemy import create_engine
import email
import pandas as pd
import smtplib
from email.mime.text import MIMEText
import datetime
import requests
from email.mime.multipart import MIMEMultipart
from pretty_html_table import build_table

# DB Credentials
user = ''
password = ''
host =  ''
port = 3306
database = ''

# Setting up DB connection
engine = create_engine('mysql+pymysql://' + user + ':' + password + '@' + host + '/' + database)

# email credentials
imap_host = 'imap.gmail.com'
imap_user = ''
imap_pass = ''

# connect to host using SSL
mailserver = smtplib.SMTP('smtp.gmail.com',587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login('', imap_pass)

query = "SELECT email_id,first_name,pincode FROM user_regn where status = 1"

df = pd.read_sql_query(query,engine)

uniq_pin_code = df.pincode.unique()

start_date = date.today()

date_list = []
slots = pd.DataFrame()

for n in range(0,7):
    date_list.append(start_date+datetime.timedelta(days=n))

for i in date_list:
    api_date = i.strftime("%d-%m-%Y")
    #print("Date: " + api_date)
    for pin in uniq_pin_code:
        #print("Pin Code: " + pin)
        response = requests.get("https://cdn-api.co-vin.in/api/v2/appointment/sessions/public/findByPin?pincode="+ pin + "&date=" + api_date + '"')
        res = response.json()
        df1 = pd.DataFrame.from_dict(res['sessions'])
        if not df1.empty:
            slots = pd.concat([slots,df1])

slots.drop(["center_id","block_name","from","to","lat","long","session_id"],axis=1,inplace=True)
slots_mod = slots[["date","pincode","min_age_limit","vaccine","fee_type","name","address","state_name","district_name","available_capacity","slots"]].sort_values(by=['date'],ascending = True)

for i in df.index:
    to = df['email_id'][i]
    receipient = df['first_name'][i]
    #print(toaddrs)
    #print(df['pincode'][i])
    df_mail = slots_mod.loc[slots_mod['pincode'] == int(df['pincode'][i])]
    #print(df_mail)
    if not df_mail.empty:
        df_mail = df_mail.rename(columns={'date': 'Date', 'pincode': 'PIN','min_age_limit':'Minimum Age Limit','vaccine':'Vaccine Type','fee_type':'Free/Paid','name':'Center Name','address':'Address','state_name':'State','address':'Address','district_name':'District','available_capacity':'Capacity','slots':'Available Slots'})
        html_table = build_table(df_mail, 'blue_light',font_size = 'small', text_align = 'middle')
        msg = MIMEText("Hi "+ receipient + ",<br><br>Please find the vaccine appointment details below for next 7 days. Kindly drop a mail to email@email.com with subject line COWIN&ltspace&gtSTOP to unsubscribe from regular updates.<br><br>" + html_table + "<br><br>Regards,<br>Vaccine Saathi",'html')
        msg['Subject'] = 'KMS Update: Vaccine Appointments Available'
        msg['From'] = "Vaccine Saathi<email@email.com>"
        mailserver.sendmail(msg['From'], to , msg.as_string())
    
