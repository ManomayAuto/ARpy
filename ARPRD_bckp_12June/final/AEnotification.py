from configparser import ConfigParser
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pyodbc
import os
import xlsxwriter
from datetime  import  date
import time
import datetime

#get emailid and password from config file
def get_email_data():           
    config = ConfigParser()
    config.read('email_config.ini')
    email_details = config["emaildata"]
    if email_details["encrypted"] == "No":
        email = email_details["emailid"]
        password = email_details["password"]
        return email, password
    elif email_details["encrypted"] == "Yes":
        return email_details["emailid"], Encryption().decrypt(email_details["password"])
   

cnxn = pyodbc.connect("Driver={SQL Server};"
                          "Server=Inbroker4;"
                          "Database=MantraDB;"
                          "Trusted_Connection=yes;")
cur = cnxn.cursor()
            
cur.execute("""Select accexe as AEName,count(PolicyNum) as PolicyCount,u.Email  as AEemailid from ae 
join Users_man u on ae.AccExe = u.Name
where ReductionCheckBox = 0 and Followups = 0 --and DateCreated < = cast(getdate() as date) 
group by AccExe,u.Email""")
records = cur.fetchall()
cur.commit()
for record in records:
    aeName=record[0]
    polCnt=record[1]
    emailae=record[2]
    if emailae is None:
        emailae=''
    email,password=get_email_data()
    sender = email
    receivers = emailae

    msg = MIMEMultipart()
    msg['subject'] = "Task Notification"
    message = """<p>Dear """+str(aeName)+"""</p>

    <p>There are tasks pending to be acted upon by you.</p>

    <p>Currently, there are """+str(polCnt)+""" Policies that might require your attention.</p>

    <p>Please ensure you are taking appropriate timely actions for each.</p>

    <p>Regards</p>
    <p>Team</p>
    <p>J.S JOHNSON</p>
    <p>Web: www.jsjohnson.com</p>
    <p>Tel: (242) 397-2100|Fax: (242) 323-3720</p>

    <p style="color:red;">**This is an auto-generated email, please do not respond to this</p>
    """
    msg.attach(MIMEText(message,'html'))
    try:
        smtpObj = smtplib.SMTP('mail.jsjemail.com')
        smtpObj.login(email,password)
        smtpObj.sendmail(sender, receivers, msg.as_string())
        print("Successfully sent email")
        
    except smtplib.SMTPException as e:
        print("Error: unable to send email")
        print(e)
        
        
