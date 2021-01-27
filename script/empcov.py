import pandas as pd
import configparser
import pyodbc
import xlsxwriter
import time
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
password = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, PWD=password)

day = datetime.datetime.now()
today = day.strftime('%y%m%d')
a = datetime.timedelta(days=1)
b = day - a
yesterday = b.strftime('%y%m%d')

query = "SELECT t01.iapnam, t02.harcd, t02.phone, t03.rdts, t03.rdrs, t04.osdate \
from hospf0062.indaccum t01 left outer join \
hospf0062.pathist t02 on t01.iahst# = t02.histn left outer join \
orderf0062.rd t03 on t01.iaord# = t03.rdor# and t01.iaacct = t03.rdpt# left outer join \
orderf0062.oeorder t04 on t01.iaacct = t04.opat# and t01.iaord# = t04.oord# \
Where t04.optlst = 'MWMC' and t03.rdts = 'COV-CM' and t04.osdate = '" + yesterday + "' order by t04.osdate"


data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
dataframe.to_excel('/home/itadmin/automation/files/employee-test-' + today + '.xlsx', index=False)

msg = MIMEMultipart()
msg['Subject'] = 'Covid-19 Employee Test Report'

# recipients to send the email to
recipients = ['setdud@mckweb.com', 'tanpar@mckweb.com', 'DesShu@MCKweb.com', 'racgei@mckweb.com']

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/employee-test-' + today + '.xlsx'

msg.attach(MIMEText("COVID19 Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()
