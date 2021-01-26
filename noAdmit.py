import pyodbc, csv, logging, os, smtplib
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import configparser
import schedule
import time

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')

server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
email = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

census = []
admit_order = []

logging.basicConfig(filename='/home/itadmin/logs/noadmit.log', level=logging.DEBUG)

#connect to as400 database
conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
cursor = conn.cursor()

census_query = "select t01.PAT# from hospf062.rmbed t01\
    inner join hospf062.patients t02 on t01.pat# = t02.patno where t02.hssvc='MIP'\
    or t02.hssvc='SIP' or t02.hssvc='OBS' or t02.hssvc='OBI' or t02.hssvc='NUR' or\
    t02.hssvc='BBN' or t02.hssvc='ICU' or t02.hssvc='CCU' order by NURST"

admit_query = "select t01.pat# from hospf062.rmbed t01\
    inner join hospf062.patients t02 on t01.pat# =t02.patno inner join \
    orderf062.oeorder t03 on t02.patno=t03.opat# WHERE oproc='ADMHOSP' \
    or oproc='ADMIPB1R'\
    or oproc='ADMIPOPR' or oproc ='ADMIP2DR' or oproc='ADMITIP' or \
    oproc='ADMOBINP' or oproc='ADMOBOBS' or oproc='IPSCU' or oproc='OBSERV'\
    or oproc='ERINPT' or oproc='EROBS' order by nurst"
try:
    cursor.execute(census_query)
except pyodbc.Error as e:
    logging.INFO(e)
    sleep(60)

for row in cursor:
    census.append(row)

cursor.execute(admit_query)

for row in cursor:
    admit_order.append(row)

data = open('/home/itadmin/automation/files/match.txt', 'w+')
data_not = open('/home/itadmin/automation/files/nomatch.txt', "w+")
conn.close()

    # census_array = []
admit_array = []
for row in census:
    d1 = str(row[0])
    if row in admit_order:
        data.write(d1 + "\n")
    else:
        admit_array.append(d1)
            #data_not.write(d1)
            #data_not.write("\n")
data_not.close()
data.close()

today = datetime.now().date()
conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
cur = conn.cursor()
report = open('/home/itadmin/automation/files/report.txt', 'w+')
report_array = []
    # data_read = open('nomatch.txt', 'r+')
report.write("MISSING ADMIT ORDERS For " + str(today) + "\n")
report.write("===============================================\n\n\n")
report.write("Nurse St      Room         Bed     Patient Name             Account Number\n\n")
for row in admit_array:
    query = "select t01.nurst, t01.room, t01.bed, t02.pname, t02.patno from hospf062.rmbed\
    t01 inner join hospf062.patients t02 on t01.pat# = t02.patno where t01.pat#='" +\
    row + "' order by t01.nurst"
    cur.execute(query)
    for r1 in cur:
        output = str(r1[0]) + "      " + str(r1[1]) + "         " + str(r1[2]) +"     " + str(r1[3]) + "             " + str(r1[4]) + "\n"
        report.write(output)
        report_array.append(output)
report.close()

msg = MIMEMultipart()
msg['Subject'] = "No Admit Order Report for " + str(today)
recipients = ['setdud@mckweb.com', 'terhur@MCKweb.com', \
    'RBootes@MCKWeb.com', \
    'SarSor@MCKweb.com', 'MadGue@MCKweb.com', 'KimSna@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/report.txt'

msg.attach(MIMEText("Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(email)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()
print("No Admit Report email sent: " + str(today))
