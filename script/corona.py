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
import schedule
import logging

logging.basicConfig(filename='/home/itadmin/logs/corona.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
password = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.OperationalError as e:
    logging.INFO("Error number {0}: {1}.".format(e.args[0],e.args[1]))
    time.sleep(60)
    print("Trying again...")

data_check = open('/home/itadmin/automation/files/coronavirus.txt', 'w+')
query = "SELECT \
                t01.patno \
                FROM HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
                HOSPF0062.CDNOTETB T02 \
                ON T02.ENCID = T01.PATNO LEFT OUTER JOIN \
                HOSPF0062.CDNTEATB T03 \
                ON T02.ENCID = T03.ENCTRID \
                AND T02.CREATEBY = T03.LSTMODBY \
                AND T03.ENCTRID = T01.PATNO LEFT OUTER JOIN \
                hospf0062.rmbed t04 on t04.pat# = t01.patno \
                WHERE T01.isadate > '2020-02-28' and t02.titl like 'Coronavirus%'"

cursor = conn.cursor()
pnumber = cursor.execute(query)
array_pnumber = []
for number in pnumber:
    n = int(number[0])

    array_pnumber.append(n)
    data_check.write(str(number) + "\n")

query2 = "select t01.pat#, t02.isadate, t01.nurst, t01.room, t01.bed  from hospf0062.rmbed t01 \
left outer join hospf0062.patients t02 on t01.pat# = t02.patno where \
nurst != 'DIAG' and nurst != 'NUR' and nurst != 'WHBC' and nurst != 'CVPR' \
and nurst != 'EOP' and PAT# > 0 ORDER BY NURST"

today = datetime.datetime.now()
c = datetime.timedelta(days=1)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
writer = pd.ExcelWriter('/home/itadmin/automation/files/covid19-' + str(yesterday) + '.xlsx', engine='openpyxl')

census = cursor.execute(query2)
new_data = []

output1 = open('/home/itadmin/automation/files/trash-data.txt', 'w+')
for data in census:
    numbers = int(data[0])
    if numbers in array_pnumber:
        out1 = str(data[0]) + "| " + str(data[1]) + "| " + str(data[2]) + "| " + str(data[3]) + "| " + str(data[4]) +"\n"
        output1.write(out1)
    else:
        out = str(data[0]) + "| " + str(data[1]) + "| " + str(data[2]) + "| " + str(data[3]) + "| " + str(data[4]) +"\n"
        new_data.append(out)

ccu = []
cdu = []
mcu = []
pcu = []
scuj = []
ssu = []

for row in new_data:
    nurse_station = row.split("|")[2]
    if 'CCU' in str(nurse_station):
        ccu.append(row)
    elif 'CDU' in str(nurse_station):
        cdu.append(row)
    elif 'MCU' in str(nurse_station):
        mcu.append(row)
    elif 'PCU' in str(nurse_station):
        pcu.append(row)
    elif 'SCUJ' in str(nurse_station):
        scuj.append(row)
    else:
        ssu.append(row)

dataframe = pd.DataFrame(ccu)
dataframe1 = pd.DataFrame(cdu)
dataframe2 = pd.DataFrame(mcu)
dataframe3 = pd.DataFrame(pcu)
dataframe4 = pd.DataFrame(scuj)
dataframe5 = pd.DataFrame(ssu)

dataframe.to_excel(writer, sheet_name="CCU", index=False, header=False)
dataframe1.to_excel(writer, sheet_name="CDU", index=False, header=False)
dataframe2.to_excel(writer, sheet_name="MCU", index=False, header=False)
dataframe3.to_excel(writer, sheet_name="PCU", index=False, header=False)
dataframe4.to_excel(writer, sheet_name="SCUJ", index=False, header=False)
dataframe5.to_excel(writer, sheet_name='SSU', index=False, header=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "covid19 Report For: " + str(today)
recipients = ['setdud@mckweb.com', 'MWMC.House.Coordinators@MCKweb.com', 'TanPar@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/covid19-' + str(yesterday) + '.xlsx'

msg.attach(MIMEText("COVID19 Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()
print("Corona mail Sent On " + str(today))
