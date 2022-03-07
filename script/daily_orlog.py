import pyodbc
import configparser
import xlsxwriter
import pandas as pd
import pymssql
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os

config = configparser.ConfigParser()
config.read('D:\\2-PROD\\config.ini')
server = config['PHS']['SERVER']
user = config['PHS']['USER']
password = config['PHS']['PASS']
eserv = config['O365']['SERVER']
euser = config['O365']['USER']
epass = config['O365']['PASS']

conn = pymssql.connect(server=server, user=user, password=password, database='asmprod')
day = datetime.datetime.now()
today = day.strftime('%Y-%m-%d')
a = datetime.timedelta(days=1)
b = day - a
yesterday = b.strftime('%Y-%m-%d')
# c = datetime.timedelta(days=32)
#e = b - c
c = datetime.datetime.today().replace(day=1)
d = c.strftime('%Y-%m-%d')
mtd_start_date = d + " 00:00:000"

query = "select cm.actcase_start_datetime, cm.pat_acct_num, cm.pat_mrn, cm.pat_displayname, \
convert(date, cm.pat_birth_datetime) as birthdate, \
pt.abbr, cio.pat_or_in_datetime, cio.pat_or_out_datetime, cp.primpract_res_name, \
cp.actual_proname as 'Procedure', cm.schedcase_start_datetime, cm.schedcase_stop_datetime \
from casemain cm left outer join psmresindroom pri on \
cm.actualroom_res_id = pri.res_id \
left outer join casepreop cpo on cm.casemain_id = cpo.casemain_id \
left outer join psmpattype pt on cpo.pattype_id = pt.pattype_id \
left outer join casepro cp on cm.casemain_id = cp.casemain_id \
left outer join caseintraop cio on cm.casemain_id = cio.casemain_id \
where cm.actcase_start_datetime >= '" + yesterday + " 00:00:00' and cm.actcase_start_datetime <= '" + yesterday + " 23:59:000 '"

query2 = "select cm.actcase_start_datetime, cm.pat_acct_num, cm.pat_mrn, cm.pat_displayname, \
convert(date, cm.pat_birth_datetime) as birthdate, \
pt.abbr, cio.pat_or_in_datetime, cio.pat_or_out_datetime, cp.primpract_res_name, \
cp.actual_proname as 'Procedure' \
from casemain cm left outer join psmresindroom pri on \
cm.actualroom_res_id = pri.res_id \
left outer join casepreop cpo on cm.casemain_id = cpo.casemain_id \
left outer join psmpattype pt on cpo.pattype_id = pt.pattype_id \
left outer join casepro cp on cm.casemain_id = cp.casemain_id \
left outer join caseintraop cio on cm.casemain_id = cio.casemain_id \
where cm.actcase_start_datetime >= '"+mtd_start_date+"' and cm.actcase_start_datetime <= '" + yesterday + " 23:59:000 '"

data = pd.read_sql(query, conn)
data2 = pd.read_sql(query2, conn)
dataframe = pd.DataFrame(data)
df2 = pd.DataFrame(data2)
out = dataframe.drop_duplicates(subset=['pat_acct_num'], keep='first')
out2 = df2.drop_duplicates(subset=['pat_acct_num'], keep='first')
writer = pd.ExcelWriter('D:\\4-FILES\\dailyOR-log-' + today + '.xlsx', engine='xlsxwriter')
out.to_excel(writer, index=False, sheet_name='Yesterday')
out2.to_excel(writer, index=False, sheet_name='MTD')
writer.save()
conn.close()

filename = 'D:\\4-FILES\\dailyOR-log-' + today + '.xlsx'
msg = MIMEMultipart()
msg['Subject'] = 'Daily OR log Report'
# recipients needs to be edited for each file
recipients = ['sdudenhofer@qhcus.com', 
                'MNelson02@qhcus.com', 
                'BFreybates@qhcus.com', 
                'LRamp@qhcus.com', 
                'tdalstra@qhcus.com', 
                'KOneal@qhcus.com', 
                'JGraff@qhcus.com', 
                'Barbie_Rebsamen@QuorumHealth.com', 
                'alyssa_harbison@quorumhealth.com', 
                'SarSor@MCKweb.com',
                'NShoemake@qhcus.com',
                'bstroud@qhcus.com']

attachment = MIMEBase('application', 'octet-stream')
msg.attach(MIMEText('Report Attached'))
attachment.set_payload(open(filename, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(filename))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.starttls()
s.login(euser, epass)
s.sendmail('sdudenhofer@qhcus.com', recipients, msg.as_string())
s.quit()
