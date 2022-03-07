 #import pyodbc
import configparser
import pandas as pd
import xlsxwriter
import pymssql
import datetime
import openpyxl
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
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']
# db = config['PHS']['DATABASE']

day = datetime.datetime.now()
today = day.strftime('%Y-%m-%d')
a = datetime.timedelta(days=35)
b = day + a
month = b.strftime('%Y-%m-%d')
start_date = today + " 00:00:000"
end_date = month + " 23:59:000"

query = "select convert(date, ptb.start_datetime) as StartDate, \
pt.name_legal,  pml.mrn, pb.name as SurgeryName, r.name as Physician \
from patbooking ptb left outer join pat pt on ptb.pat_id = pt.pat_id \
left outer join patmrnlist pml on pt.pat_id = pml.pat_id and ptb.pat_id = pml.pat_id \
left outer join visitapptlist val on val.appt_id = ptb.appt_id \
left outer join visit v on val.visit_id = v.visit_id and v.pat_id = pt.pat_id \
left outer join appt a on ptb.appt_id = a.appt_id and val.appt_id = a.appt_id \
left outer join probooking pb on ptb.appt_id = pb.appt_id \
left outer join resbooking rb on ptb.appt_id = rb.appt_id \
left outer join res r on ptb.ordering_phys_id = r.res_id \
left outer join service s on a.service_id = s.service_id \
left outer join apptstatus ats on a.apptstatus_id = ats.apptstatus_id \
left outer join apptclass atc on ptb.apptclass_id = atc.apptclass_id \
where ptb.start_datetime between '" + start_date + "' and '" + end_date + "' \
and ats.apptstatus_id != 2 \
and (r.name like 'KEIPER%' or r.name like 'ANGELES%' or r.name like 'GALLO, C%' \
or r.name like 'MILLER%' or r.name like 'KORCEK%' or r.name like 'JACKSON, L%' \
or r.name like 'LARSEN%' or r.name like 'TUMAN%' or r.name like 'MILDREN%' or \
r.name like 'TEDESCO%' or r.name like 'STRAUB%' or r.name like 'FEDOROV, A%' or \
r.name like 'BEAR%' or r.name like 'SHERMAN%' or r.name like 'HUDSON%') \
order by StartDate"

query2 = "select convert(date, ptb.start_datetime) as StartDate, \
pt.name_legal,  pml.mrn, pb.name as SurgeryName, r.name as Physician \
from patbooking ptb left outer join pat pt on ptb.pat_id = pt.pat_id \
left outer join patmrnlist pml on pt.pat_id = pml.pat_id and ptb.pat_id = pml.pat_id \
left outer join visitapptlist val on val.appt_id = ptb.appt_id \
left outer join visit v on val.visit_id = v.visit_id and v.pat_id = pt.pat_id \
left outer join appt a on ptb.appt_id = a.appt_id and val.appt_id = a.appt_id \
left outer join probooking pb on ptb.appt_id = pb.appt_id \
left outer join resbooking rb on ptb.appt_id = rb.appt_id \
left outer join res r on ptb.ordering_phys_id = r.res_id \
left outer join service s on a.service_id = s.service_id \
left outer join apptstatus ats on a.apptstatus_id = ats.apptstatus_id \
left outer join apptclass atc on ptb.apptclass_id = atc.apptclass_id \
where ptb.start_datetime between '" + start_date + "' and '" + end_date + "' \
and ats.apptstatus_id != 2 \
and (r.name like 'MILLER%' or r.name like 'ANGELES%' or r.name like 'GALLO, C%' \
or r.name like 'SHERMAN%' or r.name like 'KEIPER%') \
order by StartDate"

conn = pymssql.connect(server=server, user=user, password=password, database='phsprod')

data = pd.read_sql(query, conn)
data2 = pd.read_sql(query2, conn)
df = pd.DataFrame(data)
df1 = pd.DataFrame(data2)
count = df.shape[0]
out = df.drop_duplicates(subset = ['mrn'], keep='last')
out1 = df1.drop_duplicates(subset = ['mrn'], keep='last')
writer = pd.ExcelWriter('D:\\4-FILES\\pt-schedule-report-' + today + '.xlsx', engine='xlsxwriter')
out.to_excel(writer, index=False, sheet_name='Schedule')
out1.to_excel(writer, index=False, sheet_name='Neuro')
df['StartDate'].value_counts(sort=False).to_excel(writer, sheet_name="Total by Date")
df['Physician'].value_counts(sort=True).to_excel(writer, sheet_name="Total By Doctor")
worksheet = writer.sheets['Schedule']
worksheet1 = writer.sheets['Neuro']

worksheet.set_landscape()
worksheet1.set_landscape()
writer.save()
conn.close()

filename = 'D:\\4-FILES\\pt-schedule-report-' + str(today) + '.xlsx'
msg = MIMEMultipart()
msg['Subject'] = 'Daily PT Surgery Schedule Report'
recipients = ['sdudenhofer@qhcus.com', 'CWerthRudolph@qhcus.com']

attachment = MIMEBase('application','octet-stream')
    
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
    
