import pandas as pd
import xlsxwriter
import pymssql
import configparser
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os

# configure config parser
config = configparser.ConfigParser()
config.read('D:\\2-PROD\\config.ini')
server = config['PHS']['SERVER']
user = config['PHS']['USER']
password = config['PHS']['PASS']
eserv = config['O365']['SERVER']
euser = config['O365']['USER']
epass = config['O365']['PASS']
# connect to database
conn = pymssql.connect(server=server, user=user, password=password, database='asmprod')
cursor = conn.cursor()


today = datetime.datetime.now()
a = datetime.timedelta(days=1)
b = today + a
c = datetime.timedelta(days=1)
tomorrow = b.strftime("%Y-%m-%d")
day = today.strftime('%Y-%m-%d')

# create query
query = "select t01.pat_acct_num, t01.pat_legalname, t01.attending_res_name, \
t01.schedcase_start_datetime, substring(t02.actual_proname, 1, 50) from casemain t01 \
left outer join casepro t02 on t01.casemain_id = t02.casemain_id \
where \
t01.schedcase_start_datetime >= '" + day + "' and \
t01.schedcase_start_datetime <= '" + str(tomorrow) + "' and \
(t01.attending_res_name like 'KEIPER%' \
or t01.attending_res_name like 'ANGELES%'\
or t01.attending_res_name like 'GALLO, C%'\
or t01.attending_res_name like 'MILLER%'\
or t01.attending_res_name like 'KORCEK%'\
or t01.attending_res_name like 'JACKSON, L%'\
or t01.attending_res_name like 'LARSEN%'\
or t01.attending_res_name like 'TUMAN%'\
or t01.attending_res_name like 'MILDREN%'\
or t01.attending_res_name like 'TEDESCO%'\
or t01.attending_res_name like 'STRAUB%'\
or t01.attending_res_name like 'FEDOROV%'\
or t01.attending_res_name like 'BEAR%'\
or t01.attending_res_name like 'SHERMAN%'\
or t01.attending_res_name like 'HUDSON, J%') \
order by t01.schedcase_start_datetime"
data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
writer = pd.ExcelWriter('D:\\4-FILES\\surgery-report' + str(day) + '.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False, header=['Account Number', 
                                        'Patient Name', 
                                        'Physician Name', 
                                        'Scheduled Date', 
                                        'Procedure'])


sheet = writer.sheets['Sheet1']

sheet.set_column('A:A', 10)
sheet.set_column('B:B', 25)
sheet.set_column('C:C', 15)
sheet.set_column('D:D', 17)
sheet.set_column('E:E', 45)

sheet.set_landscape()
sheet.fit_to_pages(1, 1)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "Surgeries for: " + str(day)
recipients = ['setdud@mckweb.com', 
                'JudBre@MCKweb.com', 
                'KevWhe@MCKweb.com', 
                'RenRue@MCKweb.com', 
                'CarWer@MCKweb.com', 
                'StaGre@MCKweb.com', 
                'LisBoy@MCKweb.com', 
                'AdaMar@mckweb.com', 
                'KenTau@MCKweb.com', 
                'AngLap@MCKweb.com', 
                'AnnRee@mckweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = 'D:\\4-FILES\\surgery-report' + str(day) + '.xlsx'

msg.attach(MIMEText("Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.starttls()
s.login(euser, epass)
s.sendmail('sdudenhofer@qhcus.com', recipients, msg.as_string())
s.quit()


