import configparser
import pandas as pd
import pymssql
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
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
c = datetime.timedelta(days=32)
e = b - c
d = e.strftime('%Y-%m-%d')
mtd_start_date = d + " 00:00:000"

query = """
DECLARE @Date1 datetime
DECLARE @FirstDay datetime
DECLARE @MonthEnd datetime
SET @Date1 = GETDATE()
SET @FirstDay = DATEADD(DAY, 1, EOMONTH(@Date1, -2))
SET @MonthEnd = DATEADD(DAY, 1, @Date1 - 2)
select cm.casemain_id, cm.pat_acct_num, cm.pat_lastname, cp.primpract_res_name, 
cio.pat_or_in_datetime, cm.actcase_start_datetime, cm.actcase_stop_datetime, 
cm.schedcase_start_datetime, cm.schedcase_stop_datetime, 
cio.pat_or_out_datetime, cm.actualroom_res_abbr, cp.actual_proname as 'Procedure' 
from casemain cm left outer join psmresindroom pri on 
cm.actualroom_res_id = pri.res_id 
left outer join casepreop cpo on cm.casemain_id = cpo.casemain_id 
left outer join psmpattype pt on cpo.pattype_id = pt.pattype_id 
left outer join casepro cp on cm.casemain_id = cp.casemain_id 
left outer join caseintraop cio on cm.casemain_id = cio.casemain_id 
where cm.actcase_start_datetime between @FirstDay and @MonthEnd
order by cm.actcase_start_datetime
"""

data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
out = dataframe.drop_duplicates(subset=['pat_acct_num'], keep='first')
out = out.drop_duplicates(subset=['pat_acct_num'], keep='first')
out.loc[(out['actualroom_res_abbr'] == 'OR 03'), 'Room Name'] = "OPERATING ROOM 3"
out.loc[(out['actualroom_res_abbr'] == 'OR 02'), 'Room Name'] = "OPERATING ROOM 2"
out.loc[(out['actualroom_res_abbr'] == 'OR 01'), 'Room Name'] = "OPERATING ROOM 1"
out.loc[(out['actualroom_res_abbr'] == 'OR 04'), 'Room Name'] = "OPERATING ROOM 4"
out.loc[(out['actualroom_res_abbr'] == 'OR 05'), 'Room Name'] = "OPERATING ROOM 5"
out.loc[(out['actualroom_res_abbr'] == 'OR 06'), 'Room Name'] = "OPERATING ROOM 6"
out.loc[(out['actualroom_res_abbr'] == 'OR 07'), 'Room Name'] = "OPERATING ROOM 7"
out.loc[(out['actualroom_res_abbr'] == 'OR 08'), 'Room Name'] = "OPERATING ROOM 8"
out.loc[(out['actualroom_res_abbr'] == 'OR 09'), 'Room Name'] = "OPERATING ROOM 9"
out.loc[(out['actualroom_res_abbr'] == 'OR 10'), 'Room Name'] = "OPERATING ROOM 10"
out.loc[(out['actualroom_res_abbr'] == 'OR 11'), 'Room Name'] = "OPERATING ROOM 11"
out.loc[(out['actualroom_res_abbr'] == 'OR 12'), 'Room Name'] = "OPERATING ROOM 12"
out.drop(['actualroom_res_abbr'], axis=1, inplace=True)

df_changed = out[['casemain_id',
                    'pat_acct_num',
                    'pat_lastname',
                    'primpract_res_name',
                    'pat_or_in_datetime',
                    'actcase_start_datetime',
                    'actcase_stop_datetime',
                    'pat_or_out_datetime',
                    'schedcase_start_datetime',
                    'schedcase_stop_datetime',
                    'Room Name',
                    'Procedure']]

writer = pd.ExcelWriter('D:\\4-FILES\\MonthlyOR-log-' + today + '.xlsx', engine='xlsxwriter')
df_changed.to_excel(writer, index=False, sheet_name='OR Data')
writer.save()
conn.close()

filename = 'D:\\4-FILES\\MonthlyOR-log-' + today + '.xlsx'
msg = MIMEMultipart()
msg['Subject'] = 'Monthly OR log Report'

recipients = ['sdudenhofer@qhcus.com', 'BFreybates@qhcus.com', 'jgraff@qhcus.com']
attachment = MIMEBase('application', 'octet-stream')
msg.attach(MIMEText('Report Attached'))
attachment.set_payload(open(filename, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(filename))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.starttls()
s.login(euser, epass)
#s.sendmail('sdudenhofer@qhcus.com', recipients, msg.as_string())
#s.quit()