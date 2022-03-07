import configparser
import pyodbc
import pandas as pd
import datetime
from datetime import timedelta
import pymssql
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import smtplib
import os

config = configparser.ConfigParser()
config.read('D://2-PROD//config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pwd = config['MHD']['PASS']
eserv = config['O365']['SERVER']
euser = config['O365']['USER']
epass = config['O365']['PASS']
phsserver = config['PHS']['SERVER']
phsuser = config['PHS']['USER']
phspword = config['PHS']['PASS']
# connect to database
conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)
aconn = pymssql.connect(
    server=phsserver,
    user=phsuser,
    password=phspword,
    database='asmprod')

today = datetime.datetime.now()
a = timedelta(days=32)
b = today - a
c = timedelta(days=1)
d = today - c
month = b.strftime('%Y-%m-%d')
yesterday = d.strftime('%Y-%m-%d')

asm_query = """
DECLARE @Date1 datetime
DECLARE @FirstDay datetime
DECLARE @MonthEnd datetime
SET @Date1 = GETDATE()
SET @FirstDay = DATEADD(DAY, 1, EOMONTH(@Date1, -16))
SET @MonthEnd = EOMONTH(@Date1 - 16)
SELECT cm.pat_acct_num as PATNO, cm.pat_displayname, cm.pat_gender,
convert(varchar,
cm.pat_birth_datetime, 101)as DOB, cp.actual_proname,
cast((cm.actcase_stop_datetime-cm.actcase_start_datetime) as TIME(0))
as 'actual total case time',
convert(varchar, cm.actcase_start_datetime, 101)
as 'procdate', cp.woundclass_id, cio.asaclass_id
FROM casemain cm LEFT OUTER JOIN
casepro cp on cm.casemain_id = cp.casemain_id LEFT OUTER JOIN
caseintraop cio on cm.casemain_id = cio.casemain_id and
cp.casemain_id = cio.casemain_id
WHERE cp.actual_proname LIKE '%ARTHROPLASTY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%COLOSTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%ROBOTIC HYSTERECTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%LAMINECTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%ARTHROSCOPY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%HEMILAMINECTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%OSTEOTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%LAMINECTOMIES%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname LIKE '%COLECTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname like '%HYSTERECTOMY%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
OR cp.actual_proname like '%CORONARY ARTERY BYPASS%'
AND cm.actcase_start_datetime between @FirstDay and @MonthEnd
order by cm.actcase_start_datetime
"""

data_asm = pd.read_sql(asm_query, aconn)
df = pd.DataFrame(data_asm)

cursor = conn.cursor()
for m in df['PATNO']:
    if str(m) == 'None':
        m = 0
    patient_number = int(m)
    medicare_query = "SELECT t01.patno, t01.policy from \
        hospf0062.benefits t01 left outer join \
        hospf0062.benext t04 on t01.patno = t04.patno and t01.seqno = \
        t04.seqno WHERE t01.patno = '" + str(patient_number) + "' \
        and t04.ibinam = 'MEDICARE'"
    med_data = cursor.execute(medicare_query)
    for d in med_data:
        pat_number = str(d[0]).strip()
        pnumber = str(d[1]).strip()
        df.loc[(df['PATNO'] == pat_number), 'POLICY_NUMBER'] = pnumber


cur = conn.cursor()
for h in df['PATNO']:
    if str(h) == 'None':
        h = 0
    patientnumber = int(h)
    height_query = "SELECT t01.patno, t02.dthresp from \
            hospf0062.patients t01 LEFT OUTER JOIN \
            hospf0062.chpdtaph t02 on t01.patno = t02.dthpatno WHERE \
            t01.patno = '" + str(patientnumber) + "' and \
            t02.dthfield = 'Height'"
    height_data = cursor.execute(height_query)
    for e in height_data:
        patnumber = str(e[0]).strip()
        height = str(e[1]).strip()
        df.loc[(df['PATNO'] == patnumber), 'height'] = height


for w in df['PATNO']:
    if str(w) == 'None':
        w = 0
    pat_number = int(w)
    weight_query = "SELECT t01.patno, t02.dthresp from \
        hospf0062.patients t01 LEFT OUTER JOIN \
        hospf0062.chpdtaph t02 on t01.patno = t02.dthpatno  \
        WHERE t01.patno = '" + str(pat_number) + "' \
        and t02.dthfield = 'Weight'"
    weight_data = cursor.execute(weight_query)
    for f in weight_data:
        p_number = str(f[0]).strip()
        weight = str(f[1]).strip()
        df.loc[(df['PATNO'] == p_number), 'weight'] = weight

for d in df['PATNO']:
    if str(d) == 'None':
        d = 0
    patient_num = int(d)
    diabetes_query = "SELECT trpat#, TRRECVAL From \
                    orderf0062.nctrn where trpat# = '" + str(patient_num) \
                    + "' and trprmid ='Q0000005629'"
    diabetes_data = cursor.execute(diabetes_query)
    for g in diabetes_data:
        response = str(g[1]).strip()
        patientnum = str(g[0]).strip()
        if response  == "Denies":
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "N"
        elif response == "Type 1":
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "Y"
        elif response == "Type 2 - Insulin Controlled":
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "Y"
        elif response == "Type 2 - Oral Medication Controlled":
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "Y"
        elif response == "Type 2 - Diet Controlled":
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "Y"
        else:
            df.loc[(df['PATNO'] == patientnum), 'diabetes'] = "N"

for o in df['PATNO']:
    if str(o) == 'None':
        o = 0
    pn = int(o)
    outpatient_query = "SELECT patno, isadate, isddate from hospf0062.patients \
    where patno = " + str(pn)
    outp_data = cursor.execute(outpatient_query)
    for h in outp_data:
        admit_date = str(h[1])
        disc_date = str(h[2])
        pat_num = str(h[0])
        if admit_date ==  disc_date:
            df.loc[(df['PATNO'] == pat_num), 'outpatient'] = "Y"
        else:
            df.loc[(df['PATNO'] == pat_num), 'outpatient'] = "N"

# separate hours and minutes
for data in df['actual total case time']:
    if str(data) == 'None':
        data = '00:00:00'
    hours = str(data).split(':')[0]
    minutes = str(data).split(':')[1]
    df.loc[(df['actual total case time'] == data), 'procdurationhr'] = hours
    df.loc[(df['actual total case time'] == data), 'procdurationmin'] = minutes

# change wound class to correct values

df.loc[(df['woundclass_id'] == 9), 'swclass'] = 'C'
df.loc[(df['woundclass_id'] == 3), 'swclass'] = 'CC'
df.loc[(df['woundclass_id'] == 4), 'swclass'] = 'CO'
df.loc[(df['woundclass_id'] == 5), 'swclass'] = 'D'

# change ASA class for correct values

df.loc[(df['asaclass_id'] == 1), 'asaclass'] = '1'
df.loc[(df['asaclass_id'] == 2), 'asaclass'] = '1E'
df.loc[(df['asaclass_id'] == 3), 'asaclass'] = '2'
df.loc[(df['asaclass_id'] == 4), 'asaclass'] = '2E'
df.loc[(df['asaclass_id'] == 5), 'asaclass'] = '3'
df.loc[(df['asaclass_id'] == 6), 'asaclass'] = '3E'
df.loc[(df['asaclass_id'] == 7), 'asaclass'] = '4'
df.loc[(df['asaclass_id'] == 8), 'asaclass'] = '4E'
df.loc[(df['asaclass_id'] == 9), 'asaclass'] = '5'
df.loc[(df['asaclass_id'] == 10), 'asaclass'] = '5E'

# add missing columns

df['anesthesia'] = 'Y'
df['trauma'] = 'N'
df['scope'] = 'N'
df['emergency'] = 'N'
# df['diabetes'] = ''
df['infection'] = ''
df['closure'] = 'Pri'
df['proccode'] = ''
df['jntreptype'] = ''
df['jntreptot'] = ''
df['jntrephemi'] = ''
df['jntrepres'] = ''
# df['outpatient'] = ''
df.drop(columns=['actual total case time', 'woundclass_id'], inplace=True)
updated_dataframe = df.rename(columns={
                    'PATNO': 'patid',
                    'POLICY_NUMBER': 'medicareid',
                    'pat_displayname': 'patient name',
                    'pat_gender': 'gender',
                    'asaclass_id': 'asa',
                    'DOB': 'dob'
                    })

# update column order

df_changed = updated_dataframe[['patid',
                    'medicareid',
                    'patient name',
                    'gender',
                    'dob',
                    'height',
                    'weight',
                    'actual_proname',
                    'proccode',
                    'jntreptype',
                    'jntreptot',
                    'jntrephemi',
                    'jntrepres',
                    'procdate',
                    'diabetes',
                    'emergency',
                    'trauma',
                    'asaclass',
                    'anesthesia',
                    'procdurationhr',
                    'procdurationmin',
                    'outpatient',
                    'scope',
                    'swclass',
                    'closure',
                    'infection',
                            ]]
# print(df_changed.head)
df_changed.to_csv('D:\\4-FILES\\NHSN_data.csv', index=False)

filename = 'D:\\4-FILES\\NHSN_data.csv'
msg = MIMEMultipart()
msg['Subject'] = 'SSI Denominator Data Report'

recipients = ['sdudenhofer@qhcus.com', 'RGeissler@qhcus.com']
attachment = MIMEBase('application', 'octet-stream')
msg.attach(MIMEText('Report Attached'))
attachment.set_payload(open(filename, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(filename))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.starttls()
s.login(euser, epass)
s.sendmail('sdudenhofer@qhcus.com', recipients, msg.as_string())
s.quit()