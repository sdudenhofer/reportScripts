import pandas as pd
import configparser
import pyodbc
import datetime
#from email.utils import formatdate
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
password = config['MHD']['PASS']
eserv = config['O365']['SERVER']
euser = config['O365']['USER']
epass = config['O365']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, PWD=password)

query_omgER = "SELECT T01.PNAME,T01.ISDOB,T01.ISADATE,T01.IATME,T01.ISDDATE,T05.DCDSC, \
				T01.DIAGN, T02.PHNAME FROM HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
    HOSPF0062.PATHIST T06 ON T01.HSTNUM=T06.HISTN LEFT OUTER JOIN \
    HOSPF0062.PHYMAST T02 ON T06.NWFDOC#=T02.NWDRNUM LEFT OUTER JOIN \
    HOSPF0062.STATHD T03 ON T01.HSSVC=T03.SVCCD LEFT OUTER JOIN \
    HOSPF0062.RMBED T04 ON T01.PATNO=T04.PAT# LEFT OUTER JOIN \
    HOSPF0062.DSSTAT T05 ON T01.DCSTAT=T05.DCUBS WHERE \
				T06.NWFDOC# IN(1735, 521, 278, 1695, 978, 744, 1753, 1530, 10853, 869, 1212, 457, 459, 300, 264, 527, 273, 10854,\
    1595, 189, 1080, 10849, 1688, 1579, 515, 283, 1187, 223, 10850, 952, 10855, 1554, 274, 1713, 409, \
    270, 276, 444, 410, 704, 1651, 974, 279, 339, 9961, 1363, 757, 1746, 404, 1720, 1279, 321, 1360, 286, \
    242, 281, 232, 235, 266, 603, 418, 9971, 1721, 173, 664, 9873, 179, 275, 1166, 10097, 55, 1527, 166, \
    9962, 1652, 288, 1544, 1817) AND \
				HSSVC = 'EOP' AND ISADATE = CURRENT DATE -1 DAYS"

query_omgINP = "SELECT T01.PNAME,T01.ISDOB,T01.ISADATE,T01.IATME,T01.ISDDATE,T05.DCDSC, \
				T01.DIAGN, T02.PHNAME FROM HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
    HOSPF0062.PATHIST T06 ON T01.HSTNUM=T06.HISTN LEFT OUTER JOIN \
    HOSPF0062.PHYMAST T02 ON T06.NWFDOC#=T02.NWDRNUM LEFT OUTER JOIN \
    HOSPF0062.STATHD T03 ON T01.HSSVC=T03.SVCCD LEFT OUTER JOIN \
    HOSPF0062.RMBED T04 ON T01.PATNO=T04.PAT# LEFT OUTER JOIN \
    HOSPF0062.DSSTAT T05 ON T01.DCSTAT=T05.DCUBS \
    WHERE HSSVC IN('SIP', 'MIP', 'OBS', 'NUR', 'PED', 'BBN', 'CCU', 'GYN', 'HSI', 'ICU') \
    AND T04.NURST IN('CDU', 'ICU', 'CVIC', 'SCUJ', 'SSU', 'NUR', 'WHBC', 'PCU', 'CCU', 'MCU') AND \
    T06.NWFDOC# IN \
    (1735, 521, 278, 1695, 978, 744, 1753, 1530, 10853, 869, 1212, 457, 459, 300, 264, 527, 273, 10854, \
    1595, 189, 1080, 10849, 1688, 1579, 515, 283, 1187, 223, 10850, 952, 10855, 1554, 274, 1713, 409, \
    270, 276, 444, 410, 704, 1651, 974, 279, 339, 9961, 1363, 757, 1746, 404, 1720, 1279, 321, 1360, 286, \
    242, 281, 232, 235, 266, 603, 418, 9971, 1721, 173, 664, 9873, 179, 275, 1166, 10097, 55, 1527, 166, \
    9962, 1652, 288, 1544, 1817) \
    OR \
    HSSVC IN('CCU', 'GYN', 'HSI', 'ICU', 'NUR', 'OBS', 'PED', 'BBN', 'SIP', 'MIP') \
    AND ISDDATE>=CURRENT DATE-3 DAYS AND \
    t06.NWFDOC# IN \
    (1735, 521, 278, 1695, 978, 744, 1753, 1530, 10853, 869, 1212, 457, 459, 300, 264, 527, 273, 10854, \
    1595, 189, 1080, 10849, 1688, 1579, 515, 283, 1187, 223, 10850, 952, 10855, 1554, 274, 1713 , 409, \
    270, 276, 444, 410, 704, 1651, 974, 279, 339, 9961, 1363, 757, 1746, 404, 1720, 1279, 321, 1360, 286, \
    242, 281, 232, 235, 266, 603 , 418, 9971, 1721, 173, 664, 9873, 179, 275, 1166, 10097, 55, 1527, 166, \
    9962, 1652, 288, 1544, 1817) \
    ORDER BY PHNAME"


def quick_query(query, connection):
    data = pd.read_sql(query, connection)
    output = pd.DataFrame(data)
    return output


writer = pd.ExcelWriter('D://4-FILES//OMG_daily_doctor_report.xlsx', engine='openpyxl')

data_omgER = quick_query(query_omgER, conn)
data_omgER.to_excel(writer, sheet_name='Patients in ER', index=False)

data_omgINP = quick_query(query_omgINP, conn)
data_omgINP.to_excel(writer, sheet_name='Inpatients', index=False)

writer.save()
f = 'D://4-FILES//OMG_daily_doctor_report.xlsx'

msg = MIMEMultipart()
msg['Subject'] = "Daily Doctor Reports For Yesterday"
recipients = [
              'sdudenhofer@qhcus.com',
              'KHarlan@qhcus.com',
              'SMSmith@qhcus.com',
              'jgracen@qhcus.com'
              ]

msg.attach(MIMEText("Report Attached"))
attachment = MIMEBase('application', 'octet-stream')
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
s.starttls()
s.login(euser, epass)
s.sendmail('sdudenhofer@qhcus.com', recipients, msg.as_string())
s.quit()