import configparser
import pyodbc
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
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

conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)

query = """
select t01.nurst, t01.room, t01.bed, t04.isadate, t01.pat#,
t04.pname, T02.trrecdt, t03.prval, t02.trrecval
 FROM          	HOSPF0062.RMBED T01 LEFT OUTER JOIN
				ORDERF0062.NCTRN T02 ON T01.PAT#=T02.TRPAT# LEFT OUTER JOIN
				ORDERF0062.NCPRM T03 ON T02.TRPRMID=T03.PRID LEFT OUTER JOIN
				HOSPF0062.PATIENTS T04 ON T01.PAT#=T04.PATNO
WHERE 			T01.NURST in ('SCUJ', 'WHBC', 'PCU', 'MCU', 'CCU', 'CDU') AND
				T03.PRSTS = 'A' AND T02.TRPRMID IN ('Q0000000163', 'Q0000004013',
                'Q0000003057',  'Q0000000169', 'Q0000000168', 'Q0000000667', 'Q0000000672')
    			AND	T02.TRRECVAL NOT IN ('Peripheral;Saline Lock', 'Peripheral;Double Lumen',
                'Peripheral', 'Peripheral;Saline Lock;Single Lumen', 'Peripheral;Single Lumen',
                'Peripheral;Saline Lock;Double Lumen', 'Saline Lock')
                AND TRRECDT >= CURRENT DATE - 1 DAY
"""

data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
dataframe.to_excel("D://4-FILES//CentLines.xlsx", index=False)

msg = MIMEMultipart()
msg['Subject'] = "Central Lines Report"
recipients = ['sdudenhofer@qhcus.com',
                'HStone@qhcus.com',
                'MShelton@qhcus.com',
                'SSmith07@qhcus.com',
                'TJJohnson@qhcus.com',
                'TJennings@qhcus.com',
                'KWeeks@qhcus.com']
attachment = MIMEBase('application', 'octet-stream')
f = 'D://4-FILES//CentLines.xlsx'

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
