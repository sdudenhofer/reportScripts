from numpy.core.numeric import full_like
import pyodbc
import pandas as pd
import datetime
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from time import sleep
import smtplib
import os
import configparser

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')

server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

today = datetime.now()
day = today.strftime('%Y-%m-%d')

try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

query_total_flu = "select t01.oproc, t05.hssvc, t02.patno, t02.pname, t05.isadate, \
t03.room, t03.bed, t03.nurst, t01.osdate, \
    t04.stdsc, t01.oisrdate, t06.rdrs from \
    orderf0062.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
    left outer join hospf0062.rmbed t03 on t01.opat# = t03.pat# and t02.patno = t03.pat# \
    left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
    left outer join hospf0062.patients t05 on t01.opat# = t05.patno \
    left outer join orderf0062.rd t06 on t01.opat# = t06.rdpt# and t01.oord# = t06.rdor# \
    where oproc = 'INFLUPCR' and t06.rdrs != ' ' and t06.rdrs = 'POSITIVE' \
    and t05.hssvc != 'EOP' and stdsc != 'CANCELLED' and t03.nurst != 'NULL' \
    ORDER BY t05.isadate"

query_flu_admit = "select t01.oproc, t05.hssvc, t02.patno, t02.pname, t05.isadate, \
t03.room, t03.bed, t03.nurst, t01.osdate, \
    t04.stdsc, t01.oisrdate, t06.rdrs from \
    orderf0062.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
    left outer join hospf0062.rmbed t03 on t01.opat# = t03.pat# and t02.patno = t03.pat# \
    left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
    left outer join hospf0062.patients t05 on t01.opat# = t05.patno \
    left outer join orderf0062.rd t06 on t01.opat# = t06.rdpt# and t01.oord# = t06.rdor# \
    where oproc = 'INFLUPCR' and t06.rdrs != ' ' and t06.rdrs = 'POSITIVE' \
    and t05.hssvc != 'EOP' and stdsc != 'CANCELLED' and t03.nurst != 'NULL' \
    and t05.isadate = CURRENT DATE-1 DAYS \
    ORDER BY t05.isadate"

query_flu_icu = "select t01.oproc, t05.hssvc, t02.patno, t02.pname, t05.isadate, \
t03.room, t03.bed, t03.nurst, t01.osdate, \
    t04.stdsc, t01.oisrdate, t06.rdrs from \
    orderf0062.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
    left outer join hospf0062.rmbed t03 on t01.opat# = t03.pat# and t02.patno = t03.pat# \
    left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
    left outer join hospf0062.patients t05 on t01.opat# = t05.patno \
    left outer join orderf0062.rd t06 on t01.opat# = t06.rdpt# and t01.oord# = t06.rdor# \
    where oproc = 'INFLUPCR' and t06.rdrs != ' '  and t06.rdrs = 'POSITIVE' and t03.nurst = 'CCU' \
    and  t01.osdate >= '201022' and stdsc != 'CANCELLED' \
    ORDER BY t01.osdate"

query_covid = "select t01.oproc, t05.hssvc, t02.patno, t02.pname, t03.room, \
    t03.bed, t03.nurst, t01.osdate, \
    t04.stdsc, t01.oisrdate, t06.rdrs from \
    orderf0062.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
    left outer join hospf0062.rmbed t03 on t01.opat# = t03.pat# and t02.patno = t03.pat# \
    left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
    left outer join hospf0062.patients t05 on t01.opat# = t05.patno \
    left outer join orderf0062.rd t06 on t01.opat# = t06.rdpt# and t01.oord# = t06.rdor# \
    where oproc in ('COVID-19', 'COVID-LC', 'COVID-QL', 'COVID-CM', 'COVID-RO') \
    and t02.patno != '4106640' and t02.patno != '4106643' \
    and t03.nurst != '' and t06.rdrs like 'POS SARS%' \
    ORDER BY t05.isadate"

query_flu_deaths = "SELECT T01.HSSVC, T01.PATNO, T01.PNAME, T01.AGE, T01.SEX, \
                T01.ISADATE, T01.IATME, T01.DIAGN, \
                T01.ISDDATE, T01.DTIME, T03.DCSDS, T03.DCEXP, t04.room, t04.bed, t05.nurst \
      FROM      HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
                HOSPF0062.DSSTAT T03 \
      ON        T01.DCSTAT = T03.DCUBS \
        LEFT OUTER JOIN hospf0062.patrmbdp t04 ON t01.patno = t04.patn15 \
            LEFT OUTER JOIN hospf0062.rmbed t05 on t01.patno = t05.pat# \
            left outer join orderf0062.oeorder t07 on t01.patno = t07.opat# \
            left outer join orderf0062.rd t06 on t01.patno = t06.rdpt# and t06.rdor# = t07.oord# \
      WHERE     T01.HSSVC IN ('SIP', 'MIP', 'ICU', 'GYN', 'CCU', 'PED', 'OBS', 'OBI', 'NUR') \
        AND     ISDDATE = CURRENT DATE - 1 DAYS AND T03.DCEXP = 'Y' and t07.oproc = 'INFLUPCR' \
        and t06.rdrs = 'POSITIVE' \
      ORDER BY  T01.PNAME ASC"

query_covid_death = "SELECT T01.HSSVC, T01.PATNO, T01.PNAME, T01.AGE, T01.SEX, \
                T01.ISADATE, T01.IATME, T01.DIAGN, \
                T01.ISDDATE, T01.DTIME, T03.DCSDS, T03.DCEXP, t04.room, t04.bed, t05.nurst \
      FROM      HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
                HOSPF0062.DSSTAT T03 \
      ON        T01.DCSTAT = T03.DCUBS \
        LEFT OUTER JOIN hospf0062.patrmbdp t04 ON t01.patno = t04.patn15 \
            LEFT OUTER JOIN hospf0062.rmbed t05 on t01.patno = t05.pat# \
            left outer join orderf0062.rd t06 on t01.patno = t06.rdpt# \
            left outer join orderf0062.oeorder t07 on t01.patno = t07.opat# \
      WHERE     T01.HSSVC IN ('SIP', 'MIP', 'ICU', 'GYN', 'CCU', 'PED', 'OBS', 'OBI', 'NUR') \
        AND     ISDDATE = CURRENT DATE - 1 DAYS AND T03.DCEXP = 'Y' and t07.oproc in \
        ('COVID-19', 'COVID-LC', 'COVID-QL', 'COVID-CM', 'COVID-RO') \
        and t06.rdrs like 'POS SARS%' \
      ORDER BY  T01.PNAME ASC"

data = pd.read_sql(query_total_flu, conn)
df = pd.DataFrame(data)
ccu_count = df.loc[df['NURST'] == 'CCU'].count()
conn.close()


try:
    conn2 = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

data1 = pd.read_sql(query_flu_admit, conn2)
df1 = pd.DataFrame(data1)
flu_t_index= df.index
flu_a_index = df1.index
flu_total = len(flu_t_index)
flu_admit = len(flu_a_index)
conn2.close()

try:
    conn3 = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

data2 = pd.read_sql(query_flu_icu, conn3)
df2 = pd.DataFrame(data2)
flu_icu = df2.index
icu_total = len(flu_icu)
conn3.close()

try:
    conn4 = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

data3 = pd.read_sql(query_covid, conn4)
df3 = pd.DataFrame(data3)
covid_ih = df3.index
covid_total = len(covid_ih)
totalfc = covid_total + flu_total
conn4.close()

try:
    conn5 = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

data4 = pd.read_sql(query_flu_deaths, conn5)
df4 = pd.DataFrame(data4)
fd_count = df4.index
total_fd = len(fd_count)
conn5.close()

try:
    conn6 = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    print(e)
    sleep(60)

data5 = pd.read_sql(query_covid_death, conn6)
df5 = pd.DataFrame(data5)
out = df5.drop_duplicates(subset=['PATNO'], keep='first')
cd_count = out.index
total_cd = len(cd_count)
cfd_total = total_cd + total_fd

msg = MIMEMultipart()
msg['Subject'] = "HHS Report For: " + str(day)
recipients = ['setdud@mckweb.com']


msg.attach(MIMEText("Total Flu Positives: " + str(flu_total) +"\n" \
"Flu Admits Yesterday: " + str(flu_admit) +  "\n"\
"ICU Patients with Flu: " + str(icu_total) +  "\n"\
"Total Covid and Flu Patients: " + str(totalfc) +  "\n"\
"Deaths Yesterday with Positive Flu Test: " + str(total_fd) + "\n"\
"Deaths Yesterday with Positive Flu or Covid Test: " + str(cfd_total)))
#attachment.set_payload(open(f, 'rb').read())
#encoders.encode_base64(attachment)
#attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
#msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()

