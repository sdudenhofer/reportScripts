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
from time import sleep

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']



try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Eroor as e:
    print(e)
    sleep(60)

today = datetime.datetime.now()
c = datetime.timedelta(days=1)
d = today - c
yesterday = d.strftime("%y%m%d")

query = "SELECT t07.rmxrf as ACC, t02.stdsc as STATUS, t06.povdsc as TEST, t01.orflag as P, t03.rhiscldt, \
t03.rhcltm, t03.rhisrcdt, t03.rhrctm, t04.pname as NAME, t04.patno, t05.nurst, t05.room, t05.bed \
FROM orderf062.oeorder t01 LEFT OUTER JOIN orderf062.oeostat t02 on t01.ostat = t02.stat# \
LEFT OUTER JOIN orderf062.rh t03 on t01.opat# = t03.rhpt# and t01.oord# = t03.rhor# \
LEFT OUTER JOIN hospf062.patients t04 on t01.opat# = t04.patno \
LEFT OUTER JOIN hospf062.rmbed t05 on t01.opat# = t05.pat# and t04.patno = t05.pat# \
LEFT OUTER JOIN orderf062.rmtor t07 on t01.opat# = t07.rmpt# and t01.oord# = t07.rmor# \
LEFT OUTER JOIN orderf062.oeproc t06 on t01.oproc = t06.pproc \
WHERE t01.osdate >= '" + yesterday + "' and t01.otodpt = 'LAB' \
and t01.orflag != '1' order by t05.nurst"

data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
dataframe.to_csv('/home/itadmin/olmonitor/olmonitor.txt',  sep='|', index=False, header=False, float_format='%.f')
