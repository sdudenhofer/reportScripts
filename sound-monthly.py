import pandas as pd
import configparser
import pyodbc
import xlsxwriter
import time
import datetime
from datetime import datetime
from datetime import date
from datetime import timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os, sys
import schedule
from ftplib import FTP
from time import sleep
import ftplib


#getting login/passwords from configuration file

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
pword = config['AS400']['PASS']
fserver = config['FTP']['SERVER']
fuser = config['FTP']['USER']
fpass = config['FTP']['PASS']
fpath = config['FTP']['PATH3']

# writer = pd.ExcelWriter('sound-test.xlsx', engine='openpyxl')
conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)

today = datetime.now()
a = timedelta(days=32)
b = today - a
month = b.strftime('%Y-%m-%d')
order_day = b.strftime('%y%m%d')

query0 = "select t01.patno, t01.hstnum, t01.isdob, t01.sex, t01.fincl, t02.fdesc, t01.isadate, \
t01.iatme, t03.arrvd, t01.isddate, \
t01.dtime, t05.phname, t06.phname, t08.odobtx, \
t07.oedate, t07.oetime, t07.oproc, t01.drg, t04.drgdes, t01.diag9, t03.admhsv, t03.dischsv \
from hospf062.patients t01 left outer join \
hospf062.fcdesc t02 on t01.fincl =  t02.fcfincl left outer join \
hospf062.admreg t03 on t01.patno = t03.patno and t01.hstnum = t03.hstnum left outer join \
hospf062.drgdesc t04 on t01.drg = t04.drg left outer join \
hospf062.phymast t05 on t01.nwrefdoc = t05.nwdrnum left outer join \
hospf062.phymast t06 on t01.nwattphy = t06.nwdrnum left outer join \
orderf062.oeorder t07 on t03.patno = t07.opat# left outer join \
hospf062.username t08 on t07.ouser = t08.odobnm \
where (t01.isadate >= '" + month + "' and t01.hssvc not in ('RAD', 'CAT', 'BBN', 'NUR', 'RE2', \
'INF', 'LAB') and t01.nwattphy in ('235', '1628', '1388', '1961', '1544', '1414', \
'1958','1643','1167','1906','922','1866','1953','1939','1875','1935','1968','1952','972', \
'2088','2105','2147','2148','2149','2140','1674','2154','2160','2926','87','2936', \
'2952','3339','4058','4022','1958','2975','1636') and t07.oproc in ('DCLTCH', 'DCHOSPM', 'DCSTGHIC', 'DCAIRF', 'DCHOME', 'DCHOMEHH', \
'DCHOSP', 'DCICF', 'DCIPR', 'DCSNF', 'DCACF', 'DCFAC', \
'DCHWHH', 'DCIHOSP', 'DCOTH')) or (t01.nwattphy in ('235', '1628', '1388', '1961', '1544', '1414', \
'1958','1643','1167','1906','922','1866','1953','1939','1875','1935','1968','1952','972', \
'2088','2105','2147','2148','2149','2140','1674','2154','2160','2926','87','2936', \
'2952','3339','4058','4022','1958','2975','1636') and t07.oedate >= '" + order_day + "' and t01.hssvc not in \
('RAD', 'CAT', 'BBN', 'NUR', 'RE2', \
'INF', 'LAB') and t07.ouser in('ILIAO062', 'JSAGE062', \
'KDHUEBER', 'HGHANDOU', 'SROGER78', 'DOWENS23', 'CWHITSEL', 'PDFOX062', 'KDANIE46', 'YOUNGEL', 'HARRISED') \
and t07.oproc in ('DCLTCH', 'DCHOSPM', 'DCSTGHIC', 'DCAIRF', 'DCHOME', 'DCHOMEHH', \
'DCHOSP', 'DCICF', 'DCIPR', 'DCSNF', 'DCACF', 'DCFAC', \
'DCHWHH', 'DCIHOSP', 'DCOTH'))"



cursor = conn.cursor()
data = cursor.execute(query0)
array = []
status = ['DCLTCH', 'DCHOSPM', 'DCSTGHIC', 'DCAIRF', 'DCHOME', 'DCHOMEHH', 
'DCHOSP', 'DCICF', 'DCIPR', 'DCSNF', 'DCACF', 'DCFAC', 
'DCHWHH', 'DCIHOSP', 'DCOTH', 'CONSULT']


output = open("/home/itadmin/automation/sound-data.csv", "w+")
output.write("Med_Rec_No| Acct_No| Date_Of_Birth| Gender| FinancialClass_code| \
FinancialClass_Def| ED_Arrival_Date| ED_Arrival_Time| ED_Dispo_Date| ED_Dispo_Time| ED_Depart_Time|  \
Admiss_Date| Admiss_Time| Admiss_From_Code| Admiss_From_Definition| Discharge_Date| Discharge_Time| \
Admitting_Phys| Attending_Phys| Discharging_Phys| DC_Order_Date|DC_Order_Time| DC_Dispo_Code| \
DC_Dispo_Def| Discharge_Unit| MSDRG_code| MSDRG_Code_Des| APRDRG_Code| APRDRG_Desc| ICD1| ICD2| \
ICD3| LengthOfStay| PtStatus_Admission| PtStatus_Dicharge| ICU_Stay| Cost_Total| Cost_Direct_Rad| \
Cost_Direct_Pharm| Cost_Direct_Lab \n")
for row in data:
    order = str(row[16])
    adm_time = str(row[7])
    start_time = datetime.strptime(str(row[6]), "%Y-%m-%d")
    end_time = datetime.strptime(str(row[9]), '%Y-%m-%d')
    today = date.today()
    day = datetime.strptime(str(today), "%Y-%m-%d")
    days = end_time - start_time
    days = days.days
    disc_doc = str(row[13]).split("{")[0]
    #date_manip = list(str[14])
    #year = '20'.join(date_manip[0] + date_manip[1])
    #month = ''.join(date_manip[2] + date_manip[3])
    #day = ''.join(date_manip[4] + date_manip[5])
    #new_date = str(year) + "-" + str(month) + "-" +str(day)
    if len(adm_time) == 3:
        adm_time = '0' + adm_time
    elif len(adm_time) == 2:
        adm_time = '00' + adm_time
    else:
        adm_time = adm_time
    #if order in status:
    r = str(row[0]) + "| " + str(row[1]) + "| " + str(row[2]) + "| " + str(row[3]) + "| " + str(row[4]) + "| " + str(row[5]) + \
    "| NULL| NULL| NULL| NULL| NULL| " + str(row[6]) + "| " + str(adm_time) + "|NULL| " \
    + str(row[8]) + "| " + str(row[9]) + "| " + str(row[10]) + "| " + str(row[11]) + "| " + str(row[12]) + "| " + str(disc_doc) + "|" \
        + str(row[14]) + "|" + str(row[15]) + "|NULL| " + str(row[16]) + "|NULL|" + str(row[17]) + "| " + str(row[18]) + "|NULL|NULL| " + \
          str(row[19]) + "|NULL|NULL|" + str(days) + "| " + str(row[20]) + "| " + str(row[21]) + "|NULL|NULL|NULL|NULL|NULL"
    array.append(r)
    output.write(str(r) + "\n")
    #else:
    #    print("Skipped " + str(row[0]))
# print(array)
df = pd.DataFrame(array)
# df.columns = 

#print(df)
df.to_csv('/home/itadmin/automation/files/sound_' + str(today) + '.csv', index=False, sep="|")
output.close() 

connection = fserver
ftp = FTP()
port = 21
username = fuser
password = fpass
soundfile = "/home/itadmin/automation/files/sound_" + str(today) + ".csv"

file = open(soundfile, 'rb')

stor_sound = str("STOR sound_" + str(today) + ".csv")
ftp.connect(connection, port)
ftp.login(username, password)
ftp.cwd('/Home/Mckenzie-Willamette Medical/sound')
ftp.storbinary(stor_sound, file, 1024)
file.close()
print("Sound FILE UPLOADED!")

conn.close()
