import pyodbc
import configparser
import pandas as pd
import time
import pysftp
import datetime
import os


config = configparser.ConfigParser()
config.read('D:\\2-PROD\\config.ini')

# logging.basicConfig(filename="/home/itadmin/logs/brault.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

server = config['MHD']['ODBC']
user = config['MHD']['USER']
password = config['MHD']['PASS']
conn = pyodbc.connect(DSN=server, UID=user, PWD=password)

fserver = config['FTP']['SERVER']
fuser = config['FTP']['USER']
fpass = config['FTP']['PASS']

today = datetime.datetime.now()
d = datetime.timedelta(days=5)
c = datetime.timedelta(days=1)
a = today - d
b = today - c
demo_date = a.strftime("%Y-%m-%d")
census_date = b.strftime("%Y-%m-%d")
demoFileDate = today.strftime("%Y-%m-%d")
query_demo = "select t01.patno, t01.hstnum, t01.isadate, t01.isddate, t01.pname,\
    t01.isdob, t01.sex, t02.race, t02.martl, t02.ssn, t02.padr1, \
    t02.padr2, t02.hcity, t02.zip, t02.harcd, t02.phone, t02.c1nam,\
    t02.c1rel, t02.c1arc, t02.c1phn, t01.nwattphy, t05.phname, \
    t01.nwrefdoc, t05.phname, t01.diagn, t01.nwdocnum, t05.phname,\
    t01.ains1, t01.apln1, t01.ains2, t01.apln2, t01.ains3, t01.apln3,\
    t06.policy, t06.iname, t06.reln, \
    t10.ibinam, t10.ibiadr, \
    t10.ibiad2, t10.ibicty, t10.ibstat, t10.ibarcd, t10.ibphn1, \
    t10.ibphn2, t10.ibizip, \
    t07.policy, t07.iname, t07.reln,\
    t11.ibinam, t11.ibiadr, t11.ibiad2, \
    t11.ibicty, t11.ibstat, t11.ibarcd, t11.ibphn1, t11.ibphn2, t11.ibizip,\
    t09.policy, t09.iname, t09.reln, t08.ibinam, t08.ibiadr, t08.ibiad2, \
    t08.ibicty, t08.ibstat, t08.ibarcd, t08.ibphn1, t08.ibphn2, t08.ibizip\
    FROM hospf0062.patients t01 \
    inner join hospf0062.pathist t02 on t02.histn=t01.hstnum \
    inner join hospf0062.phymast t05 on t01.nwattphy=t05.nwdrnum \
    inner join hospf0062.admreg t12 on t01.patno=t12.patno\
    INNER join hospf0062.benefits t06 on t01.patno=t06.patno and t06.histn=t02.histn and t02.ains1=t06.insco and t01.ains1=t06.insco \
    left join hospf0062.benefits t07 on t01.patno=t07.patno and t07.histn=t02.histn and t02.ains2=t07.insco and t01.ains2=t07.insco \
    left join hospf0062.benefits t09 on t01.patno=t09.patno and t09.histn=t02.histn and t02.ains3=t09.insco \
    inner JOIN hospf0062.benext t10 on t06.patno=t10.patno and t10.seqno=t06.seqno \
    left join hospf0062.benext t11 on t07.patno=t11.patno and t11.seqno=t07.seqno\
    left join hospf0062.benext t08 on t09.patno=t08.patno and t08.seqno=t09.seqno\
    WHERE t12.admhsv='EOP' and t01.isadate='" + demo_date + "'"

data = pd.read_sql(query_demo, conn)
df = pd.DataFrame(data)

df.to_csv('D:\\4-FILES\\egoDemo-' + str(demoFileDate) + '.csv', sep='|', index=False, float_format='%.f')

query_census = "SELECT t01.isadate, t01.iatme, t01.isddate, t01.dtime, t01.pname,\
    t01.age, t01.sex, t02.padr1, t01.dcstat, t01.diagn, t04.phname,\
    t04.nwdrnum, t01.hstnum, t01.patno, t01.isdob FROM hospf0062.patients\
    t01 INNER JOIN hospf0062.pathist t02 on t01.hstnum=t02.histn \
    INNER JOIN hospf0062.admreg t03 on t01.patno=t03.patno and \
    t02.histn=t03.hstnum INNER JOIN hospf0062.phymast t04 \
    ON t04.nwdrnum=t03.nwdrnum WHERE t03.admhsv = \
    'EOP' and t01.isadate = '" + census_date + "'"

data1 = pd.read_sql(query_census, conn)
df1 = pd.DataFrame(data1)
df1.to_csv('D:\\4-FILES\\egoCensus-' + str(demoFileDate) + ".csv", sep='|', index=False, float_format='%.f')

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None


with pysftp.Connection(host=fserver, username=fuser, password=fpass, cnopts=cnopts) as sftp:
    print("Connection successfully established...")

    sftp.put('D:\\4-FILES\\egoCensus-' + str(demoFileDate) + '.csv', '/Home/Mckenzie-Willamette Medical/HMS File Transfer/egoCensus-' + str(demoFileDate) + '.csv')
    sftp.put('D:\\4-FILES\\egoDemo-' + str(demoFileDate) + '.csv', '/Home/Mckenzie-Willamette Medical/HMS File Transfer/egoDemo-' + str(demoFileDate) + '.csv')
