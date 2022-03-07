import numpy as np
import pyodbc
import pandas as pd
import datetime
import pysftp
import configparser

config = configparser.ConfigParser()
config.read('D:\\2-PROD\\config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pword = config['MHD']['PASS']
fserver = config['FTP']['SERVER']
fuser = config['FTP']['USER']
fpass = config['FTP']['PASS']
fpath = config['FTP']['PATH']

today = datetime.datetime.now()
today1 = today.isoformat()
first = str(today1).split(".")[0] + "Z"
second = str(today1).split(".")[1]
date = str(first).split(":")[0] + ":" + str(first).split(":")[1] + ":" + str(first).split(":")[2]
day = today.strftime("%Y%d%m%H%M%S")

query = """
    select t01.patno, t01.isadate, time(to_date(digits(cast(t01.iatme as dec(4,0))),'HH24MISS'))  as admit_time, 
    t01.isddate, t01.age, t02.nurst, t02.room, t02.bed from hospf0062.patients t01 left outer join hospf0062.rmbed t02 on t01.patno = t02.pat# 
    where hssvc = 'EOP' and isadate = CURRENT DATE and isddate != CURRENT DATE OR
    t02.nurst =  'EOP'
"""

conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)
cursor = conn.cursor()

cursor.execute(query)
ed_data = cursor.fetchall()
patient_number = []
df = []
for row in ed_data:
    patno = str(row[0])
    admit_date = str(row[1])
    admit_time = str(row[2])
    discharge_date = str(row[3])
    nurse_station = str(row[5])
    age = str(row[4])
    if nurse_station == "":
        isPatientBoarding = "False"
    else:
        isPatientBoarding = "True"
    query1 = "select t01.patno, t02.rdts, t02.rdrs, t01.hssvc, t01.isadate, t01.age, diagn \
            from hospf0062.patients t01 left outer join orderf0062.rd  t02 on  t01.patno = t02.rdpt# \
            where t01.patno = " + patno + " and t02.rdts like  '%COV%'"
    cursor.execute(query1)
    covid_test = cursor.fetchall()
    for row1 in covid_test:
        pat_no = str(row1[0])
        result = str(row1[2])
        if result == 'NOT DETECTED   ':
            result = "Negative"
        patient_number.append(patno)
        eop_output = patno + "," + admit_date + "T" + admit_time + "," + \
            isPatientBoarding + "," + result + "," + age
        df.append(eop_output)

dataframe = pd.DataFrame([str(sub2).split(",") for sub2 in df])
dataframe.columns = ['PATNO', 'RegistrationTimeStamp', 'isPatientBoarding', 'PatientCOVIDStatus', 'patientAge']
dataframe.drop_duplicates(subset=['PATNO'], keep='first')
dataframe['TimeStamp'] = date
dataframe['System'] = 'McKenzie Willamette Medical Center'
dataframe['Facility'] = 'McKenzie Willamette Medical Center'
dataframe['FacilityID'] = '2397477'

output = dataframe[['TimeStamp', 'RegistrationTimeStamp', 'System', 'Facility', 'FacilityID',
                    'isPatientBoarding', 'PatientCOVIDStatus', 'patientAge']]
output.to_csv('D:\\4-FILES\\NCR_McKenzieWillametteMedical_ED_Census_File_' + day + '.csv', index=False)

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

with pysftp.Connection(host=fserver, username=fuser, password=fpass, cnopts=cnopts) as sftp:
    sftp.put('D:\\4-FILES\\NCR_McKenzieWillametteMedical_ED_Census_File_' + day + '.csv', '/Home/Mckenzie-Willamette Medical/state/NCR_McKenzieWillametteMedical_ED_Census_File_' + day + '.csv')