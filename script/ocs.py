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
date = str(first).split(":")[0] + ":" + str(first).split(":")[1] + ":" + str(first).split(":")[2]

# print(date)
day = today.strftime("%Y%d%m%H%M%S")

conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)
cursor = conn.cursor()
initial_array = []
query = """
select nurst, room, bed, pat# from hospf0062.rmbed 
WHERE nurst not in ('WHBC', 'NUR', 'SSU', 'DIAG', 
'EOP', '*VNS', 'CVPR') and room !=  'CVH ' order by nurst
"""
cursor.execute(query)
query_data = cursor.fetchall()
for rows in query_data:
    nurst = rows[0]
    room = rows[1]
    bed = rows[2]
    pnumber = rows[3]
    if nurst == 'CCU ':
        careGroup = 'Adult ICU 3'
    elif nurst == 'PCU ':
        careGroup = 'Adult PCU'
    elif nurst == 'CDU ':
        careGroup = 'Adult OBS'
    else:
        careGroup = 'Adult Med Surg'
    query_output = str(nurst) + ", " + str(room) + "-" + str(bed) + ", " + str(pnumber) + ", " + str(careGroup)
    initial_array.append(query_output)

dataframe = pd.DataFrame([str(sub2).split(", ") for sub2 in  initial_array])
dataframe.columns = ['NURST', 'ROOM-BED', 'PAT#', 'LevelofCareGroup']
dataframe['TimeStamp'] = date
dataframe['System'] = 'McKenzie Willamette Medical Center'
dataframe['Facility'] = 'McKenzie Willamette Medical Center'
dataframe['FacilityID'] = '2397477'
dataframe['BedType'] = 'Regular'
dataframe['isActive'] = 'True'
dataframe['isBlocked'] = 'False'

updated_dataframe = dataframe.rename(columns={
    'NURST': 'Unit',
    })
# print(updated_dataframe['ROOM-BED'])
updated_dataframe['PAT#'] = updated_dataframe['PAT#'].astype(int)
medrecnum = updated_dataframe['PAT#'].to_list()
output = []
covid_output = []
for patient in medrecnum:
    patient = int(patient)
    patient_data = "select t01.patno, t01.hssvc, t01.isadate, time(to_date(digits(cast(t01.iatme as dec(4,0))),'HH24MISS'))  as admit_time, \
                    t01.age, diagn \
                    from hospf0062.patients t01 \
                    where t01.patno = " + str(patient)
    cursor.execute(patient_data)
    pat_data = cursor.fetchall()
    for row in pat_data:
        diagnosis = str(row[5])
        age = str(row[4])
        admit_date = str(row[2])
        admit_time = str(row[3])
        patno = str(row[0])
        service_code = str(row[1])
        if age > '90':
            age = '90+'
        output_data = patno + "," + admit_date + " " + admit_time + "," + age + "," + service_code + "," + diagnosis
        output.append(output_data)
    covid_data = "select rdpt#, rdts, rdrs from orderf0062.rd \
                 where rdpt# = " + str(patient) + " and rdts like 'COV%'"
    cursor.execute(covid_data)
    cov_data = cursor.fetchall()
    for r1 in cov_data:
        if str(r1[2]) == 'NOT DETECTED   ':
            r1[2] = "NEGATIVE"
        out = str(r1[0]) + ", " + str(r1[1]) + ", " + str(r1[2])
        covid_output.append(out)

covid_dataframe = pd.DataFrame([str(sub1).split(", ") for sub1 in covid_output])
covid_dataframe.columns = ['PAT#', 'Test', 'PatientCOVIDStatus']
covid_dataframe.drop_duplicates(subset=['PAT#'], keep='first', inplace=True)
# print(covid_dataframe)
patient_dataframe = pd.DataFrame([str(sub).split(",") for sub in output])
patient_dataframe.columns = ['PAT#', 'patientAdmitDate', 'patientAge', 'ServiceCode', 'Diagnosis', 'holder']
patient_dataframe.drop_duplicates(subset=['PAT#'], keep='first', inplace=True)
patient_dataframe['PatientAdmitDate'] = pd.to_datetime(patient_dataframe['patientAdmitDate']).dt.strftime('%Y-%m-%dT%H:%M:%SZ')
# print(patient_dataframe['PAT#'])
patient_dataframe['PAT#'] = patient_dataframe['PAT#'].astype('int64')
updated_dataframe['PAT#'] = updated_dataframe['PAT#'].astype('int64')
covid_dataframe['PAT#'] = covid_dataframe['PAT#'].astype('int64')
updated_dataframe.replace(0, np.nan, inplace=True)
updated_dataframe.drop_duplicates(subset=['ROOM-BED'], keep='first', inplace=True)
updated_dataframe['Room'] = updated_dataframe['ROOM-BED'].str.split("-", expand=True)[0]
updated_dataframe['Bed'] = updated_dataframe['ROOM-BED'].str.split("-", expand=True)[1]
updated_dataframe['isOccupied'] = np.where(updated_dataframe['PAT#'] > 0, True, False)
patient_dataframe['isPatientLevelCareICU'] = np.where(patient_dataframe['ServiceCode'] == 'ICU', True, False)


output1 = updated_dataframe.merge(patient_dataframe, how='outer', on='PAT#')
output3 = output1.merge(covid_dataframe, how='outer', on='PAT#')
# output3.drop_duplicates(subset=['PAT#'], keep='first', inplace=True)
output2 = output3[[
    'TimeStamp', 'System', 'Facility', 'FacilityID', 'Unit', 'Room', 'Bed', 'BedType',
    'LevelofCareGroup', 'isActive', 'isOccupied', 'isBlocked', 'PatientCOVIDStatus',
    'isPatientLevelCareICU', 'PatientAdmitDate', 'patientAge'
]]

output2.to_csv('D:\\4-FILES\\NCR_McKenzieWillamette_Master_Bed_Data_Extract_' + day + '.csv', index=False)

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

with pysftp.Connection(host=fserver, username=fuser, password=fpass, cnopts=cnopts) as sftp:
    sftp.put('D:\\4-FILES\\NCR_McKenzieWillamette_Master_Bed_Data_Extract_' + day + '.csv', '/Home/Mckenzie-Willamette Medical/state/NCR_McKenzieWillamette_Master_Bed_Data_Extract_' + day + '.csv')