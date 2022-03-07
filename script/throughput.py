import configparser
import pyodbc
import pandas as pd
import datetime
from datetime import timedelta

config = configparser.ConfigParser()
config.read('D://2-PROD//config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pwd = config['MHD']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)

query = """
select t01.patno, t01.isadate, t01.iatme, t01.hssvc from hospf0062.patients t01
where isadate between '2022-2-1' and '2022-02-28' and t01.hssvc not in 
('GIO', 'SOP', 'CAT', 'WCC', 'LAB', 'OS1', 'RE2', 'RAD', 'INF', 'CAR', 'MOP', 'REF')
order by isadate asc, iatme asc
"""

query2 = """
SELECT PATN15, ROOM, ASGDATE, HOUR(ASGTIME) as "ASGTIME" FROM HOSPF0062.patrmbdp 
where asgdate between '2022-2-1' and '2022-02-28'
AND ROOM NOT IN ('VMCU', 'HALL', 'VCDU', 'VCCU', 'VOR', 'VMCU', 'VSCJ', 'VCCU', 'VPCU', 'VSSU', 'VCVP',
'1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17',
'S01' , 'S02', 'S03', 'S04', 'S05', 'S06', 'S07', 'S08', 'S09', 'S10', 'S11', 'S12', 'S13',
'S14', 'S15', 'S16', 'S17', 'S18', 'S19', 'S20', 'S21', 'S22', 'S23', 'S24', 'S25', 'S26', 
'S27', 'S28', 'S29', 'S30', 'S31', 'S32', 'S33', 'S34', 'S35', 'S36', 'S37', 'S38', 'S39',
'S40', 'SER1', 'SER2')
order by asgdate asc, asgtime asc
"""

query3 = """
SELECT ROOM, NURST FROM HOSPF0062.RMBED
"""

data = pd.read_sql(query, conn)
admit_dataframe = pd.DataFrame(data)
ns_data = pd.read_sql(query3, conn)
ns_dataframe = pd.DataFrame(ns_data)
room_data = pd.read_sql(query2, conn)
room_dataframe = pd.DataFrame(room_data)

admit_dataframe['PATNO'] = admit_dataframe['PATNO'].astype('int64')
room_dataframe['PATNO'] = room_dataframe['PATN15'].astype('int64')

output = pd.merge(admit_dataframe, room_dataframe, on='PATNO')
output2 = pd.merge(output, ns_dataframe, on='ROOM')

updated = output2[['PATNO', 'ISADATE', 'IATME', 'HSSVC', 'ROOM', 'ASGDATE', 'ASGTIME', 'NURST']]
updated.drop_duplicates(subset=['PATNO'], keep='last', inplace=True)
writer = pd.ExcelWriter('throughput.xlsx', engine='xlsxwriter')
updated.to_excel(writer, sheet_name='All Data', index=False)
updated[['ASGDATE', 'NURST']].groupby(['ASGDATE', 'NURST']).size().to_excel(writer, sheet_name="Total by Move Date")
updated[['NURST', 'ASGDATE', 'ASGTIME']].groupby(['NURST', 'ASGDATE', 'ASGTIME']).size().to_excel(writer, sheet_name="Total by Nurse Station")
updated['ISADATE'].value_counts(sort=False).to_excel(writer, sheet_name="Total by Admit Date")
writer.save()