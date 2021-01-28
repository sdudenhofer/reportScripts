import configparser
import openpyxl
import pyodbc
from termcolor import colored, cprint

pass_test = colored('Passed Query Test', 'green')
fail_test = colored('Failed Query Test', 'red')

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pword = config['MHD']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)

patients = "select * from hospf0062.patients limit 5"
rmbed = "select * from hospf0062.rmbed limit 5"
pathist = "select * from hospf0062.pathist limit 5"
phymast = "select * from hospf0062.phymast limit 5"
admreg = "select * from hospf0062.admreg limit 5"
benefits = "select * from hosp0062.benefits limit 5"
benext = "select * from hospf0062.benext limit 5"
cdnteatb = "select * from hospf0062.cdnteatb limit 5"
oeostat = "select * from orderf0062.oeostat limit 5"
rd = "select * from orderf0062.rd limit 5"
oeorder = "select * from orderf0062.oeorder limit 5"
indaccum = "select * from hospf0062.indaccum limit 5"
dsstat = "select * from hospf0062.dsstat limit 5"
patrmbdp = "select * from hospf0062.patrmbdp limit 5"
ncprm = "select * from orderf0062.ncprm limit 5"
nctrn = "select * from orderf0062.nctrn limit 5"
chpdtaph = "select * from hospf0062.chpdtaph limit 5"
rh = "select * from orderf0062.rh limit 5"
rrfaxlog = "select * from hospf0062.rrfaxlog limit 5"
cdnotetb = "select * from hospf0062.cdnotetb limit 5"
rmtor = "select * from orderf0062.rmtor limit 5"
oeoproc = "select * from orderf0062.oeproc limit 5"
accumchg = "select * from hospf0062.accumchg limit 5"
chrgdesc = "select * from hospf0062.chrgdesc limit 5"
glkeysm = "select * from hospf0062.glkeysm limit 5"
fcdesc = "select * from hospf0062.fcdesc limit 5"
drgdesc = "select * from hospf0062.drgdesc limit 5"
username = "select * from hospf0062.username limit 5"
trdochp = "select * from hospf0062.trdochp limit 5"

cursor = conn.cursor()
data_patients = cursor.execute(patients)
data_rmbed = cursor.execute(rmbed)
data_pathist = cursor.execute(pathist)
data_admreg = cursor.execute(admreg)
data_benefits = cursor.execute(benefits)
data_benext = cursor.execute(benext)
data_cdnteatb = cursor.execute(cdnteatb)
data_oeostat = cursor.execute(oeostat)
data_rd = cursor.execute(rd)
data_oeorder = cursor.execute(oeorder)
data_indaccum = cursor.execute(indaccum)
data_dsstat = cursor.execute(dsstat)
data_patrmbdp = cursor.execute(patrmbdp)
data_ncprm = cursor.execute(ncprm)
data_nctrn = cursor.execute(nctrn)
data_chpdtaph = cursor.execute(chpdtaph)
data_rh = cursor.execute(rh)
data_rrfaxlog = cursor.execute(rrfaxlog)
data_cdnotetb = cursor.execute(cdnotetb)
data_rmtor = cursor.execute(rmtor)
data_oeoproc = cursor.execute(oeoproc)
data_accumchg = cursor.execute(accumchg)
data_chrgdesc = cursor.execute(chrgdesc)
data_glkeysm = cursor.execute(glkeysm)
data_fcdesc = cursor.execute(fcdesc)
data_drgdesc = cursor.execute(drgdesc)
data_username = cursor.execute(username)
data_trdochp = cursor.execute(trdochp)

def checkQuery(query):
    if query.fetchone()[0] == 1:
        cprint(pass_test)
    else:
        cprint(fail_test)

checkQuery(data_patients)
checkQuery(data_rmbed)
checkQuery(data_pathist)
checkQuery(data_admreg)
checkQuery(data_benefits)
checkQuery(data_benext)
checkQuery(data_cdnteatb)
checkQuery(data_oeostat)
checkQuery(data_rd)
checkQuery(data_oeorder)
checkQuery(data_indaccum)
checkQuery(data_dsstat)
checkQuery(data_patrmbdp)
checkQuery(data_ncprm)
checkQuery(nctrn)
checkQuery(chpdtaph)
checkQuery(rh)
checkQuery(rrfaxlog)
checkQuery(cdnotetb)
checkQuery(rmtor)
checkQuery(oeoproc)
checkQuery(accumchg)
checkQuery(chrgdesc)
checkQuery(glkeysm)
checkQuery(fcdesc)
checkQuery(drgdesc)
checkQuery(username)
checkQuery(trdochp)