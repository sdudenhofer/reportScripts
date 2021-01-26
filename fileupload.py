import pysftp
import schedule
import time
import configparser
import os
import logging

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')



logging.basicConfig(filename='/home/itadmin/logs/file_movement.log', level=logging.INFO, format='%(asctime)s %(message)s')

myhost = config['ASFTP']['SERVER']
    #ftp = FTP()
port = config['ASFTP']['PORT']
myusername = config['ASFTP']['USERNAME']
mypassword = config['ASFTP']['PASSWORD']
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None
with pysftp.Connection(host=myhost, username=myusername, password=mypassword, cnopts=cnopts) as sftp:
    print("Connection successfully established To Armus...")
    for (dirpath, dirnames, filenames)in os.walk('/home/itadmin/shared/'):
        for file in filenames:
            localFile = str(dirpath) + str(file)
                # logging.INFO('File was uploaded' + str(localFile))
            remoteFilePath = '/' + file
            sftp.put(localFile, remoteFilePath)
            os.remove(str(dirpath) + str(file))
                # logging.INFO("file removed")


