import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import csv
import string
import os
import subprocess
import shutil
import sqlite3


DatabasePath="D:\\SPLM\\Tacton_Integration\\ToProcess\\Database"
TableName="ControlTable.db"

SubProcesses="D:\\SPLM\\Tacton_Integration\\code_version_0.0.2\\pollingFiles"
InterFile="createLoadFile_v0.1.py"
#Filename123=SubProcesses+"\\"+InterFile

conn=sqlite3.connect(DatabasePath+"\\"+TableName)
c = conn.cursor()
#DBName = "controlSheet"

#df=pd.DataFrame()


def call_create_FinalLoadFile(SID,LID,LOC):

	os.environ["SID_F"]=SID
	os.environ["LID_F"]=LID
	os.environ["LOC_F"]=LOC
	subprocess.call("python D:\SPLM\Tacton_Integration\code_version_0.0.2\pollingFiles\createLoadFile_v0.1.py %SID_F% %LID_F% %LOC_F%",shell=True)

def read_from_db():
	count =0
	x=c.execute ('SELECT  * FROM  controlSheet WHERE QueryOutput=1 AND InterimLoadFIle = 1') #AND NewItemCreate=1)
	for row in c.fetchall():
		x=list(row)
		count +=1
		#print(x)
		
	
	if count == 0:
		#print("There is nothing in the database")
		pass
	else:
		call_create_FinalLoadFile(x[0],x[1],x[2])
	
	#print(x)


read_from_db()
c.close()
conn.close()