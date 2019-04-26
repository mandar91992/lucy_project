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
InterFile="createInterimFiles_v0.1.py"
Filename123=SubProcesses+"\\"+InterFile

conn=sqlite3.connect(DatabasePath+"\\"+TableName)
c = conn.cursor()
DBName = "controlSheet"

df=pd.DataFrame()


def call_create_InterimFiles(SID,LID,LOC):
	#os.system(Filename123 SID)
	#print(SID,LID,LOC)
	
	os.environ["SID"]=SID
	os.environ["LID"]=LID
	os.environ["LOC"]=LOC
	subprocess.call("python D:\SPLM\Tacton_Integration\code_version_0.0.2\pollingFiles\createInterimFiles_v0.1.py %SID% %LID% %LOC%",shell=True)

def read_from_db():
	x=c.execute ('SELECT  * FROM  controlSheet WHERE SourceBOM=1')
	for row in c.fetchall():
		x=list(row)
	#	print(x)
	call_create_InterimFiles(x[0],x[1],x[2])
	#print(x)


read_from_db()
c.close()
conn.close()