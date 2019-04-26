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

def create_table():
	try:
		c.execute('CREATE TABLE IF NOT EXISTS controlSheet (SolutionID TEXT, LineID TEXT, Path TEXT, SourceBOM INT, NewItemCreate INT, QueryOutput INT, InterimLoadFIle INT, FinalLoadFile INT, DataSets INT, TopNode TEXT, NX_1 TEXT, NX_2 TEXT)')
	except:
		print('there is a isuue creating the table')

DatabasePath="D:\\SPLM\\Tacton_Integration\\ToProcess\\Database"
TableName="ControlTable.db"



SubProcesses="D:\\SPLM\\Tacton_Integration\\code_version_0.0.2\\pollingFiles"
InterFile="createInterimFiles_v0.1.py"
Filename123=SubProcesses+"\\"+InterFile
try:
	conn=sqlite3.connect(DatabasePath+"\\"+TableName)
	c = conn.cursor()
	DBName = "controlSheet"
except:
	print('THERE IS AN ISSUE CONNECTING TO THE DATABASE')

df=pd.DataFrame()


def call_create_InterimFiles(SID,LID,LOC):
	#os.system(Filename123 SID)
	#print(SID,LID,LOC)
	
	os.environ["SID"]=SID
	os.environ["LID"]=LID
	os.environ["LOC"]=LOC
	try:
		subprocess.call("python D:\SPLM\Tacton_Integration\code_version_0.0.2\pollingFiles\createInterimFiles_v0.1.py %SID% %LID% %LOC%",shell=True)
	except:
		print('SUBPROCESS COULD NOT BE CALLED')
def read_from_db():
	try:
		x=c.execute ('SELECT  * FROM  controlSheet WHERE SourceBOM=1')
		for row in c.fetchall():
			x=list(row)
		#	print(x)
		call_create_InterimFiles(x[0],x[1],x[2])
	#print(x)
	except:
		create_table();
		print('There Is an issue creating the table')


read_from_db()
c.close()
conn.close()