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
	
def dynamic_data_entry(SID,LID,LOC,SBOM,NItem,Query,iLoad,fLoad,DS,tNode,nX1,nX2):
	try:
		c.execute("INSERT INTO controlSheet (SolutionID,LineID,Path,SourceBOM, NewItemCreate,QueryOutput,InterimLoadFIle,FinalLoadFile,DataSets,TopNode,NX_1,NX_2) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",(SID,LID,LOC,SBOM,NItem,Query,iLoad,fLoad,DS,tNode,nX1,nX2))
		conn.commit()
	except:
		print('Creating the table')
		create_table()

def remove_duplicate_lines():
	try:
		c.execute("DELETE FROM controlSheet WHERE rowid NOT IN (SELECT min(rowid) FROM controlSheet GROUP BY SolutionID, LineID)");
		conn.commit()
	except:
		print('There was a issue DELETING DUPLICATE LINES')
		create_table()


def read_from_db():
	try:
		c.execute ('SELECT  * FROM  controlSheet WHERE SourceBOM=1')
		data=c.fetchall()
		i=1
		for row in c.fetchall():
			print("ROW: ",i,"..",row)
			i +=1
	except:
		print('There is an issue READING the DATABASE')
		create_table()







DatabasePath="D:\\SPLM\\Tacton_Integration\\ToProcess\\Database"
TableName="ControlTable.db"
try:
	conn=sqlite3.connect(DatabasePath+"\\"+TableName)
	c = conn.cursor()
except:
	print('There was an error connecting to the database')


####################create_table()	#Run the first time to create the table, needs to be commented out for BAU


#### Global Variables Definition 
path = "\\\\ukthmtacl01v\\temp\\solution" 		# This is the path where you want to search for the Solution Package
ProcessingPath = "D:\\SPLM\\Tacton_Integration\\ToProcess"		#Destination of files post copy and rename, copied to the local server




######### This is where the main code starts


# this is the extension you want to detect
extension = '.xls'
i=1
Position=0
FilePath=[]
for root, dirs_list, files_list in os.walk(path):
	for file_name in files_list:
		if os.path.splitext(file_name)[-1] == extension:
			file_name_path=os.path.join(root)
			FilePath.append(os.path.join(root))
			
			
			#This line get all the files in the directory in a list	
			x = [os.path.join(r,file) for r,d,f in os.walk(file_name_path) for file in f]


			#Creating the destination Path where the files need to be copied to 
			
			x1=FilePath[Position].split("\\")[7].split(" ")[0]
			z1=FilePath[Position].split("\\")[8].split(" ")[0]
			destPath=x1+"_"+z1

			CreatePath=ProcessingPath+"\\"+destPath


			#dynamic_data_entry(SolutionID, LineID, Path, SourceBOM, QueryOutput, InterimLoadFIle, FinalLoadFile, DataSets, TopNode)
			if(len(x)>0):
				dynamic_data_entry(x1,z1,CreatePath,1,0,0,0,0,1,"NULL","0","0")
			else:
				dynamic_data_entry(x1,z1,CreatePath,1,0,0,0,0,0,"NULL","0","0")

			

			#Creates the destination Path to move the files to 
			os.mkdir(CreatePath)
			
			#Takes every file, renames it to new name and moves it to destination path
			for j in range(len(x)):

				y1=x[j].split("\\")[-1]
				y2=CreatePath.split("\\")[-1]
				UpdateFileName=y2+"_"+y1

				#Files have been renames to new file name
				#TRY CATCH SHOULD BE USED HERE BUT I AM NOT GETTING THE CODE FOR CREATING FILE
				os.rename(x[j],FilePath[Position]+"\\"+UpdateFileName)
				
				#The new file name is now moved to the new location
				shutil.move(FilePath[Position]+"\\"+UpdateFileName,CreatePath+"\\"+UpdateFileName)
			i +=1
			Position +=1

remove_duplicate_lines()

########read_from_db() #This function is for troubleshooting purposes to see the code


c.close()
conn.close()
