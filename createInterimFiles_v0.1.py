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
import sys

arg1=str(sys.argv[1])
arg2=str(sys.argv[2])
arg3=str(sys.argv[3])

#print("I am in createInterimFile")
#print("The number of arguments are:",len(sys.argv))
#print("The arguments are:",str(sys.argv))

#print(arg1,arg2,arg3)

SOURCE_FILE_PATH=arg3
FileName=arg1+"_"+arg2+"_"+"BOM Report.xls"
print("xls file is not present")
FilePath=SOURCE_FILE_PATH+"\\"+FileName


InterimLoadFile=arg1+"_"+arg2+"_"+"InterimLoadFile.xls"
AutoAssignPath=InterimLoadFilePath=SOURCE_FILE_PATH
AutoAssignFileName=arg1+"_"+arg2+"_"+"AutoAssign.pim"
AutoAssign1=AutoAssignPath+"\\"+AutoAssignFileName

QueryLoadFile=arg1+"_"+arg2+"_"+"QueryLoadFile.xml"
QueryLoadFilePath=SOURCE_FILE_PATH

InterimDataFrame=pd.DataFrame()


TemplateFilePath="D:\\SPLM\\Tacton_Integration\\code_version_0.0.2\\TemplateFiles"
QueryTemplate=TemplateFilePath+"\queryTemplate.txt"
OutputQueryFileName="VALUE"

DatabasePath="D:\\SPLM\\Tacton_Integration\\ToProcess\\Database"
TableName="ControlTable.db"

conn=sqlite3.connect(DatabasePath+"\\"+TableName)
c = conn.cursor()



####To run this file manually python createInterimFiles_v0.1.py solutionID123 Line1 D:\SPLM\Tacton_Integration\ToProcess\solutionID123_Line1


def create_qty_rolledup(BomFile):
	#print(BomFile)
	df = pd.read_excel(SOURCE_FILE_PATH+"\\"+FileName)
	df["TEMP"]=np.nan
	
	df['rev']=np.where(df['Create']=='Yes','AUTOASSIGN',0)
	df['Part No']=np.where(df['Create'] == 'Yes','AUTOASSIGN',df['Part No'])
	
	r1=df['Level'].count()
	#print("R1:",r1)
	counter = []
	i=0
	j=1
	
	for f in range(df['Level'].count()):
	#print("Data Frame value:",df.iloc[f,0])
	#print(df.iloc[i,0])
	#print("I:",i, "J:",j)
	
		if (j==r1):
			#print("I am in break",i,j,r1)
			break
	
		elif df.iloc[i,0] != df.iloc[j,0]:
			counter.append('1')
			df.iloc[i,9]='1'
	
		elif df.iloc[i,0] == df.iloc[j,0]:
			if df.iloc[i,1] != df.iloc[j,1]:
				counter.append('1')
				df.iloc[i,9]='1'
		
		#elif df.iloc[i,1] == df.iloc[j,1]:
		#	print("In second Loop")
		
		i +=1
		j+=1
	
	df.iloc[i,9]='1' #The last line in excel has to be 1
	#print(df)
	
	#print(counter)
	#print(df)
	df2=df[df['TEMP'].isnull()]
	#print(df2)	

	#print(df2.groupby(['Level','Part No']).Level.count())

	df2=df2.groupby(['Level','Part No','Description']).size().reset_index(name='count')
	df2['count']=df2['count']+1
	#df2['BOM Item No']=df['BOM Item No']
	#print (df2)


	df3=pd.merge(df,df2,on=['Level','Part No','Description'],how='left')
	df3['count']=df3['count'].fillna(1).apply(np.int64)
	df3.rename(columns={'count':'qty'},inplace=True)
	#print(df3)
	#df3['qty']=df3['count']
	#df4=df3['Level','Part No','qty']
	df4=pd.DataFrame.copy(df3)
	df4=df4.drop_duplicates(['Level','Part No','Description','qty'],keep='last')
	#print(df4)
	return (df4)


def create_InterimFile(df):
	#print(df)
	writer=ExcelWriter(InterimLoadFilePath+"\\"+InterimLoadFile)
	df.to_excel(writer,'Sheet1',index=False)
	writer.save()

	
def create_QueryLoadFile(df4):
	dfQuery=df4[df4['Part No'] != 'AUTOASSIGN']
	dfQuery=dfQuery['Part No']
	#print(dfQuery)
	
	#Converting the column Part No's into a string delimited by ;
	string1 = dfQuery.to_string(header=False, index=False).split('\n')
	vals=[','.join(ele.split()) for ele in string1]
	queryString = ";".join(vals)
	
	OutputQueryFileName = QueryLoadFilePath+"\\"+"QueryLoadFile.xls"
	#print(OutputQueryFileName)

	# Create Query.xml file from query string to get revision id's
	try:
		with open(QueryTemplate,'r') as fileTemplate:
			templatedata=fileTemplate.read()
			QUERYDATA = templatedata.replace('THMXXXXXXX',queryString)
			file = open(QueryLoadFilePath+"\\"+QueryLoadFile,"w")
			file.write(QUERYDATA)
	except:
		print("file is not available for writing")
	finally:
		fileTemplate.close()
		file.close()

def update_Data(SID,LID):
	#c.execute('SELECT * FROM controlSheet')
	c.execute ("UPDATE controlSheet SET QueryOutput = 50, InterimLoadFIle = 1, SourceBOM = 999 WHERE SolutionID = (?) and LineID = (?)",[SID,LID])
	conn.commit()

def return_AutoAssign(InterimDataFrame):
	#print(InterimDataFrame)
	InterimDataFrame=InterimDataFrame.loc[InterimDataFrame['Part No']=='AUTOASSIGN']
	
	#print(InterimDataFrame)
	
	
	String1 = "#COL level item rev name descr type link_root plant_root consumed resource required workarea attributes owner group predecessor duration act_name act_desc occ_note occ_eff abs_occ qty uom seq matrix status occs activities loadif filePath"
	String2 = "#DELIMITER #"
	try:
		file =open(AutoAssign1,'w')
		file.write(String1+"\n"+String2+"\n")
	except:
		print("cannot able to write the file")
	finally:
		file.close()
	#print(dCreate)

	dAutoAssign=pd.DataFrame.copy(InterimDataFrame)
	dAutoAssign.rename(columns={'Level':'level','Part No':'item','Description':'descr'},inplace=True)
	dAutoAssign['type']='Item'
	dAutoAssign['name']='AUTOASSIGN'
	
	#dCreate=dCreate[['level','item','rev','name','descr','type','link_root','plant_root','consumed','resource','required','workarea','attributes','owner','group','predecessor','duration','act_name','act_desc','occ_note','occ_eff','abs_occ','qty','uom','seq']]
		
	
	dAutoAssign=dAutoAssign[['level','item','rev','name','descr','type']]
	#print(dAutoAssign)
	dAutoAssign.to_csv(AutoAssign1,mode='a',sep='#',index=False,header=False)
	#print(dfAuto)
	
	
InterimDataFrame=create_qty_rolledup(FilePath)
create_InterimFile(InterimDataFrame)
create_QueryLoadFile(InterimDataFrame)
update_Data(arg1,arg2)
return_AutoAssign(InterimDataFrame)

c.close()
conn.close()
