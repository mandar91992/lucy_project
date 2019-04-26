# -*- coding: utf-8 -*-
"""
Created on Thu Apr 25 11:45:38 2019

@author: Admin
"""

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
import re
import xlrd
import xlsxwriter
####To run this file manually python createLoadFile_v0.1.py solutionID789 Line2 C:\Temp\ToProcess\solutionID789_Line2


arg1=str(sys.argv[1])
arg2=str(sys.argv[2])
arg3=str(sys.argv[3])

#print("I am in createInterimFile")
#print("The number of arguments are:",len(sys.argv))
#print("The arguments are:",str(sys.argv))

print("ARG1:",arg1,"ARG2:",arg2,"ARG3:",arg3)

path=InterimLoadFilePath=arg3

###########	TCEXCEL IMPORT PATH Initialization
SolutionFileName=arg1+"_"+arg2+".pim"
SolutionFile=arg3+"\\"+SolutionFileName


###########	PS_UPLOAD
SolutionFileName1=arg1+"_"+arg2+".txt"
SolutionFile1=arg3+"\\"+SolutionFileName1

TEMPLoadFile=InterimLoadFile=""
NewItemsFileName=arg1+"_"+arg2+".xlsx"
NewItemFilePath=arg3+"\\"+NewItemsFileName

extension = '.csv'
TempLoadFilePath=[]


def create_xlsx(mylist):
	#path='abc.xlsx' 
	workbook=xlsxwriter.Workbook(NewItemFilePath) 
	worksheet=workbook.add_worksheet() 
	r=0 
	c=0

	worksheet.write(r,c,"Part No")
	worksheet.write(r,c+1,"Description")
	r+=1
	
	for item,des in (mylist):
		worksheet.write(r,c,item)
		worksheet.write(r,c+1,des)
		r +=1
	workbook.close()
	return 0 

	
def send_emails():
    import smtplib

    #server = smtplib.SMTP_SSL('smtp.gmail.com', 465) #SMTP address of the company domain
    #server.login("Email from which email has to be sent", "password of email")
    #email_list=["love4css@gmail.com", "mandar.joshi@faithplm.com"]
    #msg = """ Hello, The report for the newly created Items are
    #       {item_name}
    #       {descriptions}
    #      """.format(item_name = item, descriptions = description) 
    #      server.sendmail("akvns152@gmail.com", email_list, msg)
    #      server.quit()

		  
def get_newItems(logFile):
	item=[]
	description=[] 
	
    try:
        #with open('tcexcel_solutionID123_Line1_AutoAssign.pim.log','r') as f:
    
    
	with open(logFile,'r') as f:
				for line in f:
							if "DEBUG(build_structure)New item created:" in line:
										first=(line.split(":")[-1][1:11])
										# print(first)
										item.append(first)
							elif "DEBUG(build_structure): Add item description" in line:
										second= (line.split("DEBUG(build_structure): Add item description")[-1][2:-2])
										# print(second)
										description.append(second)
	# print(item)
	# print(description)
	mylist=[] 
	for a,items in enumerate(item):
				for b,desc in enumerate(description):
							if a==b:
										mylist.append([items,desc])
	create_xlsx(mylist)
	# send_emails()
    except:
        print("Something went wrong when reading to the file")
    finally:
        f.close()



def create_TcExcelImportFile(dCreate):
	########## TCEXCEL_IMPORT
	
	String1 = "#COL level item rev name descr type link_root plant_root consumed resource required workarea attributes owner group predecessor duration act_name act_desc occ_note occ_eff abs_occ qty uom seq matrix status occs activities loadif filePath"
	String2 = "#DELIMITER #"
    try:
        file =open(SolutionFile,'w')
        file.write(String1+"\n"+String2+"\n")
    except:
        print("Something went wrong while writing the file")
    finally:
        file.close()
	#print(dCreate)


	dCreate=dCreate[['level','item','rev','name','descr','type','link_root','plant_root','consumed','resource','required','workarea','attributes','owner','group','predecessor','duration','act_name','act_desc','occ_note','occ_eff','abs_occ','qty','uom','seq']]
	
	dAutoAssign=pd.DataFrame.copy(dCreate)
	#dCreate=np.where(dfCreateItem['item'].values=='AUTOASSIGN')
	#print(dCreate.iloc[np.where(dCreate.item.values=='AUTOASSIGN')])
	dAutoAssign=dCreate.iloc[np.where(dCreate.item.values=='AUTOASSIGN')]
	dAutoAssign['seq']='-'
	
	dAutoAssign.to_csv(SolutionFile,mode='a',sep='#',index=False,header=False)

	
def create_PsUploadFile(dCreate):

	String1 = "#DELIMITER ,"
	String2 = "#COL level type item rev name qty seq occs revname"
    try:
        file =open(SolutionFile1,'w')
        file.write(String1+"\n"+String2+"\n")
    except:
        print("Something went wrong while writing the file")
    finally:
        file.close()
	
	dCreate=dCreate[['level','type','item','rev','name','qty','seq']]
	
	dCreate.to_csv(SolutionFile1,mode='a',sep=',',index=False,header=False)

	
	

for root, dirs_list, files_list in os.walk(path):
	for file_name in files_list:
		if os.path.splitext(file_name)[-1] == extension:
			TempLoad_path=os.path.join(root)
			TempLoadFilePath.append(os.path.join(root,file_name))
				
			#print(file_name_path)
			#print(TempLoadFilePath)
				
			#This line get all the files in the directory in a list	
			x = [os.path.join(r,file) for r,d,f in os.walk(TempLoad_path) for file in f]
			
			#print(len(x))
			#print("XXXXXXXXXXXXX:",x)
				
				


TEMPLoadFile=TempLoadFilePath[0]
#print("TempLoadFile:",TEMPLoadFile)

r=re.compile(".*InterimLoadFile*")
val=list(filter(r.match,x))

InterimLoadFile=val[0]
#print("Interim LoadFile:",InterimLoadFile)

logFormat=re.compile(".*pim.log")
logF=list(filter(logFormat.match,x))
logFile=logF[0]
#print("LogFile:",logFile)

get_newItems(logFile)
    try:
        dfNewItems = pd.read_excel(NewItemFilePath)
        print(dfNewItems)
    except:
        print("excel file is not present")
    
        

df = pd.read_csv(TEMPLoadFile,skiprows=2)
df2= pd.DataFrame.copy(df)


df2.columns=['Part No','ItemName','RevID','ProjectID','ProductGroupCode','ReleasedStatus','DateReleased','Initiator','StartDate','Approver']
df2=df2[['Part No','RevID']]

print(list(df2.columns.values))
#print(df2)


df3=pd.read_excel(InterimLoadFile)
print(df3)
#Level  BOM Item No  Part No Description  Qty Create Reuse  Requires engineering  Other         rev  count

#dfMergedNewItems = pd.merge(df3,dfNewItems,on=['Description'],how='inner',sort=False)
dfMergedNewItems=(df3[['Level','BOM Item No','Part No','Description','rev','qty']].merge(dfNewItems,on='Description',how='left'))
print(dfMergedNewItems)


'''
#df4=pd.merge(df2,df3, on=['Part No'], how='outer',sort=False)
#df4=df4.reset_index(drop=True)
#print(df3)
df4=pd.DataFrame()
df4=(df3[['Level','BOM Item No','Part No','Description','rev','qty']].merge(df2,on='Part No',how='left'))

df4['RevID']=np.where(df4['RevID'].isnull(),'AUTOASSIGN',df4['RevID'])
#print(df4)

#dfFinalMerge=(df4[['Level','BOM Item No','Part No','Description','rev','qty']].merge(dfNewIs,on='Description',how='left'))

dfCreateItem=pd.DataFrame.copy(df4)
dfCreateItem.rename(columns={'rev':'oldrev'},inplace=True)

dfCreateItem.rename(columns={'Part No':'item','RevID':'rev','Level':'level','Description':'descr','BOM Item No':'seq','ItemName':'name'},inplace=True)

print(dfCreateItem)






dfCreateItem['name']=dfCreateItem['item']
dfCreateItem['type']='Item'
dfCreateItem['type']=np.where(dfCreateItem['item'].str.contains('THM*|AUTOASSIGN',regex=True),dfCreateItem['type'],'Component')
dfCreateItem['descr']=np.where(dfCreateItem['item'].str.contains('AUTOASSIGN',regex=True),dfCreateItem['descr'],'-')



dfCreateItem['link_root']='-'
dfCreateItem['plant_root']='-'
dfCreateItem['consumed']='-'
dfCreateItem['resource']='-'
dfCreateItem['required']='-'
dfCreateItem['workarea']='-'
dfCreateItem['attributes']='-'
dfCreateItem['owner']='-'
dfCreateItem['group']='-'
dfCreateItem['predecessor']='-'
dfCreateItem['duration']='-'
dfCreateItem['act_name']='-'
dfCreateItem['act_desc']='-'
dfCreateItem['occ_note']='-'
dfCreateItem['occ_eff']='-'
dfCreateItem['abs_occ']='-'
dfCreateItem['uom']='-'



create_TcExcelImportFile(dfCreateItem)

#create_PsUploadFile(dfCreateItem)


'''