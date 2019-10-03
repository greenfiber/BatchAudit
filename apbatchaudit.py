import pyodbc
import logging
from os import listdir
from os.path import isfile, join,splitext
import pandas as pd
import xlwings as xw
from secret import secrets as secrets
cx = pyodbc.connect("DSN=gf32;UID={};PWD={}".format(
    secrets.dbusr, secrets.dbpw))

def getbatchesfromdb():
    query= '''
   SELECT distinct 
       [UDF_BATCH_NO],
	   TRANSACTIONDATE
     
        FROM [MAS_GFC].[dbo].[AP_INVOICEHISTORYHEADER]
        where convert(varchar(8),TRANSACTIONDATE,112) between '20190901' and '20190913'
        order by TRANSACTIONDATE desc
    
    '''
    rows=cx.execute(query)
    cursor = cx.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    data=[]
    for row in rows:
        data.append(row)
    return data
print("Generating batch file list...")
batchfiles = [f for f in listdir('L:\\APBatches') if isfile(join('L:\\APBatches',f)) and f.endswith('.pdf')]
print("Generating POBatches list...")
pofiles = [f for f in listdir('L:\\POBatches') if isfile(join('L:\\POBatches',f)) and f.endswith('.pdf')]

filenames =[]
pofilenames=[]
for file in batchfiles:
    filenames.append(os.path.splitext(file)[0])
for file in pofiles:
    pofilenames.append(os.path.splitext(file)[0])
db=[]
for row in rows:
    db.append(str(row.UDF_BATCH_NO))
missing=set(db)-set(filenames)
missing-=set(pofilenames)
df = pd.DataFrame(missing)
wb= xw.Book()
sheet= wb.sheets['Sheet1']
sheet.range('A1').value=df