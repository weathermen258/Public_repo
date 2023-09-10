import pyodbc
import os
mdb_file = 'C:/Users/DAO ANH CONG/AppData/Local/Programs/'+\
           'Python/Python37/Scripts/DBdienbaoKTTVHV.accdb' # open file browsing dialog box
##[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]
print(mdb_file)
mdb_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
conn = pyodbc.connect('DRIVER={}; DBQ={}'.format(mdb_driver,mdb_file))
cursor = conn.cursor()
command1= "select TxTxTxtb from CLIM \
where STNO='{}' and yYear='{}'".format("48/66","2018")
print(command1)
cursor.execute(command1)
list1 = cursor.fetchall()
final_result = [list(i) for i in list1]
print (final_result)

import pandas as pd
command0 = 'select * from CLIM'
table_data = pd.read_sql(command0,conn)
df = pd.DataFrame(table_data)
select1 = df.loc[(df['STNO']=='48/66') & (df['yYear'] == '2018')]
test3=select1[['mMonth','TxTxTxtb']]
print(test3)
test2=select1['TxTxTxtb']
print (test2)
list2 = list(test2)
sum = 0
for i in range(len(list2)):
    number = int(list2[i])/10
    sum += number
    print (number)
print(sum)
