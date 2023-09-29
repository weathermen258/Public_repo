from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
input_file = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(input_file)

import datetime
with open(input_file,'r') as dt:
   lines = dt.read()
   record_temp = lines.split("\n")
   #print(record_temp)
   for i in range(record_temp.count('')):
      record_temp.remove('')
   typh_type = record_temp[2][0:4].lower() #find obs typh type aaxx or typh
   print (typh_type)
   obs_time = record_temp[1].split()[2]
   obs_year = datetime.datetime.now().year
   obs_month = datetime.datetime.now().month
   obs_day =  obs_time[0:2]
   obs_hour = obs_time[2:4]
   obs_min = obs_time[4:6]
   data = dt.readlines()
print(obs_year, obs_month, obs_day, obs_hour, obs_min)
record_typh = []
for i in range(len(record_temp)):
   if (len(record_temp[i].split()) >= 6):
      record_typh.append(record_temp[i])
for i in range(len(record_typh)):
   print(record_typh[i])

print(data)
