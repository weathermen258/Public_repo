## import essential libraries 
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
## input typhoon name and make file
ty_name = input("Enter typhoon name:  ")
cwd = os.getcwd()
output_file = os.path.join(cwd,ty_name+".csv")
print(f'file name is {output_file}')
if os.path.isfile(output_file):
   print("File already exist, proceed !")
else:
   with open(output_file,'w') as wt:
      wt.write("{},{},{},{},{},{},{},{},{},{},{},{},{}\n"\
                .format("typhoon_name","stations","obs_year","obs_month","obs_day",\
                        "obs_hour","obs_min","wind_dir","wind_speed",\
                        "slp","delta_p","gust_dir","gust_speed"))
   wt.close()
## making array of data
stations = []
wind_speed = []
wind_dir = []
slp =[]
delta_p = []
gust_speed = []
gust_dir =[]
pmin = []
time_pmin = []
wind_max = []
time_wind_max = []

### wind direction conversion
def dir_convert(str):
   switcher = {
      '00': '00','02': 'NNE','05': 'NE','07': 'ENE', '09': 'E',
      '11': 'ESE', '14': 'SE', '16': 'SSE', '18': 'S',
      '20': 'SSW', '23': 'SW', '25': 'WSW', '27': 'W',
      '29': 'WNW', '32': 'NW', '34': 'NNW', '36': 'N','': ''
      }
   return switcher.get(str,"invalid wind direction")
## function to get data from each typh line from typh TV
def get_record_synop(str):
   print('teste')
   return
def get_record_TV(str):
   typh_len = len(str.split())
   stations.append(str.split()[2])
   wind_dir.append(dir_convert(str.split()[3][1:3]))
   wind_speed.append(str.split()[3][3:5])
   slp.append("")
   delta_p.append("")
   if (typh_len > 5):
      gust_dir.append(dir_convert(str.split()[4][1:3]))
      gust_speed.append(str.split()[4][3:5])
   else:
      gust_dir.append("")
      gust_speed.append("")
## function to get data from each typh line from typh KT

def get_record_typh(str):
   "This function get typh data"
   #print (str.split())
   typh_len = len(str.split())
   #print(typh_len)
   stations.append('48'+str.split()[2][0:3])
   wind_dir.append(dir_convert(str.split()[3][1:3]))
   wind_speed.append(str.split()[3][3:5])
   if (8 < typh_len <= 9):
      gust_dir.append(dir_convert(str.split()[7][1:3]))
      gust_speed.append(str.split()[7][3:5])
      
   if (typh_len > 9):
      if (float(str.split()[7][1:5]) < 1000):
         slp.append(1000+float(str.split()[7][1:5])/10)
      elif(float(str.split()[7][1:5]) > 1000):
         slp.append(float(str.split()[7][1:5])/10)
   else:
      slp.append("")
      
   if (typh_len > 10):
      if (int(str.split()[8][0:2]) == 58):
           delta_p.append(float(str.split()[8][2:5])/10)
      elif (int(str.split()[8][0:2]) == 59):
           delta_p.append(0 - float(str.split()[8][2:5])/10)
   else:
      delta_p.append("")
   
   if (typh_len > 10):
      gust_dir.append(dir_convert(str.split()[10][1:3]))
      gust_speed.append(str.split()[10][3:5])
   else:
      gust_dir.append("")
      gust_speed.append("")
   if (str.split()[2][0:3] == '/68' or str.split()[2][0:3] == '/81'):
      if 'E' in str:
         for i in range(len(str.split())):
            if str.split()[i][0:1] == '9':
               print('this typh is stupid')
               gust_dir.append(dir_convert(str.split()[i][1:3]))
               gust_speed.append(str.split()[i][3:5])
   #print(stations,wind_dir,wind_speed, \
   #  slp,delta_p,gust_dir,gust_speed)
   return

## Get data from input file

Tk().withdraw() # disable the root windows
input_file = askopenfilename() # open file browsing dialog box
#print(input_file)

## Read raw data from file
import datetime
with open(input_file,'r') as dt:
   lines = dt.read()
   record_temp = lines.split("\n")
   #print(record_temp)
   for i in range(record_temp.count('')):
      record_temp.remove('')
   typh_type = record_temp[2].lower() #find obs typh type aaxx or typh
   print (typh_type)
   obs_time = record_temp[1].split()[2]     #get the obs time
   print (obs_time)
   obs_year = datetime.datetime.now().year   #define year and month
   obs_month = datetime.datetime.now().month
   obs_day =  obs_time[0:2]
   obs_hour = obs_time[2:4]
   obs_min = obs_time[4:6]
   dt.close()
# print(obs_year, obs_month, obs_day, obs_hour, obs_min)
record_typh = []

for i in range(len(record_temp)):
   if (len(record_temp[i].split()) >= 5):
      record_typh.append(record_temp[i])

## get data from raw data
## (lam lai doan nay)
for i in range(len(record_typh)):
   #print(record_typh[i])
   if (len(record_typh[i].split()) >= 8):
      get_record_typh(record_typh[i])
   else:
      get_record_TV(record_typh[i])
###
      
## Write data to file
def overwrite_record(file):
   dt = open(output_file,'r')
   record_temp = dt.read()
   lines = record_temp.split("\n")
   dt.close()
   #print (lines)
   for i in range(len(stations)):
      check_line = f'{ty_name},{stations[i]},{obs_year},{obs_month},\
{obs_day},{obs_hour},{obs_min}'
      write_line = "{},{},{},{},{},{},{},{},{},{},{},{},{}"\
            .format(ty_name,stations[i],obs_year,obs_month,obs_day,\
                     obs_hour,obs_min,wind_dir[i],wind_speed[i],slp[i],\
                     delta_p[i],gust_dir[i],gust_speed[i])
      for k in range(0,len(lines)):
         if check_line in lines[k]:
            #print ('this line is replacable')
            lines[k] = write_line
            break
   
   with open('typh_temp.csv','w') as wt:
      for i in range(len(lines)):
         #print(lines[i])
         wt.write(lines[i] + '\n')
   wt.close()
   import os
   from shutil import copy2
   os.remove(output_file)
   copy2('typh_temp.csv',output_file)
   return

def write_record(file):
   with open(file,'a+') as wt:
      for i in range(len(stations)):
         wt.write("{},{},{},{},{},{},{},{},{},{},{},{},{}\n"\
            .format(ty_name,stations[i],obs_year,obs_month,\
                    obs_day,obs_hour,obs_min,wind_dir[i],\
                    wind_speed[i],slp[i],delta_p[i],gust_dir[i],\
                    gust_speed[i]))
   wt.close()
   return

# Write or overwirte data file
check_line = f'{ty_name},{stations[i]},{obs_year},{obs_month},\
{obs_day},{obs_hour},{obs_min}'
print (check_line)
rt = open(output_file,'r')
line_temp = rt.read()
#print (line_temp)
rt.close()

print (stations)
if check_line in line_temp:
   print ('This file is replacable')
   print ('******************')
   overwrite_record(output_file)
else:
   write_record(output_file)


