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
      wt.write("{},{},{},{},{},{},{},{},{},{},{},{}\n"\
                .format("stations","obs_year","obs_month","obs_day",\
                        "obs_hour","obs_min","wind_dir","wind_speed",\
                        "slp","delta_p","gust_dir","gust_speed"))

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
## function to get data from each typh line
def get_record_typh(str):
   "This function get typh data"
   print (str)
   typh_len = len(str.split())
   #print(typh_len)
   stations.append(str.split()[2][0:3])
   wind_dir.append(dir_convert(str.split()[3][1:3]))
   wind_speed.append(str.split()[3][3:5])
   if (typh_len >= 9):
      if (float(str.split()[7][1:5]) < 1000):
         slp.append(1000+float(str.split()[7][1:5])/10)
      elif(float(str.split()[7][1:5]) > 1000):
         slp.append(float(str.split()[7][1:5])/10)
   else:
      slp.append("")

   if (typh_len >= 9):
      if (int(str.split()[8][0:2]) == 58):
           delta_p.append(float(str.split()[8][2:5])/10)
      elif (int(str.split()[8][0:2]) == 59):
           delta_p.append(0 - float(str.split()[8][2:5])/10)
   else:
      delta_p.append("")
   
   if (typh_len >= 10):
      gust_dir.append(dir_convert(str.split()[10][1:3]))
      gust_speed.append(str.split()[10][3:5])
   else:
      gust_dir.append("")
      gust_speed.append("")
   #print(stations,wind_dir,wind_speed, \
   #  slp,delta_p,gust_dir,gust_speed)
   return

## Get data from input file
Tk().withdraw() # disable the root windows
input_file = askopenfilename() # open file browsing dialog box
print(input_file)

## Read raw data from file
import datetime
with open(input_file,'r') as dt:
   lines = dt.read()
   record_temp = lines.split("\n")
   #print(record_temp)
   for i in range(record_temp.count('')):
      record_temp.remove('')
   typh_type = record_temp[2][0:4].lower() #find obs typh type aaxx or typh
   print (typh_type)
   obs_time = record_temp[1].split()[2]     #get the obs time
   obs_year = datetime.datetime.now().year   #define year and month
   obs_month = datetime.datetime.now().month
   obs_day =  obs_time[0:2]
   obs_hour = obs_time[2:4]
   obs_min = obs_time[4:6]
   dt.close()
   
print(obs_year, obs_month, obs_day, obs_hour, obs_min)

record_typh = []
for i in range(len(record_temp)):
   if (len(record_temp[i].split()) >= 6):
      record_typh.append(record_temp[i])
## get data from raw data 
for i in range(len(record_typh)):
   print(record_typh[i])
   get_record_typh(record_typh[i])
   
## Write data to file
def write_record(file):
   with open(file,'a') as wt:
      for i in range(len(stations)):
         wt.write("{},{},{},{},{},{},{},{},{},{},{},{}\n"\
            .format(stations[i],obs_year,obs_month,obs_day,obs_hour,obs_min,wind_dir[i],\
            wind_speed[i],slp[i],delta_p[i],gust_dir[i],gust_speed[i]))
   wt.close
   return
write_record(output_file)  


