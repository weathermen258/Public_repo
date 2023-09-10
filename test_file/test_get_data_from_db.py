import pyodbc
import os
mdb_file = 'C:/Users/DAO ANH CONG/Documents/'+\
           'DBdienbaoKTTVHV.accdb'
##[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]
print(mdb_file)
mdb_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
conn = pyodbc.connect('DRIVER={}; DBQ={}'.format(mdb_driver,mdb_file))
cursor = conn.cursor()
stations=['48842','48/67','48/68','48/69','48840','48/70','48/72',\
          '48/66','48/74','48844','48/75','48/76','48/79','48/77',\
          '48/80','48/81','48845','48/82','48846','48/84','48/73','48/86']

## make list for meteorological elements
## function to get data from 
def get_data_db() :
   for station in stations:
      command0= "select hHOUR,TTT,RRR,DD,FF,UUU,N,PPPP,VV,R24R24,\
TnTn,TxTx,TnTn555 from Dienbaokhituong where STNO='{}' and yYear='{}' \
and mMonth='{}' and dDay='{}' and hHOUR='{}'".\
format(station,obs_year,obs_month,obs_day,obs_hour)
      #print(command0)
      cursor.execute(command0)
      record = cursor.fetchall()
      #print (record)
      if not (list(record)):
         time.append('')
         temp.append('')
         rain.append('')
         wind_dir.append('')
         wind_speed.append('')
         humid.append('')
         cloud.append('')
         slp.append('')
         visibility.append('')
         total_rain.append('')
         Tmin.append('')
         Tmax.append('')
         Tn555.append('')
      else:
         for row in record:
            if not (list(row)[0]): # get time
               time.append('')
            else:
               time.append(list(row)[0])
            if not (list(row)[1]): # get temperature
               temp.append('')
            else:
               temp.append(list(row)[1])
            if not (list(row)[2]): # get rain 6 hour
               rain.append('')
            else:
               rain.append(list(row)[2])
            if not (list(row)[3]): # get wind direction
               wind_dir.append('')
            else:
               wind_dir.append(list(row)[3])
            if not (list(row)[4]):  # get wind speed
               wind_speed.append('')
            else:
               wind_speed.append(list(row)[4])
            if not (list(row)[5]): # get humidity
               humid.append('')
            else:
               humid.append(list(row)[5])
            if not (list(row)[6]): # get cloud cover
               cloud.append('')
            else:
               cloud.append(list(row)[6])
            if not (list(row)[7]): # get sea level pressure
               slp.append('')
            else:
               slp.append(list(row)[7])
            if not (list(row)[8]): # get visibility
               visibility.append('')
            else:
               visibility.append(list(row)[8])
            if not (list(row)[9]): # get rain total 24hour
               total_rain.append('')
            else:
               total_rain.append(list(row)[9])
            if not (list(row)[10]): # get Tmin over 24h
               Tmin.append('')
            else:
               Tmin.append(list(row)[10])
            if not (list(row)[11]): # get Tmax over 24h
               Tmax.append('')
            else:
               Tmax.append(list(row)[11])
            if not (list(row)[12]): # get Tn555 re-observed
               Tn555.append('')
            else:
               Tn555.append(list(row)[12])
   return

### define the position for each element to write to excel file
def col_temp(hour):
   switcher = {
      '15':'C','18':'D','21':'E','00':'F','03':'G','06':'H',
      '09':'I','12':'J'
      }
   return switcher.get(hour,'invalid hour')
def col_rain(hour):
   switcher = {
      '18':'N','00':'O','06':'P','12':'Q'
      }
   return switcher.get(hour,'invalid hour')
def col_wind(hour):
   switcher = {
      '15':'S','18':'T','21':'U','00':'V','03':'W','06':'X',
      '09':'Y','12':'Z'
      }
   return switcher.get(hour,'invalid hour')
def col_humid(hour):
   switcher = {
      '15':'AA','18':'AB','21':'AC','00':'AD','03':'AE','06':'AF',
      '09':'AG','12':'AH'
      }
   return switcher.get(hour,'invalid hour')
def col_cloud(hour):
   switcher = {
      '15':'C','18':'D','21':'E','00':'F','03':'G','06':'H',
      '09':'I','12':'J'
      }
   return switcher.get(hour,'invalid hour')
def col_slp(hour):
   switcher = {
      '15':'K','18':'L','21':'M','00':'N','03':'O','06':'P',
      '09':'Q','12':'R'
      }
   return switcher.get(hour,'invalid hour')
def col_visi(hour):
   switcher = {
      '15':'S','18':'T','21':'U','00':'V','03':'W','06':'X',
      '09':'Y','12':'Z'
      }
   return switcher.get(hour,'invalid hour')

### Do a little test with leap year
def leap_year(obs_year):
   if (obs_year % 400 == 0) or \
      (obs_year % 4 == 0 and obs_year % 100 != 0):
      check = True
   else:
      check = False
   return check

### Define the writing to write to excel file

def write_met_obs():
   ### import neccesary packages
   import openpyxl
   from openpyxl import Workbook
   from openpyxl import load_workbook
   
   ### Define the work sheet name
   work_year = obs_year
   if (obs_hour in ['15','18','21']):    
      if ((obs_month in ['04','06','09','11']) and (obs_day == '30')) or\
         ((obs_month in ['01','03','05','07','08','10']) and (obs_day == '31')) or\
         ((leap_year(obs_year) == True) and obs_month == '02' and obs_day == '29') or\
         ((leap_year(obs_year) == False) and obs_month == '02' and obs_day == '29'):
            work_day = '01'
            if ((int(obs_month) +1) > 10):
               work_month = str(int(obs_month) +1)
            else:
               work_month = '0'+str(int(obs_month) +1)
      elif ((obs_month == '12') and (obs_day == '31')):
         work_year = obs_year + 1
         work_month = '01'
         work_day = '01'
      else:
         work_day = str(int(obs_day) + 1)
         work_month = obs_month
   else:
      work_day = int(obs_day)
      work_month = obs_month
   if int(work_day) < 10:
      work_sheet = 'Ngay' + '0' + str(int(work_day))
   else:
      work_sheet = 'Ngay' + str(work_day)
   output_file = 'Synop' + str(work_year) + work_month + '.xlsx'
   print (output_file)
   import os.path
   from shutil import copy2
   if os.path.isfile(output_file):
      print('File already exist')
   else:
      print('File doesnt exist, copy new file')
      copy2('template.xlsx',output_file)
   wb = load_workbook(output_file)
   sheet0 = wb[work_sheet]
   print(work_sheet,'***',obs_hour)
   
   ### Write met data to excel file
   for i in range(0,len(stations)):
      ### Write temperature
      cell_temp = col_temp(obs_hour) + str(4+i)
      if (temp[i] != ''):
         sheet0[cell_temp] = int(temp[i])/10
      
      ### Write rainfall
      if (obs_hour in ['18','00','06','12']):
         cell_rain = col_rain(obs_hour) + str(4+i)
         if(rain[i] != ''):
            if (float(rain[i])) < 1:
               sheet0[cell_rain] = float(rain[i])
            else:
               sheet0[cell_rain] = int(rain[i])
         else:
            sheet0[cell_rain] = '-'
      #else:
         #print('no rain at this hour')

      ### Write Wind direction and speed
      cell_wind = col_wind(obs_hour) + str(4+i)
      if (wind_dir[i] != ''):
         sheet0[cell_wind] = wind_dir[i] + wind_speed[i]
      else:
         sheet0[cell_wind] = '-'

      ### Write Humidity
      cell_humid = col_humid(obs_hour) + str(4+i)
      if (humid[i] != ''):
         sheet0[cell_humid] = round(float(humid[i]))

      ### Write cloud cover
      cell_cloud = col_cloud(obs_hour) + str(31+i)
      if (cloud[i] != ''):
         if (cloud[i][:2] == '10'):
            sheet0[cell_cloud] = cloud[i][:5]
         else:
            sheet0[cell_cloud] = cloud[i][:4]

      ### Write sea level pressure
      cell_slp = col_slp(obs_hour) + str(31+i)
      if (slp[i] != ''):
         if float(slp[i]) > 2000:
            sheet0[cell_slp] = int(slp[i])/10
         else:
            sheet0[cell_slp] = round(float(slp[i]))

      ### Write visibility
      cell_visi = col_visi(obs_hour) + str(31+i)
      if (visibility[i] != ''):
         sheet0[cell_visi] = round(float(visibility[i]))

      ### Write Tmax
      if (obs_hour == '12'):
         cell_tmax = 'M' + str(4+i)
         if (Tmax[i] != ''):
            sheet0[cell_tmax] = int(Tmax[i])/10

      ### Write Tmin
      ### Can check xem giá trị có trống hay không
      cell_tmin = 'L' + str(4+i)
      if (Tmin[i] != ''):
         sheet0[cell_tmin] = int(Tmin[i])/10
         
      ### Write Tmin that was re-observed
     #if (obs_hour in ['06','12']):
      cell_tmin = 'L' + str(4+i)
      if (Tn555[i] != ''):
         sheet0[cell_tmin] = int(Tn555[i])/10

      ### Write R24
      cell_rain24 = 'R' + str(4+i)
      if (obs_hour == '12'):
         if (total_rain[i] not in ['','0000']):
            sheet0[cell_rain24] = int(total_rain[i])/10
         else:
            sheet0[cell_rain24] = '-'
   wb.save(output_file)
   return

# get the real time for automatic data
import datetime
obs_year = datetime.datetime.utcnow().year
utc_month = datetime.datetime.utcnow().month
if (utc_month > 9):
   obs_month = utc_month
else:
   obs_month = '0' + str(utc_month)
utc_day = datetime.datetime.utcnow().day
if (utc_day > 9):
   obs_day = utc_day
else:
   obs_day = '0' + str(utc_day)
test_hour = datetime.datetime.utcnow().hour
if (0 <= test_hour < 3):
   obs_hour = '00'
elif (3 <= test_hour < 6):
   obs_hour = '03'
elif (6 <= test_hour < 9):
   obs_hour = '06'
elif (9 <= test_hour < 12):
   obs_hour = '09'
elif (12 <= test_hour < 15):
   obs_hour = '12'
elif (15 <= test_hour < 18):
   obs_hour = '15'
elif (18 <= test_hour < 21):
   obs_hour = '18'
elif (21 <= test_hour < 23):
   obs_hour = '21'
print(obs_year,obs_month,obs_day,obs_hour)

## Do a little test
### for testing purpose
obs_year = 2019
obs_month = '12'
obs_day_test = ['31']
obs_hour_test = ['00','03','06','09','12','15','18','21']
#output_file = 'sample.xlsx'
###
for obs_day in obs_day_test:
   for obs_hour in obs_hour_test:
      time = []
      temp = []
      rain = []
      wind_dir = []
      wind_speed = []
      humid = []
      cloud = []
      slp = []
      visibility = []
      total_rain = []
      Tmin = []
      Tmax = []
      Tn555 = []
      get_data_db()  ## this part get the damn data!
      #print (Tmin)
      #print (Tmax)
      #for i,k,l,m,n,o,p in zip(time,stations,temp,Tmax,Tmin,rain,total_rain):
         #print(i,'**',k,'**',l,'**',m,'**',n,'**',o,'**',p)
      write_met_obs() ### This part Write data to excel files


