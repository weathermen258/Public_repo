## This script get data from synop obs
def cloud_convert(str0):
   switcher = {
      '0': '0/10', '1': '0/10-1/10', '2': '2/10-3/10', '3': '4/10',
      '4': '5/10', '5': '6/10', '6': '7/10-8/10', '7': '9/10-10/10',
      '8': '10/10', '9': '-', '/': '-'
      }
   return switcher.get(str0,'invalid cloud cover')

### wind direction conversion
def dir_convert(str0):
   switcher = {
      '00': '00','02': 'NNE','05': 'NE','07': 'ENE', '09': 'E',
      '11': 'ESE', '14': 'SE', '16': 'SSE', '18': 'S',
      '20': 'SSW', '23': 'SW', '25': 'WSW', '27': 'W',
      '29': 'WNW', '32': 'NW', '34': 'NNW', '36': 'N','': ''
      }
   return switcher.get(str0,"invalid wind direction")

## Function to get position of a sub string inside a big string
def pos(str1,str2):
   pos = -1
   if str1 in str2:
      for i in range(len(str2.split())):
         if str1 == str2.split()[i]:
            pos = i
   return pos
####
def get_record_synop(str0):
   "This function get typh data"
   #print (str.split())
   typh_len = len(str0.split())
   #print(typh_len)
   stations.append(str0.split()[0])
   cloud_cover.append(cloud_convert(str0.split()[2][0:1]))
   wind_dir.append(dir_convert(str0.split()[2][1:3]))
   wind_speed.append(str0.split()[2][3:5])
   t2m.append(float(str0.split()[3][2:5])/10)
   tdew.append(float(str0.split()[4][2:5])/10)
   
   ## get station pressure
   station_pres_exist = 0
   for i in range(typh_len):
      if (str0.split()[i][0:1] == '3') and\
         (4 < pos(str0.split()[i],str0) < 7):
         #print (str0.split()[i])
         station_pres_exist += 1
         if (float(str0.split()[i][1:5]) < 1000):
            station_pres.append(1000 + float(str0.split()[i][1:5])/10)
         elif(float(str0.split()[i][1:5]) > 1000):
            station_pres.append(float(str0.split()[i][1:5])/10)
   if (station_pres_exist == 0):
      station_pres.append('')
   
   ### get sea level pressure
   pres_exist = 0
   for i in range(typh_len):
      if (str0.split()[i][0:1] == '4') and\
         (8 > pos(str0.split()[i],str0) > 4):
         pres_exist += 1
         #print (str0.split()[i])
         if (float(str0.split()[i][1:5]) < 1000):
            slp.append(1000 + float(str0.split()[i][1:5])/10)
         elif(float(str0.split()[i][1:5]) > 1000):
            slp.append(float(str0.split()[i][1:5])/10)
   if (pres_exist == 0):
      slp.append('')
      
   ### get delta_p
   deltap_exist = 0
   for i in range(typh_len):
      if (str0.split()[i][0:1] == '5') and\
         (pos('333',str0) <= pos(str0.split()[i],str0) <= pos('333',str0) + 3):
         deltap_exist += 1
         print(str0.split()[i])
         if (int(str0.split()[i][0:2]) == 58):
            delta_p.append(float(str0.split()[i][2:5])/10)
         elif (int(str0.split()[i][0:2]) == 59):
            delta_p.append(0 - float(str0.split()[i][2:5])/10)
   if (deltap_exist == 0):
      delta_p.append('')

   ### get gust_wind
   ## in case of typh mistyping:
   ## case 1: 555 9****
   gust_exist = 0
   for i in range(typh_len):
      if (str0.split()[i][0:1] == '9') and\
         (pos(str0.split()[i],str0) == pos('555',str0) + 1):
         gust_exist += 1
         #print(str.split()[i])
         gust_dir.append(dir_convert(str0.split()[i][1:3]))
         gust_speed.append(str0.split()[i][3:5])
   ## case 2: 911** 915**
   for i in range(typh_len):
      if (str0.split()[i][0:3] == '915') and\
         (pos(str0.split()[i],str0) > pos('333',str0)):
         gust_exist += 1
         gust_dir.append(dir_convert(str0.split()[i][3:5]))
      if (str0.split()[i][0:3] == '911') and\
         (pos(str0.split()[i],str0) > pos('333',str0)):
         gust_speed.append(str0.split()[i][3:5])
   #print ('gust_exist = ',gust_exist)
   if (gust_exist == 0):
      gust_dir.append('')
      gust_speed.append('')

   ### get water level
   sea_lev_exist = 0
   for i in range(typh_len):
      if (str0.split()[i][0:1] == 'E'):
         sea_lev_exist += 1
         sea_lev.append(str0.split()[i][1:4])
         sea_status.append(str0.split()[i][4:5])
   if (sea_lev_exist == 0):
      sea_lev.append('')
      sea_status.append('')
   return
################
def get_data():
   global stations,cloud_cover,t2m,tdew,wind_speed,wind_dir,station_pres,\
          slp,delta_p,gust_speed,gust_dir,pmin,time_pmin,dir_max,wind_max,\
          time_wind_max,sea_lev,sea_status
   global obs_day, obs_hour, obs_min, typh_type
   import datetime
   record_temp = []
   with open(input_file,'r') as dt:
      lines = dt.read()
      record_temp = lines.split("\n")
      #print(record_temp)
      for i in range(record_temp.count('')):
         record_temp.remove('')
      typh_type = (record_temp[2].split()[0]).lower() #find obs typh type aaxx or typh
      print (typh_type)
      obs_time = record_temp[1].split()[2]     #get the obs time
      print (obs_time)
      obs_day =  obs_time[0:2]
      obs_hour = obs_time[2:4]
      obs_min = obs_time[4:6]
      dt.close()
   ###
   stations = []
   cloud_cover = []
   t2m = []
   tdew = []
   wind_speed = []
   wind_dir = []
   station_pres = []
   slp =[]
   delta_p = []
   gust_speed = []
   gust_dir =[]
   pmin = []
   time_pmin = []
   dir_max = []
   wind_max = []
   time_wind_max = []
   sea_lev = []
   sea_status = []
   ###
# print(obs_year, obs_month, obs_day, obs_hour, obs_min)
   record_typh = []
   for i in range(len(record_temp)):
      if (len(record_temp[i].split()) >= 5):
         record_typh.append(record_temp[i])
   
   for i in range(len(record_typh)):
      get_record_synop(record_typh[i])
   return 
