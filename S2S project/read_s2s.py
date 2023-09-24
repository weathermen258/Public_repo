from netCDF4 import Dataset,num2date
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

### read stations file
stations = []
lats = []
lons = []
stations_bv = []
lats_bv = []
lons_bv = []

from openpyxl import Workbook
from openpyxl import load_workbook
wb = load_workbook("./nc_files/data.xlsx", data_only=True)
sheet1 = wb['data']
last_cell = len(sheet1['A'])
stations = []
lats= []
lons = []
for i in range(1,last_cell+1):
    cell0 = 'A'+str(i)
    cell1 = 'B'+str(i)
    cell2 = 'C'+str(i)
    stations.append(sheet1[cell0].value)
    lons.append(sheet1[cell1].value)
    lats.append(sheet1[cell2].value)


df_ec = pd.DataFrame()
df_ec['stations'] = stations
df_ec['lat'] = lats
df_ec['lon'] = lons

############################################################
filename = "./nc_files/raw_tp_m1_2019-12-31.nc"
ec_file=Dataset(filename,'r')
ec_lats = ec_file.variables['latitude'][:]
ec_lons = ec_file.variables['longitude'][:]
ec_xtimes = ec_file.variables['time'][:]
ec_times = []

date0 = datetime.strptime('1900-01-01 00:00:00.0','%Y-%m-%d %H:%M:%S.%f')
for time in ec_xtimes:
   time0 = (date0 + timedelta(hours=int(time))).strftime("'%Y-%m-%d %H:%M:%S")
   #print (time,'',time0)
   ec_times.append(time0)

ec_rain = ec_file.variables['tp']
#print (ec_times)
def get_index_ec(lats,lons,ec_lats,ec_lons):
   global rain_value,lat_diff,lon_diff,idx_lats, idx_lons
   idx_lats = []
   idx_lons = []
   for lat in lats:
      lat_diff = [abs(x - lat) for x in ec_lats]
      idx_lat = lat_diff.index(min(lat_diff))
      idx_lats.append(idx_lat)
   for lon in lons:
      lon_diff = [abs(x - lon) for x in ec_lons]
      idx_lon = lon_diff.index(min(lon_diff))
      idx_lons.append(idx_lon)
   return idx_lats,idx_lons
get_index_ec(lats,lons,ec_lats,ec_lons)
#print (idx_lats)
#print (idx_lons)
for i in range(0,len(ec_times),4):
   #print (i)
   rain_ec = []
   for k in range(len(idx_lats)):
      if (i>=4):
         rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]] - \
            ec_rain[i-4][idx_lats[k]][idx_lons[k]])
      else:
         rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]])
   df_ec[ec_times[i]] = rain_ec

df_ec.to_excel("./nc_files/ketqua_EC.xlsx",sheet_name='ECMWF')
