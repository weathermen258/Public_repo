from netCDF4 import Dataset,num2date
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys,os

### read stations file
stations = []
lats = []
lons = []
stations_bv = []
lats_bv = []
lons_bv = []
#input_date = sys.argv[1]
input_date = '2019-06-01'
from openpyxl import Workbook
from openpyxl import load_workbook
wb = load_workbook("./data.xlsx", data_only=True)
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
df_ec3 = pd.DataFrame()
df_ec3['stations'] = stations
df_ec3['lat'] = lats
df_ec3['lon'] = lons
df_ec5 = pd.DataFrame()
df_ec5['stations'] = stations
df_ec5['lat'] = lats
df_ec5['lon'] = lons
df_ec7 = pd.DataFrame()
df_ec7['stations'] = stations
df_ec7['lat'] = lats
df_ec7['lon'] = lons
############################################################
filename = "./ec_raw/raw_tp_m1_"+input_date+".nc"
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
for i in range(0,len(ec_times),12):
   #print (i)
   rain_ec = []
   for k in range(len(idx_lats)):
      if (i>=12):
         if (ec_rain[i][idx_lats[k]][idx_lons[k]] - \
            ec_rain[i-12][idx_lats[k]][idx_lons[k]] >= 0):
            rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]] - \
               ec_rain[i-12][idx_lats[k]][idx_lons[k]])
         else:
            rain_ec.append(0)
      else:
         rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]])
   df_ec3[ec_times[i]] = rain_ec

df_ec3 = df_ec3.drop(columns =['lat','lon'],axis=1)
df_ec3.rename(columns={'stations':'Date'},inplace=True)
df_ec3.set_index('Date',inplace=True)
df_T_3 = df_ec3.T
for i in range(0,len(ec_times),20):
   #print (i)
   rain_ec = []
   for k in range(len(idx_lats)):
      if (i>=20):
         if (ec_rain[i][idx_lats[k]][idx_lons[k]] - \
            ec_rain[i-20][idx_lats[k]][idx_lons[k]] >= 0):
            rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]] - \
               ec_rain[i-20][idx_lats[k]][idx_lons[k]])
         else:
            rain_ec.append(0)
      else:
         rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]])
   df_ec5[ec_times[i]] = rain_ec

df_ec5 = df_ec5.drop(columns =['lat','lon'],axis=1)
df_ec5.rename(columns={'stations':'Date'},inplace=True)
df_ec5.set_index('Date',inplace=True)
df_T_5 = df_ec5.T
for i in range(0,len(ec_times),28):
   #print (i)
   rain_ec = []
   for k in range(len(idx_lats)):
      if (i>=28):
         if (ec_rain[i][idx_lats[k]][idx_lons[k]] - \
            ec_rain[i-28][idx_lats[k]][idx_lons[k]] >= 0):
            rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]] - \
               ec_rain[i-28][idx_lats[k]][idx_lons[k]])
         else:
            rain_ec.append(0)
      else:
         rain_ec.append(ec_rain[i][idx_lats[k]][idx_lons[k]])
   df_ec7[ec_times[i]] = rain_ec

df_ec7 = df_ec7.drop(columns =['lat','lon'],axis=1)
df_ec7.rename(columns={'stations':'Date'},inplace=True)
df_ec7.set_index('Date',inplace=True)
df_T_7 = df_ec7.T
output_file = "./ketqua_EC_"+input_date+".xlsx"
#temp_file = "./temp.csv"
#df_T.to_csv(temp_file,index=True,index_label='Date')
with pd.ExcelWriter(output_file) as writer:
   #df_ec.to_excel(writer,sheet_name='EC')
   df_T_3.to_excel(writer,sheet_name='EC_3Ngay')
   df_T_5.to_excel(writer,sheet_name='EC_5ngay')
   df_T_7.to_excel(writer,sheet_name='EC_7Ngay')