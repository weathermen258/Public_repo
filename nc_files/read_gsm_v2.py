from netCDF4 import Dataset
import pandas as pd
from datetime import datetime, timedelta
import sys

ini_date = sys.argv[1]
### read stations file

from openpyxl import Workbook
from openpyxl import load_workbook
def get_stations(filename):
   wb = load_workbook(filename, data_only=True)
   sheet1 = wb['data']
   last_cell = len(sheet1['A'])
   global stations, lats, lons
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
   return stations, lats, lons
filename0 = 'stations_banve.xlsx'
get_stations (filename0)
#print (stations)
df_gsm = pd.DataFrame()
df_gsm['stations'] = stations
df_gsm['lat'] = lats
df_gsm['lon'] = lons

##############################################
filename = 'GSM_'+str(ini_date)+'.nc'
nc_file=Dataset(filename,'r')
nc_lats = nc_file.variables['lat_0'][:]
nc_lons = nc_file.variables['lon_0'][:]
xtimes = nc_file.variables['forecast_time0'][:]
times = []
date0 = datetime.strptime(filename[4:14],'%Y%m%d%H')
for t in xtimes:
   time0 = (date0 + timedelta(hours=int(t))).strftime("%Y%m%d%H")
   times.append(time0)
   #print (t,time0)
rain = nc_file.variables['APCP_P8_L1_GLL0_acc']

def get_index(lats,lons,nc_lats,nc_lons):
   global rain_value,lat_diff,lon_diff,idx_lats, idx_lons
   idx_lats = []
   idx_lons = []
   for lat in lats:
      lat_diff = [abs(x - lat) for x in nc_lats]
      idx_lat = lat_diff.index(min(lat_diff))
      idx_lats.append(idx_lat)
   for lon in lons:
      lon_diff = [abs(x - lon) for x in nc_lons]
      idx_lon = lon_diff.index(min(lon_diff))
      idx_lons.append(idx_lon)
   return idx_lats,idx_lons

get_index(lats,lons,nc_lats,nc_lons)

#print (idx_lats)
#print (idx_lons)

for i in range(len(times)):
   rain_value = []
   for k in range(len(idx_lats)):
      if (i>=1):
         rain_value.append(rain[i][idx_lats[k]][idx_lons[k]]- rain[i-1][idx_lats[k]][idx_lons[k]])
      else:
         rain_value.append(rain[i][idx_lats[k]][idx_lons[k]])
   df_gsm[times[i]] = rain_value

df_hour = pd.DataFrame()
df_hour['stations'] = stations
df_hour['lat'] = lats
df_hour['lon'] = lons
for i in times:
   
   date_0 =  datetime.strptime(i,'%Y%m%d%H')
   date_00 = date_0.strftime('%Y-%m-%d %H:%M:%S')
   date_1 = (date_0 - timedelta(hours=1)).strftime('%Y-%m-%d %H:%M:%S')
   date_2 = (date_0 - timedelta(hours=2)).strftime('%Y-%m-%d %H:%M:%S')
   #print (date_00, date_1, date_2)
   df_hour[date_2] = df_gsm[i].apply(lambda x: pd.Series(x/3))
   df_hour[date_1] = df_gsm[i].apply(lambda x: pd.Series(x/3))
   df_hour[date_00] = df_gsm[i].apply(lambda x: pd.Series(x/3))


print ('################################### check 1')
df_hour = df_hour.drop(columns =['lat','lon'],axis=1)
df_hour.rename(columns={'stations':'Date'},inplace=True)
df_hour.set_index('Date',inplace=True)
#print (df_hour)
print ("################################# check 1.5")
df_hour_transpose = df_hour.T
df_hour_transpose.rename_axis('Date')
#print(df_hour_transpose)
print ('#################################### check 2')
#df_hour_transpose.set_index(pd.DatetimeIndex(df_hour_transpose['Date']), inplace=True)
#print (df_hour_transpose)
#df_hour_transpose.rename(columns={})
#df_hour_transpose = df_hour_transpose.drop(df_hour_transpose.index[[0,1,2,3]])
#print (df_hour_transpose['stations'])
import os
results_dir = "/home/cong/scripts/get_gsm/results"
output_file = os.path.join(results_dir,'kq_gsm_'+times[0]+'.xlsx')
with pd.ExcelWriter(output_file) as writer:
   df_gsm.to_excel(writer,sheet_name='3hours')
   df_hour.to_excel(writer,sheet_name='hours')
   df_hour_transpose.to_excel(writer,sheet_name='hours_ngang')
test_file = os.path.join(results_dir,'testfile.csv')
dfs_file = os.path.join(results_dir,'kq_gsm_'+ times[0] +'.dfs0')
df_hour_transpose.to_csv(test_file,index=True,index_label='Date')
df_tv = pd.read_csv(test_file,parse_dates=True,index_col='Date')
#print (df_tv)
print ('###################################### check 3')
tar_dir = "/home/phongdubao/NWP/GSM"
from shutil import copy2
excel_file = os.path.join(tar_dir,'kq_gsm_'+times[0]+'.xlsx')
copy2(output_file,excel_file)
print ('###################################### check 3.1')
print (pd.__version__)
import mikeio
print (mikeio.__version__)
print ("#####################################check 3.2")
from mikeio import Dfs0
print ("#####################################check 3.2")
df_tv.to_dfs0(dfs_file)
print ("###################################### check 4")
###############################################
dfs0_file = os.path.join(tar_dir,'kq_gsm_'+times[0]+'.dfs0')
copy2(dfs_file,dfs0_file)
############################################################
