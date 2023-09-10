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

df_cfs = pd.DataFrame()
df_cfs['stations'] = stations
df_cfs['lat'] = lats
df_cfs['lon'] = lons

df_ec = pd.DataFrame()
df_ec['stations'] = stations
df_ec['lat'] = lats
df_ec['lon'] = lons
##############################################
filename = "./nc_files/prate.2020010100.daily.nc"
nc_file=Dataset(filename,'r')
nc_lats = nc_file.variables['latitude'][:]
nc_lons = nc_file.variables['longitude'][:]
xtimes = nc_file.variables['time'][:]
times = []
for time in xtimes:
   time0 = datetime.fromtimestamp(time).strftime("%Y%m%d")
   times.append(time0)
rain = nc_file.variables['prate']

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

for i in range(46):
   rain_value = []
   for k in range(len(idx_lats)):
      rain_value.append(rain[i][idx_lats[k]][idx_lons[k]])
   df_cfs[times[i]] = rain_value

df_cfs.to_excel("./nc_files/ketqua_CFS.xlsx",sheet_name='CFS')
############################################################
