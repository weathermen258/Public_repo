from netCDF4 import Dataset,num2date
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
### read stations file
stations = []
lats = []
lons = []
stations_bv = []
lats_bv = []
lons_bv = []
ini_date = sys.argv[1]
from openpyxl import Workbook
from openpyxl import load_workbook
wb = load_workbook("./Toa_do_LV_sCa.xlsx", data_only=True)
sheet1 = wb.worksheets[0]
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
filename = "./ec/"+ini_date+".nc"
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
for i in range(0,len(ec_times)):
   rain_value = []
   for k in range(len(idx_lats)):
      if (i >=1):
         if (ec_rain[i][idx_lats[k]][idx_lons[k]] - ec_rain[i-1][idx_lats[k]][idx_lons[k]] >= 0):
            rain_value.append(ec_rain[i][idx_lats[k]][idx_lons[k]] - ec_rain[i-1][idx_lats[k]][idx_lons[k]])
         else:
            rain_value.append(0)
      else:
         rain_value.append(ec_rain[i][idx_lats[k]][idx_lons[k]])
   df_ec[ec_times[i]] = rain_value
df_cfs0 = df_ec.drop(columns =['lat','lon'],axis=1)
df_cfs0.rename(columns={'stations':'Date'},inplace=True)
df_cfs0.set_index('Date',inplace=True)
df_T = df_cfs0.T
output_file = "./ketqua_EC_sCa"+ini_date+".xlsx"
temp_file = "./temp.csv"
with pd.ExcelWriter(output_file) as writer:
   df_ec.to_excel(writer,sheet_name='EC')
   df_T.to_excel(writer,sheet_name='EC_T')
col_names = df_T.columns
#print (col_names)
from mikeio import Dfs0, ItemInfo, EUMType, EUMUnit
for i in range(len(col_names)):
   df_col = df_T.iloc[:,i]
   df_col.to_csv(temp_file,index=True,index_label='Date')
   df_tv = pd.read_csv(temp_file,parse_dates=True,index_col='Date')
   if ("RAINFALL" in col_names[i]):
      dfs_file = './songca_ec/'+ini_date+'/'+col_names[i]+'.dfs0'
   elif ('X' in col_names[i]):
      dfs_file = './songca_ec/'+ini_date+'/'+col_names[i]+'.dfs0'
   else:
      dfs_file = './songca_ec/'+ini_date+'/'+col_names[i]+'_Rainfall.dfs0'
   df_tv.to_dfs0(dfs_file,items=[ItemInfo('Weighted',EUMType.Rainfall,EUMUnit.millimeter)])
