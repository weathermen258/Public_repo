from netCDF4 import Dataset
import pandas as pd
from datetime import datetime, timedelta
import sys


### read stations file
stations = []
lats = []
lons = []
stations_bv = []
lats_bv = []
lons_bv = []
#ini_date = sys.argv[1]
#ini_date = '2023-09-14'
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

df_cfs = pd.DataFrame()
df_cfs['stations'] = stations
df_cfs['lat'] = lats
df_cfs['lon'] = lons

##############################################
filename = "./cfs/2023-09-01.nc"
nc_file=Dataset(filename,'r')
nc_lats = nc_file.variables['latitude'][:]
nc_lons = nc_file.variables['longitude'][:]
xtimes = nc_file.variables['time'][:]
times = []
day0 = datetime.strptime('1970-01-01 00:00:00.0',"%Y-%m-%d %H:%M:%S.%f")
for i in range(46):
   time0 = (day0 + timedelta(seconds=xtimes[i])).strftime("%Y-%m-%d %H:%M:%S")
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

for i in range(len(times)):
   rain_value = []
   for k in range(len(idx_lats)):
      rain_value.append(rain[i][idx_lats[k]][idx_lons[k]])
   df_cfs[times[i]] = rain_value
df_cfs0 = df_cfs.drop(columns =['lat','lon'],axis=1)
df_cfs0.rename(columns={'stations':'Date'},inplace=True)
df_cfs0.set_index('Date',inplace=True)
df_T = df_cfs0.T
output_file = "./ketqua_CFS.xlsx"
temp_file = "./temp.csv"
with pd.ExcelWriter(output_file) as writer:
   df_cfs.to_excel(writer,sheet_name='CFS')
   df_T.to_excel(writer,sheet_name='CFS_T')
col_names = df_T.columns
print (col_names)
from mikeio import Dfs0, ItemInfo, EUMType, EUMUnit
for i in range(len(col_names)):
   df_col = df_T.iloc[:,i]
   df_col.to_csv(temp_file,index=True,index_label='Date')
   df_tv = pd.read_csv(temp_file,parse_dates=True,index_col='Date')
   name = df_tv.columns[0]
   df_tv.rename(columns={name:'Weighted'},inplace=True)
   if ("RAINFALL" in col_names[i]):
      dfs_file = './songca_cfs/'+col_names[i]+'.dfs0'
   elif ('X' in col_names[i]):
      dfs_file = './songca_cfs/'+col_names[i]+'.dfs0'
   else:
      dfs_file = './songca_cfs/'+col_names[i]+'_Rainfall.dfs0'
   df_tv.to_dfs0(dfs_file,items=[ItemInfo('Weighted',EUMType.Rainfall,EUMUnit.millimeter)])
print (df_tv)

#test_file = os.path.join(results_dir,'testfile.csv')
#dfs_file = os.path.join(results_dir,'kq_gsm_'+ times[0] +'.dfs0')
#df_hour_transpose.to_csv(test_file,index=True,index_label='Date')
#df_tv = pd.read_csv(test_file,parse_dates=True,index_col='Date')
############################################################
