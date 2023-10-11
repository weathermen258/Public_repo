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
ini_date = sys.argv[1]
date00 = sys.argv[2]
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
##############################################
def read_cfs(ini_date):
   global df_cfs
   df_cfs = pd.DataFrame()
   df_cfs['stations'] = stations
   df_cfs['lat'] = lats
   df_cfs['lon'] = lons
   filename = "./cfs/"+ini_date+".nc"
   nc_file=Dataset(filename,'r')
   nc_lats = nc_file.variables['latitude'][:]
   nc_lons = nc_file.variables['longitude'][:]
   xtimes = nc_file.variables['time'][:]
   #print (xtimes)
   global times
   times = []
   day0 = datetime.strptime('1970-01-01 00:00:00.0',"%Y-%m-%d %H:%M:%S.%f")
   for i in range(len(xtimes)):
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
   for i in range(len(times)):
      rain_value = []
      for k in range(len(idx_lats)):
         rain_value.append(rain[i][idx_lats[k]][idx_lons[k]])
      df_cfs[times[i]] = rain_value
   df_cfs0 = df_cfs.drop(columns =['lat','lon'],axis=1)
   df_cfs0.rename(columns={'stations':'Date'},inplace=True)
   df_cfs0.set_index('Date',inplace=True)
   df_T = df_cfs0.T
   return (df_T)
df_T0 = read_cfs(ini_date=date00)
df_T1 = read_cfs(ini_date=ini_date)
#print (df_T0)
#print (df_T1)
df_T = pd.concat([df_T0, df_T1], axis=0)
df_T = df_T[~df_T.index.duplicated(keep='first')]
df_T = df_T.sort_index()
df_T = df_T.iloc[0:45]
print (df_T)
output_file = "./ketqua_CFS_sCa"+ini_date+".xlsx"
temp_file = "./temp.csv"
with pd.ExcelWriter(output_file) as writer:
   df_cfs.to_excel(writer,sheet_name='CFS')
   df_T.to_excel(writer,sheet_name='CFS_T')
col_names = df_T.columns
#print (df_T.iloc[0:3])
#df_3Ngay = pd.DataFrame(columns=df_T.columns, index=None)
#df_5Ngay = pd.DataFrame(columns=df_T.columns, index=None)
#df_7Ngay = pd.DataFrame(columns=df_T.columns, index=None)
from mikeio import ItemInfo
from mikeio import EUMType,EUMUnit
for j in [3,5,7]:
   df_new = pd.DataFrame(columns=df_T.columns, index=None)
   #print (df_new)
   for k in range(j-1,len(df_T),j):
      print (k)
      #print (df_T.iloc[(k-3):k].sum(axis=0))
      df_new.loc[times[k]] = df_T.iloc[(k-j+1):(k+1)].sum(axis=0)
      #df_new = pd.concat([df_new,df_T.loc[(k-3):k].sum(axis=0)])
   print (df_new)
   for i in range(len(col_names)):
      df_col = df_new.iloc[:,i]
      df_col.to_csv(temp_file,index=True,index_label='Date')
      df_tv = pd.read_csv(temp_file,parse_dates=True,index_col='Date')
      #name = df_tv.columns[0]
      #df_tv.rename(columns={name:'Weighted'},inplace=True)
      if ("RAINFALL" in col_names[i]):
         dfs_file = './songca_cfs/'+ini_date+"/"+str(j)+'/'+col_names[i]+'.dfs0'
      elif ('X' in col_names[i]):
         dfs_file = './songca_cfs/'+'/'+ini_date+"/"+str(j)+'/'+col_names[i]+'.dfs0'
      else:
         dfs_file = './songca_cfs/'+ini_date+"/"+str(j)+'/'+col_names[i]+'_Rainfall.dfs0'
      df_tv.to_dfs0(dfs_file,items=[ItemInfo('Weighted',EUMType.Rainfall,EUMUnit.millimeter)])
############################################################
