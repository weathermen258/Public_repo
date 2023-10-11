import mikeio
import os
cwd = os.getcwd()
file = os.path.join(cwd,'BAN VE 83_Rainfall.dfs0')
print (file)
ds = mikeio.read(file)
print (ds)
df = ds.to_dataframe()
print (df)
