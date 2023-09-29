import mikeio
ds = mikeio.read('.\songma_ec\TRUNGSON_Rainfall.dfs0')
print (ds)
df = ds.to_dataframe()
print (df)
