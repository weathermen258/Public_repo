import pandas as pd
df1 = pd.DataFrame([('microsoft inc',2),
('Google Inc.',4),
('amazon, Inc.',5)],columns = ('A','id'))

df2 = pd.DataFrame([('(01780-500-01)','237489', '- 342','API',   1),
('(409-6043-01)','234324', ' API','Other   ',2),
('23423423','API', 'NaN','NaN',     3),
('(001722-5e240-60)','NaN', 'NaN','Other',   4),
('(0012172-52411-60)','32423423','   NaN','Other',   4),
('29849032-29482390','API', '    Yes','     False',   5),
('329482030-23490-1','API', '    Yes','     False',   6)],
columns = ['B','C','D','E','id'])

df1  =df1.set_index('id')
df1.drop_duplicates(inplace=True)
df2  = df2.set_index('id')
df3  = df1.join(df2,how='outer')
print (df1)
print (df2)
print (df3)