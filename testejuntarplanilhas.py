import os
import pandas as pd
import numpy

pasta = 'C:/Users/Leandro/Documents/python/Excel com Python/testejuntar'

df = []

for file in os.listdir[pasta]:
    if file.endswith('.xlsx'):
      print('Loading file {0}...'.format(file))
      df.append(pd.read_excel(os.path.join(pasta,file)))

print(len(df))

df_master = pd.concat(df, axis=0)
df_master.to_excel('C:/Users/Leandro/Documents/python/Excel com Python/testejuntar/P123.xlsx', index=False)