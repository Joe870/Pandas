import pandas as pd 

df = pd.read_csv('data2.csv')
df['Date'] = pd.to_datetime(df['Date'])

print(df.to_string())

df.loc[7, "duration"] = 45 

for x in df.index:
    if df.loc[x, "duration"] > 120:
        df.loc[x, "duratiom"] = 120

for x in df.index:
    if df.loc[x, "duration"] > 120: 
        df.drop(x, inplace = True)

print(df.duplicated())
df.drop_duplicates(inplace = True)