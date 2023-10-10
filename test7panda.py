import pandas as pd 

df = pd.read_csv('data2.csv')

df.fillna(130, inplace = True)

# replace null values in the "calories" columns with the number 130
df["calories"].fillna(130, inplace = True)

# mean is het gemiddelde van de kolom calories
x = df["calories"].mean()
df["calories"].fillna(x, inplace = True)

# median is de mediaan ofwel het middelste getal
y = df["calories"].median()
df["calories"].fillna(y, inplace = True) 

# mode is het gegeven wat het vaakst voorkomt
z = df["calories"].mode()[0]
df["calories"].fillna(z, inplace = True)