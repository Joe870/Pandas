import pandas as pd 

df = pd.read_csv('data2.csv')
# Dropna verwijdert de rijden met lege waarden
# De dropna methode maakt een nieuwe dataframe aan en verandert de originele dataframe niet.
new_df = df.dropna()

print(new_df.to_string())

# Als je wil dat de originele dataframe wel verandert wordt dan gebruik je het inplace = True argument
df.dropna(inplace = True)
print(df.to_string)