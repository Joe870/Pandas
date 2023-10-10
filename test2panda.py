import pandas as pd 

data = {
    "calories": [420, 380, 390],
    "duration": [50, 40, 45]
}

df = pd.DataFrame(data)

print(df)
print(df.loc[0])
print(df.loc[[0,1]])

da = pd.DataFrame(data, index = ["day 1", "day 2", "day 3"])
print(da)

print(da.loc["day2"])