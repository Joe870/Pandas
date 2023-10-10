import pandas as pd

a = [1,7,2]

myvar = pd.Series(a)

print(myvar)
print(myvar[0])

b = [3,9,6]

myver = pd.Series(b, index = ["x","y","z"])

print(myver)
print(myver["x"])

calories = {"day1": 420, "day2": 380, "day3": 390}
myvor = pd.Series(calories)
print(myvor)

myvur = pd.Series(calories, index = ["day1", "day2"])
print(myvur)

data = {
    "calories": [420, 380, 390],
    "duration": [50, 40, 45]
}

myvers = pd.DataFrame(data)

print(myvers)