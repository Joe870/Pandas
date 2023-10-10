import pandas as pd

# create a dictionary of data
data = {'name': ['Alice', 'Bob', 'Charlie', 'David'], 'age': [25, 30, 35, 40], 'country': ['USA', 'canada', 'Australia', 'UK']}

# create a pandas dataframe from the dictionary
df = pd.DataFrame(data)

# display the Dataframe
print(df)