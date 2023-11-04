import pandas as pd

# Create a sample DataFrame
data = {'Column1': [1, 2, 30, 4],
        'Column2': [0, 2, 2, 4],
        'Column3': [2, 2, 3, 4]}
df = pd.DataFrame(data)

# Get all columns with names starting with "Column"
columns_to_compare = [col for col in df.columns if col.startswith('Column')]

# Check if these columns have the same values in each row
df['AllColumnsEqual'] = df[columns_to_compare].eq(df[columns_to_compare].iloc[:, 0], axis=0).all(axis=1)

# Display the DataFrame
print(df)

data = pd.DataFrame(data)
data.set_index('Column1',inplace = True)
print(data)

print(data.index[2])