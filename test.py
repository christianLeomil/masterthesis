import pandas as pd

# Create a sample DataFrame
data = {
    'Column1': [10, 20, 30],
    'Column2': [5, 15, 25],
    'Column3': [2, 4, 6]
}

df = pd.DataFrame(data)

# Calculate the sum of all other columns for each row
row_sums = df.drop('Column1', axis=1).sum(axis=1)

# Add the row sums as a new column in the DataFrame
df['RowSums'] = row_sums

print(df)
