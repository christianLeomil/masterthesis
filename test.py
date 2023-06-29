import pandas as pd

# Create a sample DataFrame
data = {'Name': ['Alice', 'Bob', 'Charlie', 'Dave', 'Eve'],
        'Age': [25, 30, 35, 40, 45]}
df = pd.DataFrame(data)

# Filter rows based on a value in a column
filtered_df = df[df['Name'] == 'Alice']

# Reset the index
filtered_df = filtered_df.reset_index(drop=True)

# Print the filtered DataFrame with reset index
print(filtered_df)