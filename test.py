import pandas as pd

path_output = './output/'
path_input = './input/'

df_final = pd.read_excel(path_output + 'df_final.xlsx')

list_columns = [s for s in df_final.columns if 'op_cost' in s]
df_costs = df_final[list_columns]

list_columns = [s for s in df_final.columns if 'inv_cost' in s]
df_inv = df_final[list_columns]

list_columns = [s for s in df_final.columns if 'emissions' in s]

df_costs.to_excel(path_output + 'df_costs.xlsx')
df_costs.to_excel(path_output + 'df_investments.xlsx')
df_costs.to_excel(path_output + 'df_emissions.xlsx')