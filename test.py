import pandas as pd

path_output = './output/'
path_input = './input/'

df_final = pd.read_excel(path_output + 'df_final.xlsx')

list_columns = [s for s in df_final.columns if 'op_cost' in s]

df_costs = df_final[list_columns]

df_costs.to_excel(path_output + 'teste.xlsx')