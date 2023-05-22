import pandas as pd

path_input = './masterthesis/input/'
path_output = './masterthesis/output/'

df_input_other = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'other')
valor = df_input_other.loc[df_input_other['Parameter'] == 'E_bat_max', 'Value'].values[0]
print(valor)
print(type(valor))



df_input_other.to_excel(path_output + 'test.xlsx')