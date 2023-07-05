import pandas as pd

path_output = './output/'
path_input = './input/'
name_file = 'variable_values.xlsx'

df_variable_values = pd.read_excel(path_output + name_file)
df_aux = pd.read_excel(path_output + 'df_aux.xlsx')

list_columns = df_variable_values.columns
list_temp = []
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    for j in list_columns:
        if '_'+ element +'_' in j:
            list_temp.append(j)
        elif element +'_' in j:
            list_temp.append(j)
        elif 'from_'+ element in j:
            list_temp.append(j)
        elif 'to_' + element in j:
            list_temp.append(j)

complimentary_list = list(set(list_columns) - set(list_temp))
list_temp = list_temp + complimentary_list
list_temp.remove('TimeStep')
list_temp.insert(0,'TimeStep')

df_variable_values = df_variable_values.reindex(columns = list_temp)
df_variable_values.to_excel(path_output + 'test.xlsx',index = False)