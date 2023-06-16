import pandas as pd
from openpyxl import load_workbook

def matrix_creator(df_elements):
    list_elements = []
    list_type = []
    for i in df_elements.columns:
        for j in range(0, df_elements[i].iloc[0]):
            list_elements.append(i + str(j+1))
            list_type.append(i)
    df_matrix = pd.DataFrame(0, index=list_elements, columns=list_elements)
    df_aux = pd.DataFrame({'element':list_elements,
                           'type':list_type})

    return df_matrix , df_aux

# path_input = './input/'
# path_output = './output/' 
# name_file = 'df_input.xlsx'

# first run this
# df_elements = pd.read_excel(path_input + name_file, sheet_name = 'elements')
# df_matrix = matrix_creator(df_elements, path_input, name_file)
# df_matrix.to_excel(path_output + 'df_matrix.xlsx')

# than copy and paste df_matrix to sheet 'conect' and run this
# df_conect = pd.read_excel(path_input + name_file,sheet_name='conect',index_col=0)
# df_conect.index.name = None