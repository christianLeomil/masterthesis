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

def connection_creator(df_conect):

    for i in df_conect.index: 
        for j in df_conect.columns:
            if df_conect.loc[i,j] != 0:
                df_conect.loc[i,j] = 'P_' + j + '_' + i
    
    list_sub = ['P_from_' + i for i in df_conect.columns]
    df_conect.columns = list_sub

    list_sub = ['P_to_' + i for i in df_conect.index]
    df_conect.index = list_sub

    # creating equation in the 'to' direction
    list_expressions = []
    for i in df_conect.index:
        list_exp_partial = []
        for j,m in enumerate(df_conect.columns):
            if df_conect.loc[i,m] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_conect.loc[i,m] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_conect.loc[i,m] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    # creating equation in the 'from' direction
    for i in df_conect.columns:
        list_exp_partial = []
        for j,m in enumerate(df_conect.index):
            if df_conect.loc[m,i] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_conect.loc[m,i] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_conect.loc[m,i] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    return df_conect,list_expressions

# path_input = './input/'
# path_output = './output/' 
# name_file = 'df_input.xlsx'

# df_conect = pd.read_excel(path_input + name_file, sheet_name = 'conect',index_col = 0)
# df_conect.index.name = None

# [df_conect,list_expressions] = connection_creator(df_conect)

# print(df_conect)
# print(list_expressions)