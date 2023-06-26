import pandas as pd
from openpyxl import load_workbook

def matrix_creator(df_elements):
    list_elements = []
    list_type = []
    for i in df_elements.columns:
        for j in range(0, df_elements[i].iloc[0]):
            list_elements.append(i + str(j+1))
            list_type.append(i)
    df_matrix = pd.DataFrame(0, index = list_elements, columns = list_elements)
    df_aux = pd.DataFrame({'element': list_elements,
                           'type':list_type})

    return df_matrix , df_aux

def connection_creator(df_conect):

    #getting seuquence of classes that will receive methods with expressions later on
    # first through rows
    list_attr_classes = []
    for i in df_conect.index:
        if any(expr !=0 for expr in df_conect.loc[i]):
            list_attr_classes.append(i)
    # now through columns
    for i in df_conect.columns:
        if any(expr !=0 for expr in df_conect[i]):
            list_attr_classes.append(i)


    #building variable names inside connection matrix
    for i in df_conect.index: 
        for j in df_conect.columns:
            if df_conect.loc[i,j] != 0:
                df_conect.loc[i,j] = 'P_' + j + '_' + i
    
    #creating variable names derived from columns of conection matrix
    list_sub = ['P_from_' + i for i in df_conect.columns]
    df_conect.columns = list_sub

    #creating variable names derived from index of conection matrix
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

    list_con_variables = []
    for i in df_conect.index:
        for j in df_conect.columns:
            if df_conect.loc[i,j] != 0:
                list_con_variables.append(df_conect.loc[i,j])
                list_con_variables.append(i)
                list_con_variables.append(j)
    list_con_variables = list(set(list_con_variables))

    return df_conect, list_expressions, list_con_variables, list_attr_classes


def write_excel(df_matrix, path_input):
    with pd.ExcelWriter(path_input + 'df_input.xlsx',mode = 'a',engine = 'openpyxl', if_sheet_exists='replace') as writer:
        df_matrix.to_excel(writer,sheet_name = 'conect')