import pandas as pd

def connection_creator(df_conect,df_aux,path_output):

    list_variables = []
    list_variables_aux = []

    for i in df_conect.index:
        for j in df_conect.columns:
            if df_conect.loc[i,j] != 0:
                list_variables.append('P_to_' + i)
                if i not in ['demand','net']:
                    list_variables_aux.append('t,' + df_aux[i].iloc[0])
                else:
                    list_variables_aux.append('t')
                break

    for i in df_conect.columns:
        for j in df_conect.index:
            if df_conect.loc[j,i] != 0:
                list_variables.append('P_from_' + i)
                if i not in ['demand','net']:
                    list_variables_aux.append('t,' + df_aux[i].iloc[0])
                else:
                    list_variables_aux.append('t')   
                break

    for i in df_conect.index:
        for j in df_conect.columns:
            if df_conect.loc[i,j] != 0:
                list_variables.append('P_' + j +  '_' + i)
                if (i not in ['demand','net']) and (j not in ['demand','net']):
                    list_variables_aux.append('t,'+df_aux[i].iloc[0]+','+df_aux[j].iloc[0])
                elif (i in ['demand','net']) and (j in ['demand','net']):
                    list_variables_aux.append('t')
                else:
                    if  i in ['demand','net']:
                        list_variables_aux.append('t,'+ df_aux[j].iloc[0])
                    else:
                        list_variables_aux.append('t,'+ df_aux[i].iloc[0])

    df_variables = pd.DataFrame({'variables': list_variables,
                                'aux':list_variables_aux})

    for i in df_conect.index:
        for j in df_conect.columns:
            if df_conect.loc[i,j] == 1:
                df_conect.loc[i,j] = 'P_' + j + '_' + i

    df_conect.columns  = ['P_from_' + i for i in df_conect.columns]
    df_conect.index = ['P_to_' + i for i in df_conect.index]

    df_variables.to_excel(path_output + 'df_variables.xlsx')
    df_conect.to_excel(path_output + 'df_conect.xlsx')
    
    return df_conect, df_variables

def exp_creator(df_conect,df_variables):

    list_expressions = []
    
    #looping through columns, building "energy to" expressions
    for i in df_conect.index:
        list_exp_partial = []
        for j,m in enumerate(df_conect.columns):
            if df_conect.loc[i,m] != 0:
                if list_exp_partial == []:
                    suffix = df_variables[df_variables['variables'] == i]['aux'].iloc[0]
                    list_exp_partial.append(i + '['+ suffix + ']' + ' == ')
                    suffix = df_variables[df_variables['variables'] == df_conect.loc[i,m]]['aux'].iloc[0]
                    list_exp_partial[-1] = list_exp_partial[-1] +  df_conect.loc[i,m] + '['+ suffix + ']'
                else:
                    suffix = df_variables[df_variables['variables'] == df_conect.loc[i,m]]['aux'].iloc[0]
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + ' + df_conect.loc[i,m] + '['+ suffix + ']'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1] 
            list_expressions.append(list_exp_partial)

    #looping through columns, building "energy from" expressions
    for i in df_conect.columns:
        list_exp_partial = []
        for j,m in enumerate(df_conect.index):
            if df_conect.loc[m,i] != 0:
                if list_exp_partial == []:
                    suffix = df_variables[df_variables['variables'] == i]['aux'].iloc[0]
                    list_exp_partial.append(i + '['+ suffix + ']' + ' == ')
                    suffix = df_variables[df_variables['variables'] == df_conect.loc[m,i]]['aux'].iloc[0]
                    list_exp_partial[-1] = list_exp_partial[-1] +  df_conect.loc[m,i] + '['+ suffix + ']'
                else:
                    suffix = df_variables[df_variables['variables'] == df_conect.loc[m,i]]['aux'].iloc[0]
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + ' + df_conect.loc[m,i] + '['+ suffix + ']'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1] 
            list_expressions.append(list_exp_partial)

    return list_expressions

path_output = './output/'
path_input = './input/'
name_file = 'df_input.xlsx'
df_conect = pd.read_excel(path_input + name_file,sheet_name='conect',index_col=0)
df_conect.index.name = None
df_aux = pd.read_excel(path_input + name_file,sheet_name='aux')

[df_conect, df_variables] = connection_creator(df_conect,df_aux,path_output)
list_expressions = exp_creator(df_conect,df_variables)
print(list_expressions)
