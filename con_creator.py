import pandas as pd

def connection_creator(df_conect,df_aux):

    #generate all variables:

    list_variables = []
    list_variables_aux = []

    #checkin destination
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
    
    return df_conect, df_variables
    
