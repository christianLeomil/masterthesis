import pandas as pd
import pyomo.environ as pyo

def connection_to(df_conect_to):

    list_energy_from = []
    list_energy_to_total = []
    for i in range(0, len(df_conect_to)):
        list_energy_to = []
        for j in df_conect_to.columns:
            if df_conect_to[j].iloc[i] == 1:
                list_energy_to.append(j)
        if list_energy_to != []:
            list_energy_to_total.append(list_energy_to)
            list_energy_from.append(df_conect_to.index[i])

    df_conect_to = pd.DataFrame({'from': list_energy_from, 'to': list_energy_to_total})
    print(df_conect_to)
    print('\n')

    df_sets = pd.DataFrame({'name':['pv','bat'],
                            'var':['n','m'],
                            'set':['model.PV','model.BAT']})

    list_expressions = []
    list_sets = []
    for i in df_conect_to.index:
        list_expression_conect =[]
        if df_conect_to['from'].iloc[i] in ['demand','net']:
            list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t] == ')
            list_sets.append('')
        else:
            list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t][' + df_sets['var'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0] + '] == ')
            list_sets.append(df_sets['set'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0])
        for index,j in enumerate(df_conect_to['to'].iloc[i]):
            if index != 0:
                list_expression_conect[-1] = list_expression_conect[-1] + ' +'
            if j not in ['demand','net']:
                list_expression_conect[-1] = list_expression_conect[-1] + ' sum(model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t][' + str(df_sets['var'][df_sets['name']==j].iloc[0]) +'] for ' + str(df_sets['var'][df_sets['name']==j].iloc[0]) + ' in ' + str(df_sets['set'][df_sets['name']==j].iloc[0]) + ')'
            else:
                list_expression_conect[-1] = list_expression_conect[-1] + ' model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t]'
        list_expressions.append(list_expression_conect)

    return list_expressions, list_sets


def connect_from():

    return None

path_input = './input/'
name_file = 'df_input.xlsx'
df_connect_to = pd.read_excel(path_input + name_file , sheet_name = 'connect_to',index_col = 0)
df_connect_to.index.name = None

[list_expressions,list_sets] = connection_to(df_connect_to)
print(list_expressions)
print(list_sets)
