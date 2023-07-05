import pandas as pd
import classes

path_input ='./input/'
path_output = './output'
name_file = 'df_input.xlsx'

def aux_creator(df_elements):
    list_elements = []
    list_type = []
    for i in df_elements.index:
        for j in range(0,df_elements['# components'].loc[i]):
            list_type.append(i)
            name_element = i + str(j+1)
            list_elements.append(name_element)
    df_aux = pd.DataFrame({'element': list_elements,
                           'type':list_type})
    list_con_electric = []
    list_con_thermal = []
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        element_type = df_aux['type'].iloc[i]

        myClass = getattr(classes,element_type)()
        energy_type = getattr(myClass,'energy_type')
        if energy_type['electric'] == 'yes':
            list_con_electric.append(element)
        if energy_type['thermal'] == 'yes':
            list_con_thermal.append(element)
    df_con_electric = pd.DataFrame(0,columns = list_con_electric, index = list_con_electric)
    df_con_thermal = pd.DataFrame(0,columns = list_con_thermal, index = list_con_thermal)

    return  df_con_electric, df_con_thermal, df_aux

def write_excel(df,path_input,name_sheet):
    with pd.ExcelWriter(path_input + 'df_input.xlsx',mode = 'a',engine = 'openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer,sheet_name = name_sheet)

def connection_creator(df_con_electric,df_con_thermal):
    list_attr_classes = []

    # getting classes that will receive connecion constraints later on. INDEX, ELECTRIC
    for i in df_con_electric.index:
        if any(expr != 0 for expr in df_con_electric.loc[i]):
            list_attr_classes.append(i)
    # getting classes that will receive connecion constraints later on. COLUMNS, ELECTRIC
    for i in df_con_electric.columns:
        if any(expr != 0 for expr in df_con_electric[i]):
            list_attr_classes.append(i)

    # getting classes that will receive connecion constraints later on. INDEX, THERMAL
    for i in df_con_thermal.index:
        if any(expr != 0 for expr in df_con_thermal.loc[i]):
            list_attr_classes.append(i)
    # getting classes that will receive connecion constraints later on. COLUMNS, THERMAL
    for i in df_con_thermal.columns:
        if any(expr != 0 for expr in df_con_thermal[i]):
            list_attr_classes.append(i)


    #building variable names inside connection matrix ELECTRIC
    for i in df_con_electric.index: 
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                df_con_electric.loc[i,j] = 'P_' + j + '_' + i

    #building variable names inside connection matrix THERMAL
    for i in df_con_thermal.index: 
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                df_con_thermal.loc[i,j] = 'Q_' + j + '_' + i

    
    #creating variable names derived from columns of conection matrix
    list_sub = ['P_from_' + i for i in df_con_electric.columns]
    df_con_electric.columns = list_sub
    list_sub = ['Q_from_' + i for i in df_con_thermal.columns]
    df_con_thermal.columns = list_sub

    #creating variable names derived from index of conection matrix
    list_sub = ['P_to_' + i for i in df_con_electric.index]
    df_con_electric.index = list_sub
    list_sub = ['Q_to_' + i for i in df_con_thermal.index]
    df_con_thermal.index = list_sub


    # creating equation in the 'to' direction ELECTRIC
    list_expressions = []
    for i in df_con_electric.index:
        list_exp_partial = []
        for j,m in enumerate(df_con_electric.columns):
            if df_con_electric.loc[i,m] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_con_electric.loc[i,m] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_con_electric.loc[i,m] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    # creating equation in the 'from' direction ELECTRIC
    for i in df_con_electric.columns:
        list_exp_partial = []
        for j,m in enumerate(df_con_electric.index):
            if df_con_electric.loc[m,i] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_con_electric.loc[m,i] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_con_electric.loc[m,i] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    # creating equation in the 'to' direction THERMAL
    for i in df_con_thermal.index:
        list_exp_partial = []
        for j,m in enumerate(df_con_thermal.columns):
            if df_con_thermal.loc[i,m] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_con_thermal.loc[i,m] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_con_thermal.loc[i,m] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    # creating equation in the 'from' direction THERMAL
    for i in df_con_thermal.columns:
        list_exp_partial = []
        for j,m in enumerate(df_con_thermal.index):
            if df_con_thermal.loc[m,i] !=0:
                if list_exp_partial == []:
                    list_exp_partial.append('model.'+ i + '[t] == model.' + df_con_thermal.loc[m,i] + '[t]')
                else:
                    list_exp_partial[-1] = list_exp_partial[-1] + ' + model.' + df_con_thermal.loc[m,i] + '[t]'
        if list_exp_partial != []:
            list_exp_partial = list_exp_partial[-1]
            list_expressions.append(list_exp_partial)

    # creating list with variables that need to be created in the abstract model ELECTRIC
    list_con_variables = []
    for i in df_con_electric.index:
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                list_con_variables.append(df_con_electric.loc[i,j])
                list_con_variables.append(i)
                list_con_variables.append(j)
    list_con_variables = list(set(list_con_variables))

    # creating list with variables that need to be created in the abstract model THERMAL
    for i in df_con_thermal.index:
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                list_con_variables.append(df_con_thermal.loc[i,j])
                list_con_variables.append(i)
                list_con_variables.append(j)
    list_con_variables = list(set(list_con_variables))

    return df_con_thermal, df_con_electric, list_expressions, list_con_variables, list_attr_classes

def objective_constraint_creator(df_aux): # this function creates the constraints in order for the objective function to work
    list_buy_constraint = ['model.total_buy[t] == '] 
    list_sell_constraint = ['model.total_sell[t] == ']
    df_aux = df_aux[df_aux['type'] == 'net' ].reset_index(drop = True)
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        if i == 0:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' model.' + element + '_cost_buy_electric[t]' + ' + model.' + element + '_cost_buy_thermal[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' model.' + element + '_cost_sell_electric[t]' + ' + model.' + element + '_cost_sell_thermal[t]'
        else:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' model.' + element + '_cost_buy_electric[t]' + ' + model.' + element + '_cost_buy_thermal[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' + model.' + element + '_cost_sell_electric[t]' + ' + model.' + element + '_cost_sell_thermal[t]'

    list_objective_constraints = list_buy_constraint + list_sell_constraint
    return list_objective_constraints

# df_elements = pd.read_excel(path_input + name_file,index_col=0,sheet_name = 'elements test')
# df_elements.index.name = None
# # print(df_elements)

# [df_con_electric,df_con_thermal,df_aux] = aux_creator(df_elements)

# write_excel(df_con_electric,path_input,'conect_electric')
# write_excel(df_con_thermal,path_input,'conect_thermal')

# input('Press enter do continue...')

# df_con_electric = pd.read_excel(path_input + name_file, sheet_name = 'conect_electric',index_col=0)
# df_con_electric.index.name = None

# df_con_thermal = pd.read_excel(path_input + name_file, sheet_name = 'conect_thermal', index_col=0)
# df_con_thermal.index.name = None

# [df_con_thermal, df_con_electric, list_expressions, 
#  list_con_variables, list_attr_classes] = connection_creator(df_con_electric, df_con_thermal)

# list_objective_constraints = objective_constraint_creator(df_aux)

# print('\n#df_con_electric:')
# print(df_con_electric)
# print('\n#df_con_thermal:')
# print(df_con_thermal)
# print('\n#list_con_variables:')
# print(list_con_variables)
# print('\n#list_attr_classes:')
# print(list_attr_classes)
# print('\n#list_expressions:')
# print(list_expressions)
# print('\n#df_aux:')
# print(df_aux)
# print('\n#list_constraints_objective:')
# print(list_constraints_objective)