import pandas as pd
import classes
# from openpyxl import load_workbook

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

    return  df_con_electric, df_con_thermal

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


def write_excel(df,path_input,name_sheet):
    with pd.ExcelWriter(path_input + 'df_input.xlsx',mode = 'a',engine = 'openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer,sheet_name = name_sheet)


def objective_constraint_creator(df_aux): # this function creates the constraints in order for the objective function to work
    list_buy_constraint = ['model.total_buy[t] == '] 
    list_sell_constraint = ['model.total_sell[t] == ']
    df_aux = df_aux[df_aux['type'] == 'net' ].reset_index(drop = True)
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        if i == 0:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' model.' + element + '_buy[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' model.' + element + '_sell[t]'
        else:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' + model.' + element + '_buy[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' + model.' + element + '_sell[t]'

    list_objective_constraints = list_buy_constraint + list_sell_constraint
    return list_objective_constraints