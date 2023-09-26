import pandas as pd
import classes
import sys
import importlib

path_input ='./input/'
path_output = './output'
name_file = 'df_input.xlsx'

def write_to_financial_model(df_variable_values, path_output, boolean):
    columns = list(df_variable_values.columns)
    investment_costs = [i for i in columns if 'inv_cost' in i]
    operational_costs = [i for i in columns if 'op_cost' in i]

    df_cost = df_variable_values[operational_costs]
    df_cost['total cost'] = df_cost.sum(axis = 1)

    df_investment = df_variable_values[investment_costs]
    df_investment['total investment'] = df_investment.sum(axis = 1)

    df_variable_values = df_variable_values[operational_costs + investment_costs]

    cash = 0
    list_cash_flow = []
    for i in df_cost['total cost'].index:
        cash = cash - df_cost['total cost'].iloc[i] - df_investment['total investment'].iloc[i]
        list_cash_flow.append(cash)

    df_cash_flow = pd.DataFrame({'cash flow':list_cash_flow})

    write_excel(df_cash_flow,path_output,'cash flow','df_financial_data.xlsx',boolean)
    write_excel(df_cost,path_output,'df_cost','df_financial_data.xlsx',boolean)
    write_excel(df_investment,path_output,'df_investment','df_financial_data.xlsx',boolean)
    write_excel(df_variable_values,path_output,'input data','df_financial_data.xlsx',boolean)

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
    
    list_electric_load = []
    list_electric_source = []
    list_thermal_load = []
    list_thermal_source = []    

    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        element_type = df_aux['type'].iloc[i]
        
        myClass = getattr(classes,element_type)
        component_type = getattr(myClass,'component_type')

        if component_type['electric_load'] == 'yes':
            list_electric_load.append(element)
        if component_type['electric_source'] == 'yes':
            list_electric_source.append(element)
        if component_type['thermal_load'] == 'yes':
            list_thermal_load.append(element)
        if component_type['thermal_source'] == 'yes':
            list_thermal_source.append(element)

    df_con_electric = pd.DataFrame(0,columns = list_electric_source, index = list_electric_load)
    df_con_thermal = pd.DataFrame(0,columns = list_thermal_source, index = list_thermal_load)

    return  df_con_electric, df_con_thermal, df_aux
        
def write_excel(df,path, name_sheet, name_file, boolean):
    with pd.ExcelWriter(path + name_file,mode = 'a',engine = 'openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer,sheet_name = name_sheet, index = boolean)

def connection_creator(df_con_electric, df_con_thermal):
    list_attr_classes = []
    list_attr_classes = list_attr_classes + df_con_electric.index.to_list()
    list_attr_classes = list_attr_classes + df_con_electric.columns.to_list()

    list_attr_classes = list_attr_classes + df_con_thermal.index.to_list()
    list_attr_classes = list_attr_classes + df_con_thermal.columns.to_list()

    #changing names of variables P_to_demand and Q_to_demand, values usually given as input, turning them into parameters
    list_attr_classes = [s.replace('P_to_demand','param_P_to_demand') for s in list_attr_classes]
    list_attr_classes = [s.replace('Q_to_demand','param_Q_to_demand') for s in list_attr_classes]

    #building variable names inside connection matrix ELECTRIC
    for i in df_con_electric.index: 
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                if i == j:
                    print('==========ERROR==========')
                    print('Energy cannot be transferred from ' + i + ' to ' + j)
                    sys.exit()
                else:
                    df_con_electric.loc[i,j] = 'P_' + j + '_' + i

    #building variable names inside connection matrix THERMAL
    for i in df_con_thermal.index: 
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                if i == j:
                    print('==========ERROR==========')
                    print('Energy cannot be transferred from ' + i + ' to ' + j)
                    sys.exit()
                else:
                    df_con_thermal.loc[i,j] = 'Q_' + j + '_' + i

    #creating variable names derived from columns of conection matrix
    list_sub = ['P_from_' + i for i in df_con_electric.columns]
    df_con_electric.columns = list_sub
    list_sub = ['Q_from_' + i for i in df_con_thermal.columns]
    df_con_thermal.columns = list_sub

    #creating variable names derived from index of conection matrix
    list_sub = ['P_to_' + i for i in df_con_electric.index]
    list_sub = [s.replace('P_to_demand','param_P_to_demand') for s in list_sub]
    list_sub = [s.replace('P_to_charging_station','param_P_to_charging_station') for s in list_sub]
    print('\n->->->->->-list_sub_electric')
    print(list_sub)
    df_con_electric.index = list_sub
    list_sub = ['Q_to_' + i for i in df_con_thermal.index]
    list_sub = [s.replace('Q_to_demand','param_Q_to_demand') for s in list_sub]
    print('\n->->->->->-list_sub_thermal')
    print(list_sub)
    df_con_thermal.index = list_sub

    # creating equations in the 'to' direction ELECTRIC
    list_expressions = []
    for i in df_con_electric.index:
        list_exp_partial = 'model.' + i + '[t] == 0'
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                list_exp_partial = list_exp_partial + ' + model.' + df_con_electric.loc[i,j] + '[t]'
        list_expressions.append(list_exp_partial)

    # creating equations in the 'from' direction ELECTRIC
    for i in df_con_electric.columns:
        list_exp_partial = 'model.' + i + '[t] == 0'
        for j in df_con_electric.index:
            if df_con_electric.loc[j,i] != 0:
                list_exp_partial = list_exp_partial + ' + model.' + df_con_electric.loc[j,i] + '[t]'
        list_expressions.append(list_exp_partial)

    # creating equation in the 'to' direction THERMAL
    for i in df_con_thermal.index:
        list_exp_partial = 'model.' + i + '[t] == 0'
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                list_exp_partial = list_exp_partial + ' + model.' + df_con_thermal.loc[i,j] + '[t]'
        list_expressions.append(list_exp_partial)

    # creating equation in the 'from' direction THERMAL
    for i in df_con_thermal.columns:
        list_exp_partial = 'model.' + i + '[t] == 0'
        for j in df_con_thermal.index:
            if df_con_thermal.loc[j,i] != 0:
                list_exp_partial = list_exp_partial + ' + model.' + df_con_thermal.loc[j,i] + '[t]'
        list_expressions.append(list_exp_partial)

    # creating list with variables that need to be created in the abstract model ELECTRIC
    list_con_variables = []
    list_con_variables = list_con_variables + df_con_electric.columns.to_list()
    list_con_variables = list_con_variables + df_con_electric.index.to_list()
    for i in df_con_electric.index:
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                list_con_variables.append(df_con_electric.loc[i,j])
    list_con_variables = list(set(list_con_variables))

    # creating list with variables that need to be created in the abstract model THERMAL
    list_con_variables = list_con_variables + df_con_thermal.columns.to_list()
    list_con_variables = list_con_variables + df_con_thermal.index.to_list()
    for i in df_con_thermal.index:
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                list_con_variables.append(df_con_thermal.loc[i,j])
    list_con_variables = list(set(list_con_variables))

    return df_con_thermal, df_con_electric, list_expressions, list_con_variables, list_attr_classes

def objective_constraint_creator(df_aux,time_span,receding_horizon): # this function creates the constraints for the objective function to work
    list_buy_constraint = ['model.total_buy[t] == '] 
    list_sell_constraint = ['model.total_sell[t] == ']
    df_temp = df_aux[df_aux['type'] == 'net' ].reset_index(drop = True)
    for i in df_temp.index:
        element = df_temp['element'].iloc[i]
        if i == 0:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' model.' + element + '_buy_electric[t]' + ' + model.' + element + '_buy_thermal[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' model.' + element + '_sell_electric[t]' + ' + model.' + element + '_sell_thermal[t]'
        else:
            list_buy_constraint[-1] = list_buy_constraint[-1] + ' + model.' + element + '_buy_electric[t]' + ' + model.' + element + '_buy_thermal[t]'
            list_sell_constraint[-1] = list_sell_constraint[-1] + ' + model.' + element + '_sell_electric[t]' + ' + model.' + element + '_sell_thermal[t]'
    
    list_objective_constraints = list_buy_constraint + list_sell_constraint


    list_operation_costs_total = []
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        element_type = df_aux['type'].iloc[i]
        method = getattr(classes,element_type)
        if hasattr(method,'constraint_operation_costs'):
            if list_operation_costs_total == [] :
                list_operation_costs_total.append('model.total_operation_cost[t] == ' + ' model.' + element + '_op_cost[t]')
            else:
                list_operation_costs_total[-1] = list_operation_costs_total[-1] + '+ model.'+ element + '_op_cost[t]'

    list_objective_constraints = list_objective_constraints + list_operation_costs_total

    list_emissions_constraint = []
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        element_type = df_aux['type'].iloc[i]
        method = getattr(classes,element_type)
        if hasattr(method,'constraint_emissions'):
            if list_emissions_constraint == [] :
                list_emissions_constraint.append('model.total_emissions[t] == ' + 'model. ' + element +'_emissions[t]')
            else:
                list_emissions_constraint[-1] = list_emissions_constraint[-1] + ' + model. ' + element +'_emissions[t]'
    
    # print('\nEmission Constriants')
    # print(list_emissions_constraint)

    list_objective_constraints = list_objective_constraints + list_emissions_constraint

    return list_objective_constraints

def organize_output_columns(df_variable_values,df_aux):
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

    return df_variable_values

def breaking_dataframe(df,horizon,saved_position):
    list_split = []
    for i in range(0,len(df),saved_position):
        # print('\n------------'+str(i))
        if i == 0:
            list_split.append(df.iloc[i : i + horizon])
        else:
            list_split.append(df.iloc[i-1 : i+horizon])
        # print(list_split[-1])
    return list_split

def save_variables_last_time_step(df_input_others,df_variables_last_time_step):
    for j in df_variables_last_time_step.index:
        if df_variables_last_time_step['Parameter'].iloc[j] in df_input_others['Parameter'].tolist():
            df_input_others.loc[df_input_others['Parameter']==df_variables_last_time_step['Parameter'].iloc[j], 'Value'] = df_variables_last_time_step['Value'].iloc[j]
        else:
            row_to_append = df_variables_last_time_step.loc[j]
            df_input_others = df_input_others.append(row_to_append, ignore_index = True)
            
    return df_input_others

