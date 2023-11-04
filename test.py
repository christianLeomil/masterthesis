
import pandas as pd
import inspect
import classes
import sys
import numpy as np

# paths and name of input file
path_input = './input/'
path_output = './output/'
name_file = 'input.xlsx'

control = classes.control(path_input, name_file)

def write_excel(df,path, name_sheet, name_file, boolean):
    with pd.ExcelWriter(path + name_file,mode = 'a',engine = 'openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer,sheet_name = name_sheet, index = boolean)

def write_avaliable_elements_and_domain_names(control):
    class_names = [name for name, obj in inspect.getmembers(classes) if inspect.isclass(obj)]
    list_elements = []
    for i in class_names:
        if hasattr(getattr(classes,i),'domain_type'):
            if len(getattr(classes,i).domain_type['energy_domains']) <= control.number_energy_domains:
                list_elements.append(i)
            
    df_elements = pd.DataFrame({'class_type': list_elements,
                               '# components': [0]*len(list_elements)})
    write_excel(df_elements,control.path_input,'microgrid_components', 'input.xlsx', False)

    list_domains = []
    for i in range(control.number_energy_domains):
        list_domains.append('domain' + str(i + 1))
    df_domain = pd.DataFrame(data = [[0] * len(list_domains)], columns = list_domains)
    df_domain = df_domain.T
    df_domain.columns = ['domain_names']
    df_domain.index.name = None
    write_excel(df_domain,control.path_input, 'energy_domains_names', 'input.xlsx', True)

    return

def create_element_df_and_domain_selection_df(df_elements,df_domains):  
    #checking if all domains were named, otherwise, print error message
    if all(isinstance(s,str) for s in df_domains['domain_names'].tolist()) and all(s != ' ' for s in df_domains['domain_names'].tolist()):
        pass
    else:
        print('\n==========ERROR==========')
        print('All domains must be named with a string and cannot be empty values\n')
        sys.exit()

    # creating element df
    list_elements = []
    list_type = []
    for i in df_elements.index:
        for j in range(0,df_elements.loc[i,'# components']):
            list_type.append(i)
            name_element = i + str(j+1)
            list_elements.append(name_element)

    #checking if there is any demand in the inputs, otherwise, print error message
    if not any('demand' in s for s in list_elements):
        print('\n==========ERROR==========')
        print('Model must have a demand in order to run optimization\n')
        sys.exit()

    df_aux = pd.DataFrame({'element': list_elements,
                           'type':list_type})
    
    # creating element df
    list_component_domain = []
    list_component_source_domain = []
    list_component_load_domain = []
    for i in df_aux.index:
        # element = df_aux['element'].iloc[i]
        element_type = df_aux['type'].iloc[i]
        list_component_domain.append(getattr(classes,element_type).domain_type['energy_domains'])
        list_component_source_domain.append(getattr(classes,element_type).domain_type['source_domains'])
        list_component_load_domain.append(getattr(classes,element_type).domain_type['load_domains'])

    df_domain_selection = pd.DataFrame({'element':list_elements,
                                        'component_domains':list_component_domain,
                                        'component_source_domains':list_component_source_domain,
                                        'component_load_domains':list_component_load_domain,
                                        'optimization_domains':[df_domains['domain_names'].tolist()] * len(list_component_domain)})
    
    biggest_list = max(list_component_domain,key = len)
    for i in range(len(biggest_list)):
        list_action = []
        name_column = 'domain_choice' +str(i + 1)
        for j in range(len(df_domain_selection)):
            if len(list_component_domain[j]) >= (i + 1):
                list_action.append('insert here')
            else:
                list_action.append(0)
        df_domain_selection[name_column] = list_action
    write_excel(df_domain_selection,control.path_input, 'domain_selection', 'input.xlsx', False)

    return df_aux

def create_connection_revenue_and_stock_matrices(df_domain, df_domain_selection):
    #checking if given domain names match pre-defined names  
    list_columns = [s for s in df_domain_selection.columns if 'domain_choice' in s]
    for i in list_columns:
        for j in df_domain_selection.index:
            if df_domain_selection.loc[j,i] != 0:
                if df_domain_selection.loc[j,i] not in df_domain['domain_names'].tolist():
                    print('\n==========ERROR==========')
                    print(f"'{df_domain_selection.loc[j,i]}' is not one of the defined domain names. Please start again.")
                    sys.exit()

    #checking if element has the same domains inserted for different energy types.
    df_temp = df_domain_selection[list_columns].copy()
    for i in range(len(df_temp)):
        if df_temp.iloc[i].duplicated().any():
            print('\n==========ERROR==========')
            print(f"Row of the element {df_domain_selection.index[i]} in df_domain_selection has duplicated values. Please select only one domain for each type of energy of the component")
            sys.exit()

    #checking if elements have more domains defines than they are design to produce/consume
    df_temp = df_domain_selection[list_columns].copy()
    for i in range(len(df_temp)):
        if sum(1 for s in df_temp.iloc[i] if s != 0) > len(eval(df_domain_selection['component_domains'].iloc[i])):
            print('\n==========ERROR==========')
            print(f"Row of the element {df_domain_selection.index[i]} has more domains defined than what the component is designed to handle.")
            sys.exit()

    #getting inputs and seprating element into sources and loads of the selected domains
    list_source_final = []
    list_load_final = []
    for i in df_domain.index:
        domain = df_domain.loc[i,'domain_names']
        list_source_domain = []
        list_load_domain = []
        for j in df_domain_selection.index:
            list_component_domain = eval(df_domain_selection.loc[j,'component_domains'])
            list_component_source_domain = eval(df_domain_selection.loc[j,'component_source_domains'])
            list_component_load_domain = eval(df_domain_selection.loc[j,'component_load_domains'])
            for m,n in enumerate(list_columns):
                if df_domain_selection.loc[j,n] == domain:
                    if list_component_domain[m] in list_component_source_domain:
                        list_source_domain.append(j)
                    if list_component_domain[m] in list_component_load_domain:
                        list_load_domain.append(j) 

        list_source_final.append(list_source_domain)
        list_load_final.append(list_load_domain)

    #creating and writing connection tables for each domain
    for i in range(len(df_domain)):
        df_temp = pd.DataFrame(0,index = list_load_final[i], columns = list_source_final[i])
        name_sheet = 'connect_domain_' + df_domain['domain_names'].iloc[i]
        write_excel(df_temp, path_input, name_sheet,'input.xlsx',True)

    #creating and writing revenue tables for each domain
    for i in range(len(df_domain)):
        df_temp = pd.DataFrame(0,index = list_load_final[i], columns = list_source_final[i])
        name_sheet = 'revenue_domain_' + df_domain['domain_names'].iloc[i]
        write_excel(df_temp, path_input, name_sheet,'input.xlsx',True)

    #creating and writing stock tables for each domain
    for i in range(len(df_domain)):
        df_temp = pd.DataFrame(0,index = list_load_final[i], columns = list_source_final[i])
        if len(df_temp.filter(like = 'net', axis = 0)) > 0:
            df_temp = df_temp.filter(like = 'net', axis = 0)
            name_sheet = 'stock_domain_' + df_domain['domain_names'].iloc[i]
            write_excel(df_temp, path_input, name_sheet,'input.xlsx',True)

    return 

def create_connection_equations(df_domains):
    list_attr_classes = []
    list_expressions = []
    list_con_variables = []
    list_connection_matrices = []
    for i in df_domains.index:
        domain_name = df_domains.loc[i,'domain_names']
        sheet_name = 'connect_domain_' + domain_name
        df_connect = pd.read_excel(path_input + name_file, sheet_name = sheet_name,index_col=0)
        df_connect.index.name = None

        list_attr_classes = list_attr_classes + df_connect.index.to_list()
        list_attr_classes = list_attr_classes + df_connect.columns.to_list()

        #building variable names inside connection
        for j in df_connect.index: 
            for k in df_connect.columns:
                if df_connect.loc[j,k] != 0:
                    if j == k:
                        print('\n==========ERROR==========')
                        print('Energy cannot be transferred from ' + j + ' to ' + k +'\n')
                        sys.exit()
                    else:
                        df_connect.loc[j,k] = domain_name + '_' + k + '_' + j

        #creating variable names derived from columns of conection matrix
        list_sub = [domain_name + '_from_' + s for s in df_connect.columns]
        df_connect.columns = list_sub

        #creating variable names derived from index of conection matrix
        list_sub = [domain_name + '_to_' + s for s in df_connect.index]
        list_sub = [s.replace(domain_name + '_to_demand','param_P_to_demand') for s in list_sub]
        list_sub = [s.replace(domain_name + '_to_charging_station','param_P_to_charging_station') for s in list_sub]
        df_connect.index = list_sub

        df_connect.to_excel(path_output + 'df_connect_' + domain_name + '.xlsx')

        # creating equations in the 'to' direction of connection matrix
        for j in df_connect.index:
            list_exp_partial = 'model.' + j + '[t] == 0'
            for k in df_connect.columns:
                if df_connect.loc[j,k] != 0:
                    list_exp_partial = list_exp_partial + ' + model.' + df_connect.loc[j,k] + '[t]'
            list_expressions.append(list_exp_partial)
        
        # creating equations in the 'from' direction of connection matrix
        for j in df_connect.columns:
            list_exp_partial = 'model.' + j + '[t] == 0'
            for k in df_connect.index:
                if df_connect.loc[k,j] != 0:
                    list_exp_partial = list_exp_partial + ' + model.' + df_connect.loc[k,j] + '[t]'
            list_expressions.append(list_exp_partial)

        # creating list with variables that need to be created in the abstract model ELECTRIC
        list_con_variables = list_con_variables + df_connect.columns.to_list()
        list_con_variables = list_con_variables + df_connect.index.to_list()
        for j in df_connect.index:
            for k in df_connect.columns:
                if df_connect.loc[j,k] != 0:
                    list_con_variables.append(df_connect.loc[j,k])
        list_con_variables = list(set(list_con_variables))

        list_connection_matrices.append(df_connect)

    print('\n---------------This is the list_attr_classes---------------')
    print(list_attr_classes)
    print('\n---------------This is the list_expressions---------------')
    print(list_expressions)
    print('\n---------------This is the list_con_variables---------------')
    print(list_con_variables)

    return list_connection_matrices, list_expressions, list_con_variables, list_attr_classes

def create_revenue_equations(df_domains):
    list_attr_classes = []
    list_expressions = []
    list_con_variables = []
    list_connection_matrices = []
    for i in df_domains.index:
        domain_name = df_domains.loc[i,'domain_names']
        sheet_name = 'revenue_domain_' + domain_name
        df_connect = pd.read_excel(path_input + name_file, sheet_name = sheet_name,index_col=0)
        df_connect.index.name = None

    return     

write_avaliable_elements_and_domain_names(control)
print("\nPlease insert the number of each avaliable element in sheet 'microgrid_components' of the 'input.xlsx' file.")
print("Please also input the name of the chosen domains in the sheet 'energy_domains_names' of the input.xlsx file.")
input("After inserting values please press enter...")

df_elements = pd.read_excel(path_input + name_file, index_col=0, sheet_name = 'microgrid_components')
df_elements.index.name = None
df_domains = pd.read_excel(path_input + name_file, index_col = 0, sheet_name = 'energy_domains_names')
df_domains.index.name = None

df_aux = create_element_df_and_domain_selection_df(df_elements,df_domains)
print("\nPlease select the domain for each element in the sheet 'domain_selection' of the 'input.xlsx' file.")
print("In this table, fill the cells containing 'insert here' with desired value. Do not change cells containing 0.")
input("After inserting values please press enter...")

df_domain_selection = pd.read_excel(path_input + name_file, index_col = 0, sheet_name = 'domain_selection')
df_domain_selection.index.name = None

create_connection_revenue_and_stock_matrices(df_domains, df_domain_selection)
print("\nPlease insert the connection between elements for the selected domains in sheet 'connect_domain_' of the 'input.xlsx' file.")
print("In this table,define the connection between elements inserting an 'x' in the matrix where the connection exists.")
input("After inserting values please press enter...")

[list_connection_matrices, 
 list_expressions, 
 list_con_variables, 
 list_attr_classes] = create_connection_equations(df_domains)

create_revenue_equation(df_domains)