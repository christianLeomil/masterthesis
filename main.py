import pyomo.environ as pyo
import pandas as pd
import classes
import utils
import inspect
import textwrap
import warnings
import time
import chime

chime.theme('mario')

start_time = time.time()
warnings.filterwarnings("ignore", '.*')

# ---------------------------------------------------------------------------------------------------------------------
# region reading data, creating list with connection variables, list with string of connection constraint, and list with objective constraints

# paths and name of input file
path_input = './input/'
path_output = './output/'
name_file = 'input.xlsx'

# creating instance of class that contains infos of how optimization is going to be
control = classes.control(path_input, name_file)

# utils.write_avaliable_elements_and_domain_names(control)
# print("\nPlease insert the number of each avaliable element in sheet 'microgrid_components' of the 'input.xlsx' file.")
# print("Please also input the name of the chosen domains in the sheet 'energy_domains_names' of the input.xlsx file.")
# input_time_1 = time.time()
# chime.info(sync = True)
# input("After inserting values please press enter...")
# input_time_2 = time.time()
# print('please wait...')

df_elements = pd.read_excel(path_input + name_file, index_col=0, sheet_name = 'microgrid_components')
df_elements.index.name = None
df_domains = pd.read_excel(path_input + name_file, index_col = 0, sheet_name = 'energy_domains_names')
df_domains.index.name = None

df_aux = pd.read_excel(control.path_output + 'df_aux.xlsx')

# df_aux = utils.create_element_df_and_domain_selection_df(df_elements,df_domains,control)
# print("\nPlease select the domain for each element in the sheet 'domain_selection' of the 'input.xlsx' file.")
# print("In this table, fill the cells containing 'insert here' with desired value. Do not change cells containing '0'.")
# input_time_3 = time.time()
# chime.info(sync = True)
# input("After inserting values please press enter...")
# input_time_4 = time.time()
# print('please wait...')

df_domain_selection = pd.read_excel(path_input + name_file, index_col = 0, sheet_name = 'domain_selection')
df_domain_selection.index.name = None

# utils.create_connection_revenue_and_stock_matrices(df_domains, df_domain_selection,control)
# print("\nPlease insert the connection between elements for the selected domains in sheet 'connect_domain_' of the 'input.xlsx' file.")
# print("Please also insert the revenue values for energy flows and energy flows that will be sold on the stock exchange in sheets 'revenue_domain and 'stock_domain_ in the 'input.xlsx' file")
# print("In this table,define the connection between elements inserting an 'x' in the matrix where the connection exists.")
# input_time_5 = time.time()
# chime.info(sync = True)
# input("After inserting values please press enter...")
# input_time_6 = time.time()
# print('please wait...')

[list_connection_matrices,
 list_expressions, 
 list_con_variables, 
 list_attr_classes] = utils.create_connection_equations(df_domains,control)

df_input_other = pd.read_excel(path_input + name_file, sheet_name = 'param_scalars')

[df_input_other, 
 list_expressions_rev, 
 list_variables_rev, 
 list_parameters_rev, 
 list_parameters_rev_value, 
 list_revenue_total, 
 list_correl_elements] = utils.create_revenue_and_stock_equations(df_domains,df_input_other,control)

[list_operation_costs_total,
 list_investment_costs_total, 
 list_emissions_total] = utils.objective_expression_creator(df_aux)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating abstract model and creating temporal set

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()
model.starting_index = pyo.Param(initialize = 0)
model.time_step = pyo.Param(initialize = 1)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating instances of selected elements of the energy system

for i in df_aux.index:
    globals()[df_aux['element'].iloc[i]] = getattr(classes,df_aux['type'].iloc[i])(df_aux['element'].iloc[i], control)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region RENAMING VARIABLES from METHODS of elements selected for the energy system

list_elements = df_aux['element']
list_types = df_aux['type']

for i,n in enumerate(list_elements):
    element = n
    element_type = list_types[i]
    methods = inspect.getmembers(globals()[element],inspect.ismethod)
    for method_name,method in methods:
        if method_name.startswith('constraint'):
                original_method = getattr(globals()[element],method_name)
                source_code = inspect.getsource(original_method)
                source_code = textwrap.dedent(source_code)
                #-----------replacing name of component in variables-----------#
                modified_source_code = source_code.replace(element_type,element)

                #-----------replacing name of connection variables with names of domains-----------#
                list_indexes = eval(df_domain_selection.loc[n,'component_domains'])
                for i,k in enumerate(list_indexes):
                    to_be_replaced = k + 'from_' + n
                    replacement = df_domain_selection.loc[n, 'domain_choice' + str(i + 1)] +'_from_' + n
                    modified_source_code = modified_source_code.replace(to_be_replaced, replacement)

                    to_be_replaced = k + 'to_' + n
                    replacement = df_domain_selection.loc[n, 'domain_choice' + str(i + 1)] +'_to_' + n
                    modified_source_code = modified_source_code.replace(to_be_replaced, replacement)
                #-----------replacing name of connection variables with names of domains-----------#

                # print(modified_source_code)
                compiled_code = compile(modified_source_code, "<string>", "exec")
                namespace = {}
                exec(compiled_code, namespace)
                modified_method = namespace[method_name]
                setattr(globals()[element], method_name , modified_method)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region RENAMING (self.)PARAMETERS from classes of elements selected for the energy system

for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    attribute_dict = list(vars(globals()[element]).keys())
    attribute_dict_altered  = [s.replace(class_type,element) for s in attribute_dict]

    for j in range(0,len(attribute_dict_altered)):
         if class_type in attribute_dict[j]:
            setattr(globals()[element],attribute_dict_altered[j],getattr(globals()[element],attribute_dict[j]))
            delattr(globals()[element],attribute_dict[j])

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region RENAMING VARIABLES from list_var of components selected for the energy system

# print('========================Renaming list_var parameters')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    list_var = globals()[element].list_var
    list_var = [s.replace(class_type,element) for s in list_var]
    setattr(globals()[element],'list_var',list_var)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region if size optimization yes, then choose parameters that are going to be optimized

if control.design_optimization == 'yes':
    list_param_to_var = []
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        method_value = list(vars(globals()[element]).keys())
        method_value = [s for s in method_value if 'param_' in s]
        # method_value = getattr(globals()[element],'list_param')
        list_param_to_var = list_param_to_var + method_value

    df_design_optimization = pd.DataFrame({'list_altered':list_param_to_var})
    df_design_optimization['choice'] = 0
    df_design_optimization['lower bound'] = 0
    df_design_optimization['upper bound'] = 0

    #writes to input file in order
    with pd.ExcelWriter(path_input + 'input.xlsx',mode = 'a', engine = 'openpyxl',if_sheet_exists= 'replace') as writer:
        df_design_optimization.to_excel(writer,sheet_name = 'parameters_to_variables',index = False)
    input_time_7 = time.time()
    chime.info(sync = True)
    input("\nPlease insert the parameters that are going to be optimized n the sheet 'param_to_variables' of the file 'input.xlsx'...")
    input_time_8 = time.time()
    print('please wait...')

    df_design_optimization = pd.read_excel(path_input + 'input.xlsx',sheet_name = 'parameters_to_variables', index_col = 0)
    df_design_optimization.index.title = None
    df_design_optimization = df_design_optimization[df_design_optimization['choice'] == 1]

    list_altered_variables = df_design_optimization.index.tolist()
    list_upper_value = df_design_optimization['upper bound'].tolist()
    list_lower_value = df_design_optimization['lower bound'].tolist()
    
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        list_original_param = list(vars(globals()[element]).keys())
        list_original_param = [s for s in list_original_param if 'param_' in s]
        
        list_altered_var = []
        list_text_altered_var = []
 
        for ind,j in enumerate(list_altered_variables):
            if j in list_original_param:
                list_altered_var.append(j)
                text ='within = pyo.NonNegativeReals, bounds = ('+str(list_lower_value[ind])+','+str(list_upper_value[ind])+')'
                list_text_altered_var.append(text)
                delattr(globals()[element],j)

        setattr(globals()[element],'list_altered_var', list_altered_var)
        setattr(globals()[element],'list_text_altered_var', list_text_altered_var)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding parameters from revenue connections to abstract model

for i,j,k in zip(list_correl_elements,list_parameters_rev,list_parameters_rev_value):
    setattr(globals()[i],j,k)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding parameters and variables to abstract model


#add PARAMETERS and SERIES from classes to abstract model:
print('\n------------------------paramaters from classes')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    list_parameters_series = list(vars(globals()[element]).keys())
    print('\n------------' + element)
    for j in list_parameters_series:
        if j.startswith('param'):
            print(j)
            if isinstance(getattr(globals()[element],j),list):
                exec(f"model.add_component('{j}',pyo.Param(model.HOURS))")
            else:
                exec(f"model.add_component('{j}',pyo.Param())")


#add VARIABLES from classes to abstract model:
print('\n------------------------variables from classes')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    if hasattr(globals()[element],'list_var'):
        list_classes_varibles = globals()[element].list_var
        # list_classes_variables = [s.replace(class_type, element) for s in globals()[element].list_var] 
        for j,m in enumerate(list_classes_varibles):
            specifications = globals()[element].list_text_var[j]
            text = 'model.HOURS,' + specifications
            print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Var({text}))")


#add ALTERED VARIABLES from classes to abstract model:
print('\n------------------------altered variables from classes')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    if hasattr(globals()[element],'list_altered_var'):
        list_classes_variables = globals()[element].list_altered_var
        # list_classes_variables = [s.replace(class_type, element) for s in globals()[element].list_altered_var]
        for j,m in enumerate(list_classes_variables):
            specifications = globals()[element].list_text_altered_var[j]
            text = specifications
            print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Var({text}))")


# dynamically adding VARIABLES FROM CONNECTIONS to abstract model
print('\n------------------------variables from connections')
for i in list_con_variables:
    if not i.startswith('param_'):
        if not i.startswith('P_to_charging_station'):
            print(i)
            text = 'model.HOURS, within = pyo.NonNegativeReals'
            exec(f"model.add_component('{i}',pyo.Var({text}))")


#dynamically adding VARIABLES FROM REVENUE CONNECTIONS to abstract model
print('\n------------------------variables from revenue connections')
for i in list_variables_rev:
    print(i)
    text = 'model.HOURS, within = pyo.NonNegativeReals'
    exec(f"model.add_component('{i}',pyo.Var({text}))")


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding connection constraints to instances of classes

# loopin throuhg list_of_expressions and giving each class the respective connection equation
constraint_number = 1
for i,n in enumerate(list_attr_classes):

    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_con_' + str(constraint_number)

    # print('\n-=-=-=-=-=-=-=' + list_expressions[i])
    # print(list_expressions[i])
    def method_wrapper(model,t,expr = list_expressions[i]):
        return dynamic_method(model,t,expr)
    setattr(globals()[n], method_name, method_wrapper)
    constraint_number += 1

for i in df_aux.index:
    print('------------------' + df_aux['element'].iloc[i])
    for j in dir(globals()[df_aux['element'].iloc[i]]):
        if not j.startswith('__'):
            if callable(getattr(globals()[df_aux['element'].iloc[i]],j)):
                print(j)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating extra series and variables

model.total_revenue = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.total_operation_costs = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.total_emissions = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.total_investment_costs = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region create constraints from created classes:

print('\n=============Creating constraints from classes=============')

constraint_num = 1
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    methods = inspect.getmembers(globals()[element],inspect.isfunction)
    print('------------' + element)
    for method_name,method in methods:
        if not method_name.startswith('__'):
            method = getattr(globals()[element],method_name)
            model.add_component('Constraint_class_'+ str(constraint_num), pyo.Constraint(model.HOURS, rule = method))
            print('-' + str(constraint_num) + '-' + method_name + ' ----> ' + 'Constraint_class_'+ str(constraint_num))
            # if 'Constraint_class_'+ str(constraint_num) == 'Constraint_class_43':
            #     print(inspect.getsource(method))
            constraint_num += 1

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating objective function and deactivating objectives

# adding total operation and investment costs, total revenue, and total emissions constraints to objective class
constraint_num = 1
objective_class = classes.objective('objective')

for i in list_revenue_total: # adding total revenue constraint to objective class
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    print(f"\n----This is the constraint_objective_{constraint_num}: {i}")
    constraint_num += 1


for i in list_operation_costs_total: # adding total operation costs constraint to objective class
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    print(f"\n----This is the constraint_objective_{constraint_num}: {i}")
    constraint_num += 1


for i in list_investment_costs_total: # adding total investment constraint to objective class
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    print(f"\n----This is the constraint_objective_{constraint_num}: {i}")
    constraint_num += 1


for i in list_emissions_total: # adding total emissions constraint to objective class
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    print(f"\n----This is the constraint_objective_{constraint_num}: {i}")
    constraint_num += 1


for i in list_expressions_rev: # adding connection revenue constraints to objective class (revenue and compensation constraints)
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    print(f"\n----This is the constraint_objective_{constraint_num}: {i}")
    constraint_num += 1
    

# checking content of class objective:
print('\n------------------objective')
for i in dir(objective_class):
    if not i.startswith('__'):
        original_method = getattr(objective_class,i)
        if callable(original_method):
            source_code = inspect.getsource(original_method)
            print(i)


# adding constraints from class OBJECTIVE to abstract model
print('\n=============Creating constraints from objecive class=============')
constraint_num = 1
element = objective_class
methods = inspect.getmembers(element, inspect.isfunction)
print('------------' + 'objective_class')
for method_name,method in methods:
    if not method_name.startswith('__'):
        method = getattr(element, method_name)
        model.add_component('Constraint_objective_'+ str(constraint_num), pyo.Constraint(model.HOURS, rule = method))
        print('-' + str(constraint_num) + '-' + method_name + ' ----> ' + 'Constraint_objective_'+ str(constraint_num))
        constraint_num += 1


# creating cost objective of abstract model
def cost_objective(model,t):
    return sum(model.total_operation_costs[t] + model.total_investment_costs[t] - model.total_revenue[t] for t in model.HOURS)
if control.opt_objective == 'minimize':
    model.costObjective = pyo.Objective(rule = cost_objective, sense = pyo.minimize)
else:
    model.costObjective = pyo.Objective(rule = cost_objective, sense = pyo.maximize)

# creating emission objective of abstract model
def emission_objective(model,t):
    return sum(model.total_emissions[t] for t in model.HOURS)
if control.opt_objective == 'minimize':
    model.emissionObjective = pyo.Objective(rule = emission_objective, sense = pyo.minimize)
else:
    model.emissionObjective = pyo.Objective(rule = emission_objective, sense = pyo.maximize)

if control.opt_equation == 'cost_objective':
    model.emissionObjective.deactivate()
    print('emission objective deactivated')
else:
    model.costObjective.deactivate()
    print('cost objective deactivated')

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region checking if simulation will have receiding horizon and then splitting input_series if it does.

df_input_series = pd.read_excel(path_input + name_file, sheet_name = 'param_series', nrows = control.time_span - 1)

if control.receding_horizon == 'yes':
    list_split = utils.breaking_dataframe(df_input_series, control.optimization_horizon, control.control_horizon)
else:
    list_split = []
    list_split.append(df_input_series)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding connection methods to created classes

df_final = pd.DataFrame()
for k,df in enumerate(list_split):
    print('K iteraction is ' + str(k))
    # df.to_excel(path_output +'teste/df_split'+str(k)+'.xlsx')

    if k != 0:
        last_time_step_index = df['HOURS'].iloc[0]
        print('--- the last time_step_is: '+ str(last_time_step_index))
        
        # model.starting_index = pyo.Param(initialize = last_time_step_index)
        df_input_other.loc[df_input_other['Parameter'] == 'starting_index','Value'] = last_time_step_index
        list_columns = df_time_dependent_variable_values.columns
        list_values = []

        for i in list_columns:
            list_values.append(df_time_dependent_variable_values.loc[df_time_dependent_variable_values['TimeStep'] == last_time_step_index,i].values[0])

        list_columns = ['param_' + s + '_starting_index' for s in list_columns]
        df_variables_last_time_step = pd.DataFrame({'Parameter': list_columns,
                                                    'Value': list_values})
        
        # df_variables_last_time_step.to_excel(path_output + '/teste/df_variables_last_time_step'+str(k)+'.xlsx',index = False) 
        df_input_other = utils.save_variables_last_time_step(df_input_other,df_variables_last_time_step)

    # endregion
    # ---------------------------------------------------------------------------------------------------------------------
    # region reading data for parameters and series. If no input is given default value is set

    #reading data
    data = pyo.DataPortal()
    data['HOURS'] = df['HOURS'].tolist()
    # data['starting_index'] = {None:control.starting_index}
    data['starting_index'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'starting_index', 'Value'].values[0]}

    #getting list with all needed PARAMETERS and SERIES of created classes and reading data, or getting default values from classes
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        class_type = df_aux['type'].iloc[i]
        list_parameters_series = list(vars(globals()[element]).keys())
        list_parameters_series = [s for s in list_parameters_series if 'param_' in s]
        # print('\n>>>>>>>>>>>>'+element)
        # print(list_parameters_series)
        for j in list_parameters_series:
            if isinstance(getattr(globals()[element],j),list):
                try:
                    data[j] = df.set_index('HOURS')[j].to_dict()
                except Exception:
                    df_temp = pd.DataFrame({'HOURS': data['HOURS'],j : getattr(globals()[element],j)})
                    data[j] = df_temp.set_index('HOURS')[j].to_dict()
            else:
                try:
                    data[j] = {None:df_input_other.loc[df_input_other['Parameter'] == j, 'Value'].values[0]}
                except Exception:
                    value = getattr(globals()[element],j)
                    data[j] = {None: value}

    # endregion
    # ---------------------------------------------------------------------------------------------------------------------
    # region starting model

    #generating instance
    instance = model.create_instance(data)

    # printing constriants to check (uncomment to see all contraints)
    # print("Constraint Expressions:")
    # for constraint in instance.component_objects(pyo.Constraint):
    #     for index in constraint:
    #         print(f"{constraint}[{index}]: {constraint[index].body}")


    # endregion
    # ---------------------------------------------------------------------------------------------------------------------
    # region solving model

    optimizer = pyo.SolverFactory('cplex')
    results = optimizer.solve(instance)

    # # Displaying the results
    # instance.pprint()
    # instance.display()

    # endregion
    # ---------------------------------------------------------------------------------------------------------------------
    # region exporting results

    # separating time dependend and time independent variables to be exported, only important when receding horizon is turned off
    variable_names_time_dependent = []
    variable_names_scalar = []
    variable_values_scalar = []
    for var_component in instance.component_objects(pyo.Var):
        if len(var_component) == 1:
            variable_names_scalar.append(var_component.name)
            variable_values_scalar.append(pyo.value(getattr(instance,var_component.name)))
        else:
            for var in var_component.values():
                variable_names_time_dependent.append(var.name)
    for i in range(len(variable_names_time_dependent)):
        # Find the index of '['
        index = variable_names_time_dependent[i].find('[') 
        # Remove the text after '[' including '['
        variable_names_time_dependent[i] = variable_names_time_dependent[i][:index]
    variable_names_time_dependent = list(set(variable_names_time_dependent))

    # Create an empty DataFrame to store the variable values
    df_time_dependent_variable_values = pd.DataFrame(columns=['TimeStep'] + variable_names_time_dependent)

    # Iterate over the time steps and extract the variable values
    for t in instance.HOURS:
        row = {'TimeStep': t}
        for var_name in variable_names_time_dependent:
            var_value = getattr(instance, var_name)
            row[var_name] = pyo.value(var_value[t])
        df_time_dependent_variable_values = df_time_dependent_variable_values.append(row, ignore_index = True)

    # Organize and export the DataFrame with the variable values
    df_time_dependent_variable_values = utils.organize_output_columns(df_time_dependent_variable_values,df_aux)

    if control.receding_horizon == 'yes':
        if k == 0:
            df_final = pd.concat([df_final, df_time_dependent_variable_values.loc[0 : control.control_horizon-1]], ignore_index = True)
        else:
            df_final = pd.concat([df_final, df_time_dependent_variable_values.loc[1 : control.control_horizon]], ignore_index = True)
    else:
        df_final = df_time_dependent_variable_values.copy()

if control.design_optimization == 'yes':
    df_scalar_variable_values = pd.DataFrame([variable_values_scalar], columns = variable_names_scalar).T
    df_scalar_variable_values.columns = ['value']
    df_scalar_variable_values.to_excel(path_output + 'df_scalar_variables.xlsx')

# printing scalar parameters for abstract model
param_names_scalar = []
param_values_scalar = []
for param_component in instance.component_objects(pyo.Param):
    if len(param_component) == 1:
        param_names_scalar.append(param_component.name)
        param_values_scalar.append(pyo.value(getattr(instance,param_component.name)))

df_scalar_param = pd.DataFrame([param_values_scalar], columns = param_names_scalar).T
df_scalar_param.columns = ['value']
df_scalar_param.to_excel(path_output + 'df_scalar_param.xlsx')

df_final.to_excel(path_output + 'df_final.xlsx',index = False)
utils.financial_analysis(control)
utils.emissions_analysis(control)
utils.charts_generator(control,df_aux,df_domains)

#endregion 

# Record the end time
end_time = time.time()

# Calculate and print the elapsed time
# if control.design_optimization == 'yes':
#     elapsed_time = end_time - (input_time_2 - input_time_1) - (input_time_4 - input_time_3) - (input_time_6 - input_time_5) - (input_time_8 - input_time_7) -start_time
# else:
#     elapsed_time = end_time - (input_time_2 - input_time_1) - (input_time_4 - input_time_3) - (input_time_6 - input_time_5) -start_time

elapsed_time = end_time - start_time

print(f"Elapsed time: {elapsed_time} seconds")
chime.success(sync=True)