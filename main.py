import pyomo.environ as pyo
import pandas as pd
import classes
import utils
import inspect
import textwrap
import warnings
import time

start_time = time.time()
warnings.filterwarnings("ignore", '.*') 

# ---------------------------------------------------------------------------------------------------------------------
# region reading data, creating connections variables, connections constraints, and objevtive constriants\

# paths and name of input filexz
path_input = './input/'
path_output = './output/'
name_file = 'df_input.xlsx'

# creating instance of class that contains infos of how optimization is going to be
control = classes.control(path_input,name_file)

#getting list of elements of energy system
df_elements = pd.read_excel(path_input + name_file, index_col=0, sheet_name = 'elements')
df_elements.index.name = None

#preparing data on elements of the energy system for the rest of the file
[df_con_electric, df_con_thermal, df_aux] = utils.aux_creator(df_elements)
df_aux.to_excel(path_output + 'df_aux.xlsx',index = False)

# writing dataframes on inputfile to input 
# utils.write_excel(df_con_electric,path_input,'conect_electric','df_input.xlsx', True)
# utils.write_excel(df_con_thermal,path_input,'conect_thermal','df_input.xlsx', True)
# input("\nPlease insert the connection between elements of energy system and press enter to continue...")

#reading inputs for the connections between elements of the energy system written in the input file
df_con_electric = pd.read_excel(path_input + name_file, sheet_name = 'conect_electric',index_col=0)
df_con_electric.index.name = None
df_con_thermal = pd.read_excel(path_input + name_file, sheet_name = 'conect_thermal', index_col=0)
df_con_thermal.index.name = None

#creating variables from connections exporting dataframes for check
[df_con_thermal, df_con_electric, list_expressions, 
 list_con_variables, list_attr_classes] = utils.connection_creator(df_con_electric, df_con_thermal)
df_con_electric.to_excel(path_output + 'df_con_electric.xlsx')
df_con_thermal.to_excel(path_output + 'df_con_thermal.xlsx')

# creating constriants that will turn into the objevtive functions
list_objective_constraints = utils.objective_constraint_creator(df_aux,control.time_span,control.receding_horizon)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating abstract model and creating temportal set

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()
model.starting_index = pyo.Param(initialize = 0)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating instances of selected elements of the energy system

for i in df_aux.index:
    # globals()[df_aux['element'].iloc[i]] = getattr(classes,df_aux['type'].iloc[i])(df_aux['element'].iloc[i], control.time_span - 1, control.receding_horizon)
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
                modified_source_code = source_code.replace(element_type,element)
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
# region RENAMING VARIABLES from list_var of elements selected for the energy system

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

if control.df.loc['size_optimization','value'] == 'yes':
    list_param_to_var = []
    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        method_value = list(vars(globals()[element]).keys())
        method_value = [s for s in method_value if 'param_' in s]
        # method_value = getattr(globals()[element],'list_param')
        list_param_to_var = list_param_to_var + method_value

    df_size_optimization = pd.DataFrame({'list_altered':list_param_to_var})
    df_size_optimization['choice'] = 0
    df_size_optimization['lower bound'] = 0
    df_size_optimization['upper bound'] = 0

    #writes to input file in order
    with pd.ExcelWriter(path_input + 'df_input.xlsx',mode = 'a', engine = 'openpyxl',if_sheet_exists= 'replace') as writer:
        df_size_optimization.to_excel(writer,sheet_name = 'parameters to variables',index = False)
    input('Please select parameters that are going to be optimized and press enter...')

    df_size_optimization = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'parameters to variables', index_col = 0)
    df_size_optimization.index.title = None
    df_size_optimization = df_size_optimization[df_size_optimization['choice'] == 1]

    list_altered_variables = df_size_optimization.index.tolist()
    list_upper_value = df_size_optimization['upper bound'].tolist()
    list_lower_value = df_size_optimization['lower bound'].tolist()
    
    
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


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region for loop for receiding horizon

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

model.time_step = pyo.Param(initialize = 1)

model.total_buy = pyo.Var(model.HOURS, within= pyo.NonNegativeReals) # variable contained in objective constraints
model.total_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals) # variable contained in objective constraints
model.total_operation_cost = pyo.Var(model.HOURS, within = pyo.Reals) # variable contained in objective constraints
model.total_emissions = pyo.Var(model.HOURS, within = pyo.NonNegativeReals) # variable contained in objective constraints

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

#adding total cost and total buy constraint to objective class
constraint_num = 1
objective_class = classes.objective('objective')
for i in list_objective_constraints:
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_objective_' + str(constraint_num) 

    def method_wrapper(model,t,expr = i):
        return dynamic_method(model,t,expr)
    setattr(objective_class , method_name, method_wrapper)
    constraint_num += 1

#checking content of class objective:
print('\n------------------objective')
for i in dir(objective_class):
    if not i.startswith('__'):
        original_method = getattr(objective_class,i)
        if callable(original_method):
            source_code = inspect.getsource(original_method)
            print(i)

#creating constraints from class objective
method = getattr(objective_class,'constraint_objective_1')  # constraint for calculating bought energy
model.add_component('Constraint_objective_buy',pyo.Constraint(model.HOURS, rule = method))
method = getattr(objective_class,'constraint_objective_2') # constraint for calculating sold energy
model.add_component('Constraint_objective_sell',pyo.Constraint(model.HOURS, rule = method))
method = getattr(objective_class,'constraint_objective_3') #constraint for opreational costs
model.add_component('Constraint_objective_operation',pyo.Constraint(model.HOURS, rule = method))
method = getattr(objective_class,'constraint_objective_4') #constraint for calculating emissions
model.add_component('Constraint_objective_emissions',pyo.Constraint(model.HOURS, rule = method))

#creating cost objective of abstract model
def cost_objective(model,t):
    return sum(model.total_buy[t] + model.total_operation_cost[t] -  model.total_sell[t] for t in model.HOURS)
if control.opt_objective == 'minimize':
    model.costObjective = pyo.Objective(rule = cost_objective, sense = pyo.minimize)
else:
    model.costObjective = pyo.Objective(rule = cost_objective, sense = pyo.maximize)

#creating emission objective of abstract model
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
# region reading series and parameters

df_input_series = pd.read_excel(path_input + name_file, sheet_name = 'series', nrows = control.time_span - 1)
df_input_other = pd.read_excel(path_input + name_file, sheet_name = 'other')

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region checking if simulation will have receiding horizon and then splitting input_series if it does.

if control.receding_horizon == 'yes':
    list_split = utils.breaking_dataframe(df_input_series,control.horizon,control.saved_position)
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
    # data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}
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
    # df_time_dependent_variable_values.to_excel(path_output +'/time dependant variables/'+ 'df_time_dependent_variable_values' + str(k) + '.xlsx',index = False)

    if control.receding_horizon == 'yes':
        if k == 0:
            df_final = pd.concat([df_final, df_time_dependent_variable_values.loc[0 : control.saved_position-1]], ignore_index = True)
        else:
            df_final = pd.concat([df_final, df_time_dependent_variable_values.loc[1 : control.saved_position]], ignore_index = True)
    else:
        df_final = df_time_dependent_variable_values.copy()

if control.size_optimization == 'yes':
    df_scalar_variable_values = pd.DataFrame([variable_values_scalar], columns = variable_names_scalar).T
    df_scalar_variable_values.columns = ['value']
    df_scalar_variable_values.to_excel(path_output + 'df_scalar_variable_values.xlsx')

df_final.to_excel(path_output + 'df_final.xlsx',index = False)
utils.financial_analysis(control)

#endregion 

# Record the end time
end_time = time.time()

# Calculate and print the elapsed time
elapsed_time = end_time - start_time
print(f"Elapsed time: {elapsed_time} seconds")