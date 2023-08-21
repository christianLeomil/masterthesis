import pyomo.environ as pyo
import pandas as pd
import classes
import functions
import inspect
import textwrap
import warnings

warnings.filterwarnings("ignore", '.*')

# ---------------------------------------------------------------------------------------------------------------------
# region reading data

path_input = './input/'
path_output = './output/'
name_file = 'df_input.xlsx'

control = getattr(classes,'control')(path_input)

df_input_series = pd.read_excel(path_input +name_file, sheet_name = 'series', nrows = 999)
df_input_other = pd.read_excel(path_input + name_file, sheet_name = 'other')

df_elements = pd.read_excel(path_input + name_file,index_col=0,sheet_name = 'elements')
df_elements.index.name = None

[df_con_electric, df_con_thermal,df_aux] = functions.aux_creator(df_elements)
df_aux.to_excel(path_output + 'df_aux.xlsx',index = False)

# functions.write_excel(df_con_electric,path_input,'conect_electric','df_input.xlsx', True)
# functions.write_excel(df_con_thermal,path_input,'conect_thermal','df_input.xlsx', True)
# input("Press Enter to continue...")

df_con_electric = pd.read_excel(path_input + name_file, sheet_name = 'conect_electric',index_col=0)
df_con_electric.index.name = None

df_con_thermal = pd.read_excel(path_input + name_file, sheet_name = 'conect_thermal', index_col=0)
df_con_thermal.index.name = None

[df_con_thermal, df_con_electric, list_expressions, 
 list_con_variables, list_attr_classes] = functions.connection_creator(df_con_electric, df_con_thermal)

df_con_electric.to_excel(path_output + 'df_con_electric.xlsx')
df_con_thermal.to_excel(path_output + 'df_con_thermal.xlsx')

list_objective_constraints = functions.objective_constraint_creator(df_aux)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating abstract model

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating classes for elements selected from existing classes

#create one instance of the class of each element chosen
for i in df_aux.index:
    globals()[df_aux['element'].iloc[i]] = getattr(classes,df_aux['type'].iloc[i])(df_aux['element'].iloc[i])

# # comented code below is for checking what the global variables in this program are.  
# def print_global_variables():
#     global_vars = globals()
#     for var_name, var_value in global_vars.items():
#         print(var_name, "=", var_value)
#         print('\n')

# print_global_variables()
# print('\n')

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region renaming variables from methods from created classes

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
# region checking if optimization should do also size optimization of components



# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding parameters and variables to abstract model

#add SERIES PARAMETERS from classes to abstract model:
print('\n------------------------series from classes') 
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    if hasattr(globals()[element],'list_series'):
        list_classes_series = [s.replace(class_type, element) for s in globals()[element].list_series]
        for j,m in enumerate(list_classes_series):
            specifications = globals()[element].list_text_series[j]
            text = specifications
            print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Param({text}))")


#add PARAMETERS from classes to abstract model:
print('\n------------------------paramaters from classes')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    if hasattr(globals()[element],'list_param'):
        list_classes_parameters = [s.replace(class_type, element) for s in globals()[element].list_param]
        for j,m in enumerate(list_classes_parameters):
            specifications = globals()[element].list_text_param[j]
            text = specifications
            print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Param({text}))")

#add VARIABLES from classes to abstract model:
print('\n------------------------variables from classes')
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    if hasattr(globals()[element],'list_var'):
        list_classes_variables = [s.replace(class_type, element) for s in globals()[element].list_var]
        for j,m in enumerate(list_classes_variables):
            specifications = globals()[element].list_text_var[j]
            text = 'model.HOURS,' + specifications
            print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Var({text}))")

# dynamically adding variables from connections to abstract model
print('\n------------------------variables from connections')
for i in list_con_variables:
    if not i.startswith('P_to_demand'):
        if not i.startswith('Q_to_demand'):
            if not i.startswith('P_to_charging_station'):
                print(i)
                text = 'model.HOURS, within = pyo.NonNegativeReals'
                exec(f"model.add_component('{i}',pyo.Var({text}))")

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating extra series and variables

model.time_step = pyo.Param()

model.total_buy = pyo.Var(model.HOURS, within= pyo.NonNegativeReals) # variable contained in objective constraints
model.total_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals) # variable contained in objective constraints
model.total_operation_cost = pyo.Var(model.HOURS, within = pyo.Reals) # variable contained in objective constraints
model.total_emissions = pyo.Var(model.HOURS, within = pyo.NonNegativeReals) # variable contained in objective constraints

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding connection methods to created classes

# loopin throuhg list_of_expressions and giving classes their equations:
constraint_number = 1
for i,n in enumerate(list_attr_classes):

    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'constraint_con_' + str(constraint_number)

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
# region create constraints from created classes:

constraint_num = 1
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    methods = inspect.getmembers(globals()[element],inspect.isfunction)
    for method_name,method in methods:
        if not method_name.startswith('__'):
            method = getattr(globals()[element],method_name)
            model.add_component('Constraint_class_'+ str(constraint_num), pyo.Constraint(model.HOURS, rule = method))
            constraint_num += 1

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region objective

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
print('------------------objective')
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

if control.opt_equation == 'cost objective':
    model.emissionObjective.deactivate()
else:
    model.costObjective.deactivate()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region reading data

#reading data
data = pyo.DataPortal()
data['HOURS'] = df_input_series['HOURS'].tolist()
data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}

#getting list with all needed PARAMETERS of created classes and reading data, or getting default values from classes
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    list_parameters =  getattr(globals()[element],'list_param')
    for j in list_parameters:
        name_param = j.replace(class_type,element)
        try:
            data[name_param] = {None:df_input_other.loc[df_input_other['Parameter'] == name_param, 'Value'].values[0]}
        except Exception:
            value = getattr(globals()[element],j)
            data[name_param] = {None: value}

#getting list with all needed SERIES of created classes and reading data, or getting default values
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    list_series =  getattr(globals()[element],'list_series')
    for j in list_series:
        name_series = j.replace(class_type,element)
        try:
            data[name_series] = df_input_series.set_index('HOURS')[name_series].to_dict()
        except Exception:
            list_par = getattr(globals()[element], j)
            if not isinstance(list_par,list):
                list_par = [list_par] * len(data['HOURS'])
                data[name_series] = {hour_value: getattr(globals()[element], j) for hour_value in data['HOURS']}
            else:
                df_test= pd.DataFrame({'HOURS': data['HOURS'],name_series : list_par})
                data[name_series] = df_test.set_index('HOURS')[name_series].to_dict()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region starting model

#generating instance
instance = model.create_instance(data)

#printing constriants to check (uncomment to see all contraints)
# print("Constraint Expressions:")
# for constraint in instance.component_objects(pyo.Constraint):
#     for index in constraint:
#         print(f"{constraint}[{index}]: {constraint[index].body}")

#solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
# instance.display()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region exporting results

variable_names =[]
for var_component in instance.component_objects(pyo.Var):
    for var in var_component.values():
        variable_names.append(var.name)

for i in range(len(variable_names)):
    # Find the index of '['
    index = variable_names[i].find('[') 
    # Remove the text after '[' including '['
    variable_names[i] = variable_names[i][:index]

variable_names = list(set(variable_names))

# Create an empty DataFrame to store the variable values
df_variable_values = pd.DataFrame(columns=['TimeStep'] + variable_names)

# Iterate over the time steps and extract the variable values
for t in instance.HOURS:
    row = {'TimeStep': t}
    for var_name in variable_names:
        var_value = getattr(instance, var_name)
        row[var_name] = pyo.value(var_value[t])
    df_variable_values = df_variable_values.append(row, ignore_index=True)

# Organize and export the DataFrame with the variable values
df_variable_values = functions.organize_output_columns(df_variable_values,df_aux)
df_variable_values.to_excel(path_output + 'variable_values.xlsx',index = False)

functions.write_to_financial_model(df_variable_values, path_output, False)

#endregion