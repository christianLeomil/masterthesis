import pyomo.environ as pyo
import pandas as pd
import classes
import functions
import inspect
import textwrap
import warnings
import copy

warnings.filterwarnings("ignore", '.*')

# ---------------------------------------------------------------------------------------------------------------------
# region reading data

path_input = './input/'
path_output = './output/'
name_file = 'df_input.xlsx'

df_input_series = pd.read_excel(path_input +name_file, sheet_name= 'series')
df_input_other = pd.read_excel(path_input +name_file, sheet_name= 'other')
df_elements = pd.read_excel(path_input + name_file, sheet_name = 'elements')
[df_matrix, df_aux] = functions.matrix_creator(df_elements)
df_aux.to_excel(path_output + 'df_aux.xlsx',index = False)

# functions.write_excel(df_matrix,path_input)

df_matrix.to_excel(path_output + 'df_matrix.xlsx') 

input("Press Enter to continue...")

df_conect = pd.read_excel(path_input + name_file, sheet_name = 'conect',index_col = 0)
df_conect.index.name = None
[df_conect, list_expressions, list_con_variables, list_attr_classes] = functions.connection_creator(df_conect)
df_conect.to_excel(path_output + 'df_conect.xlsx')

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region abstract creating model

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating classes for elements selected from existing classes

#create one class of each element chosen
for i in df_aux.index:
    globals()[df_aux['element'].iloc[i]] = getattr(classes,df_aux['type'].iloc[i])

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region renaming variables from methods from created classes

#looping through created classes and renaming variables in equations
for i in df_aux.index:
    print('---------------------------------' + df_aux['element'].iloc[i])
    for j in dir(globals()[df_aux['element'].iloc[i]]):
        if j.startswith('__'):
            pass
        else:
            print(j)
            # getting original method from class
            original_method = getattr(globals()[df_aux['element'].iloc[i]],j)
            # getting equation from method
            source_code = inspect.getsource(original_method)
            # removing indentation from equation text
            source_code = textwrap.dedent(source_code)
            # Replace the parameter name
            modified_source_code = source_code.replace(df_aux['type'].iloc[i],df_aux['element'].iloc[i])
            print(modified_source_code)
            # Compile the modified source code
            compiled_code = compile(modified_source_code, "<string>", "exec")
            # Create a namespace dictionary for execution
            namespace = {}
            # Execute the compiled code in the namespace
            exec(compiled_code, namespace)
            # Get the modified method from the namespace
            modified_method = namespace[j]
            # Set the modified method as the new method_1
            setattr(globals()[df_aux['element'].iloc[i]], j , modified_method)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding parameters and variables to abstract model

#create parameters from classes:
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    myObj = globals()[element]
    myObj = myObj()
    if hasattr(myObj,'list_param'):
        list_classes_parameters = [s.replace(class_type, element) for s in myObj.list_param]
        for j,m in enumerate(list_classes_parameters):
            specifications = myObj.list_text_param[j]
            text = specifications
            # print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Param({text}))")

#create variables from classes:
for i in df_aux.index:
    element = df_aux['element'].iloc[i]
    class_type = df_aux['type'].iloc[i]
    myObj = globals()[element]
    myObj = myObj()
    if hasattr(myObj,'list_var'):
        list_classes_variables = [s.replace(class_type, element) for s in myObj.list_var]
        for j,m in enumerate(list_classes_variables):
            specifications = myObj.list_text_var[j]
            text = 'model.HOURS,' + specifications
            # print(m)
            # print(specifications)
            exec(f"model.add_component('{m}',pyo.Var({text}))")

# dynamically creating variables from connections
for i in list_con_variables:
    if i in ['P_to_demand1']:
        pass
    else:
        print(i)
        text = 'model.HOURS, within = pyo.NonNegativeReals'
        exec(f"model.add_component('{i}',pyo.Var({text}))")

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating sets

model.time_step = pyo.Param()
model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
model.P_to_demand1 = pyo.Param(model.HOURS) #time series with solar energy
model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region adding connection methods to created classes

# loopin throuhg list_of_expressions and giving classes their equations:
constraint_number = 1
for i,n in enumerate(list_attr_classes):

    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'Constraint_con_' + str(constraint_number)

    def method_wrapper(model,t,expr = list_expressions[i]):
        return dynamic_method(model,t,expr)
    setattr(globals()[n], method_name, method_wrapper)
    constraint_number += 1

for i in df_aux.index:
    print('------------------'+df_aux['element'].iloc[i])
    for j in dir(globals()[df_aux['element'].iloc[i]]):
        if not j.startswith('__'):
            print(j) 

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region create constraints from created classes:


constraint_num = 1
for i in df_aux.index:
    for j in dir(globals()[df_aux['element'].iloc[i]]):
        if not j.startswith('__'):
            method = getattr(globals()[df_aux['element'].iloc[i]],j)
            model.add_component('Constraint_class_'+ str(constraint_num), pyo.Constraint(model.HOURS, rule = method))
            constraint_num += 1


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region objective


#objective function
def objective_rule(model,t):
     return sum(model.E_buy[t] * model.costBuy[t] - model.E_sell[t] * model.costSell[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule = objective_rule,sense= pyo.minimize)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region reading data

#reading data
data = pyo.DataPortal()
data['HOURS'] = df_input_series['HOURS'].tolist()
data['P_solar'] = df_input_series.set_index('HOURS')['P_solar'].to_dict()
data['P_to_demand1'] = df_input_series.set_index('HOURS')['P_to_demand1'].to_dict()
data['costBuy'] = df_input_series.set_index('HOURS')['costBuy'].to_dict()
data['costSell'] = df_input_series.set_index('HOURS')['costSell'].to_dict()

data['pv1_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'pv1_eff', 'Value'].values[0]}
data['pv1_area'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'pv1_area', 'Value'].values[0]}

data['pv2_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'pv2_eff', 'Value'].values[0]}
data['pv2_area'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'pv2_area', 'Value'].values[0]}

data['bat1_E_max'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_E_max', 'Value'].values[0]}
data['bat1_starting_SOC'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_starting_SOC', 'Value'].values[0]}
data['bat1_c_rate_ch'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_c_rate_ch', 'Value'].values[0]}
data['bat1_c_rate_dis'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_c_rate_dis', 'Value'].values[0]}
data['bat1_ch_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_ch_eff', 'Value'].values[0]}
data['bat1_dis_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat1_dis_eff', 'Value'].values[0]}

# data['bat2_E_max'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_E_max', 'Value'].values[0]}
# data['bat2_starting_SOC'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_starting_SOC', 'Value'].values[0]}
# data['bat2_c_rate_ch'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_c_rate_ch', 'Value'].values[0]}
# data['bat2_c_rate_dis'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_c_rate_dis', 'Value'].values[0]}
# data['bat2_ch_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_ch_eff', 'Value'].values[0]}
# data['bat2_dis_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat2_dis_eff', 'Value'].values[0]}

data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region starting model

#generating instance
instance = model.create_instance(data)

#printing constriants to check (uncomment to see all contraints)
print("Constraint Expressions:")
for constraint in instance.component_objects(pyo.Constraint):
    for index in constraint:
        print(f"{constraint}[{index}]: {constraint[index].body}")

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

# Display the DataFrame with the variable values
df_variable_values.to_excel(path_output + 'variable_values.xlsx',index = False)

#endregion