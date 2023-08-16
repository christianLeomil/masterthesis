import pyomo.environ as pyo
import pandas as pd

path_output = './output/'
# -------------------------------------------------------------------------------------------------------------
# region creating model

model = pyo.AbstractModel()
model.HOURS = pyo.Set()

model.cost_x = pyo.Param(model.HOURS)
model.cost_y = pyo.Param(model.HOURS)
model.demand = pyo.Param(model.HOURS)

model.quant_x = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
model.quant_y = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
model.quant_z = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
model.tranfer_z = pyo.Var(model.HOURS, within = pyo.Reals)

# endregion
# -------------------------------------------------------------------------------------------------------------
# region creating constraint

def objective_rule(model, t):
    return sum(model.quant_x[t] * model.cost_x[t] + model.quant_y[t] * model.cost_y[t] 
               + model.quant_z[t] * model.cost_y[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule=objective_rule, sense=pyo.minimize)

def constraint_1(model,t):
    if t <= 10:
        return model.quant_z[t] >= 50
    else:
        return model.quant_x[t] + model.quant_y[t] >= 100
model.constraint1 = pyo.Constraint(model.HOURS, rule = constraint_1)

def constraint_2(model, t):
    return model.quant_x[t] + model.quant_y[t] + model.quant_z[t] == model.demand[t]
model.constraint2 = pyo.Constraint(model.HOURS, rule=constraint_2)

def constraint_3(model,t):
    if t != 1:
        return model.quant_z[t] == model.quant_z[t-1] + model.tranfer_z[t]
    else:
        return model.quant_z[t] == 100
model.constraint_3 = pyo.Constraint(model.HOURS, rule = constraint_3)

# endregion
# -------------------------------------------------------------------------------------------------------------
# region added from receding horizon

list_hours = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
              16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,
              28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39,
              40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]

list_cost_x = [0.41, 0.44, 0.35, 0.4, 0.5, 0.33, 0.31,
               0.33, 0.39, 0.32, 0.41, 0.47, 0.45, 0.42,
               0.3, 0.37, 0.39, 0.3, 0.33, 0.41, 0.34,
               0.49, 0.36, 0.49, 0.33, 0.49, 0.41, 0.31,
               0.48, 0.47, 0.36, 0.36, 0.42, 0.45, 0.34,
               0.31, 0.37, 0.37, 0.42, 0.35, 0.5, 0.41,
               0.45, 0.37, 0.34, 0.42, 0.47, 0.43, 0.32,
               0.36]

list_cost_y = [0.46, 0.34, 0.44, 0.33, 0.4, 0.48, 0.38,
               0.33, 0.31, 0.46, 0.32, 0.38, 0.47, 0.35,
               0.35, 0.47, 0.38, 0.35, 0.46, 0.36, 0.48, 0.38,
               0.37, 0.31, 0.41, 0.31, 0.38, 0.45, 0.5, 0.37,
               0.4, 0.46, 0.44, 0.43, 0.5, 0.43, 0.43, 0.43,
               0.48, 0.36, 0.31, 0.47, 0.38, 0.32, 0.33, 0.44,
               0.38, 0.42, 0.45, 0.33]

list_demand = [808, 997, 896, 802, 955, 982, 964, 836, 955, 970, 987,
               868, 955, 853, 968, 942, 892, 985, 990, 977, 957, 975,
               925, 801, 949, 985, 825, 932, 908, 981, 833, 831, 946,
               995, 981, 803, 821, 822, 835, 936, 898, 954, 986, 826,
               904, 997, 807, 937, 964, 852]

df = pd.DataFrame({'HOURS': list_hours,
                   'cost_x': list_cost_x,
                   'cost_y': list_cost_y,
                   'demand': list_demand
                   })

data = pyo.DataPortal()
data['HOURS'] = df['HOURS'].tolist()
data['cost_x'] = df.set_index('HOURS')['cost_x'].to_dict()
data['cost_y'] = df.set_index('HOURS')['cost_y'].to_dict()
data['demand'] = df.set_index('HOURS')['demand'].to_dict()

# Define the horizon length and the number of time periods
horizon_length = 10
num_time_periods = len(list_hours) - horizon_length + 1

for t in range(num_time_periods):
    print(f"===== Time Period: {t} =====")

    # Create a subset of the hours for the current horizon
    horizon_hours = list_hours[t : t + horizon_length]

    # Create a dictionary for the subset of data
    subset_data = {
        'HOURS': horizon_hours,
        'cost_x': {hour: data['cost_x'][hour] for hour in horizon_hours},
        'cost_y': {hour: data['cost_y'][hour] for hour in horizon_hours},
        'demand': {hour: data['demand'][hour] for hour in horizon_hours}
    }

    # Create a DataPortal instance with the subset data
    subset_data_portal = pyo.DataPortal()
    subset_data_portal['HOURS'] = subset_data['HOURS']
    subset_data_portal['cost_x'] = subset_data['cost_x']
    subset_data_portal['cost_y'] = subset_data['cost_y']
    subset_data_portal['demand'] = subset_data['demand']

    # Create an instance with the subset data
    instance = model.create_instance(subset_data_portal)

    # Solve the instance
    optimizer = pyo.SolverFactory('cplex')
    results = optimizer.solve(instance)

    if results.solver.termination_condition != pyo.TerminationCondition.optimal:
        print(f"Solver terminated with status: {results.solver.termination_condition}")

    # Display the results
    instance.display()

    #printing excel
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
    for i in instance.HOURS:
        row = {'TimeStep': i}
        for var_name in variable_names:
            var_value = getattr(instance, var_name)
            row[var_name] = pyo.value(var_value[i])
        df_variable_values = df_variable_values.append(row, ignore_index=True)

    # Organize and export the DataFrame with the variable values
    df_variable_values.to_excel(path_output + 'test_'+str(t)+ '.xlsx',index = False)

print("Receding horizon optimization completed.")

# endregion