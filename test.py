# import pyomo.environ as pyo

# # Create an abstract model
# model = pyo.AbstractModel()

# # Define a set of time steps
# model.time = pyo.Set(initialize=[1, 2, 3, 4, 5, 6, 7])

# # Define a variable
# model.x = pyo.Var(model.time, within=pyo.NonNegativeReals, initialize=2)

# # Define an objective (just for demonstration purposes)
# def objective_rule(model):
#     return sum(model.x[t] for t in model.time)

# model.obj = pyo.Objective(rule=objective_rule, sense=pyo.minimize)

# # def constraint_rule(model, t):
# #     if t == 1:
# #         return pyo.Constraint.Skip  # Skip the constraint for t=1
# #     return model.x[t] >= 1.5 * model.x[t - 1]
# # model.constraintRule = pyo.Constraint(model.time, rule=constraint_rule)

# def constraint_rule(model, t):
#     if t == 1:
#         return model.x[t] >= 1.5 * model.x[t-1]
# model.constraintRule = pyo.Constraint(model.time, rule=constraint_rule)

# # Create an instance of the abstract model
# instance = model.create_instance()

# # Skip optimization for t=1 and manually set values
# timestep_to_skip = 1
# if timestep_to_skip in instance.time:
#     instance.x[timestep_to_skip].fix(5.0)  # Fix the variable value for t=1

# # Solve the model
# solver = pyo.SolverFactory('cplex')
# solver.solve(instance)

# # Display the results
# for t in instance.time:
#     print(f'x[{t}] = {pyo.value(instance.x[t])}')


import pandas as pd

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

horizon = 5
saved_position = 4

path_input = './input/'
path_output = './output/'

df = pd.read_excel(path_input + 'df_input.xlsx', sheet_name = 'series', nrows = 30)

for i,df in enumerate(breaking_dataframe(df,horizon,saved_position)):
    df.to_excel(path_output + '/teste/df_split' + str(i) + '.xlsx')