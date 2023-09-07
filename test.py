import pyomo.environ as pyo

# Create an abstract model
model = pyo.AbstractModel()

# Define a set of time steps
model.time = pyo.Set(initialize=[1, 2, 3, 4, 5, 6, 7])

# Define a variable
model.x = pyo.Var(model.time, within=pyo.NonNegativeReals, initialize = 10)

# Define an objective (just for demonstration purposes)
def objective_rule(model):
    return sum(model.x[t] for t in model.time)

model.obj = pyo.Objective(rule=objective_rule, sense=pyo.minimize)

def constraint_rule(model,t):
    return model.x[t] >=  5
model.constraintRule = pyo.Constraint(model.time,rule = constraint_rule)

# Create an instance of the abstract model
instance = model.create_instance()

# Hardcode the value of 'x' at time step 5 after creating the instance
fixed_value = 10.0
variable_name = 'x'
text = 'instance.'+variable_name+'[5].fix(fixed_value)'
exec(text)
# instance.x[5].fix(fixed_value)

# Solve the model
solver = pyo.SolverFactory('cplex')
solver.solve(instance)

# Display the results
for t in instance.time:
    print(f'x[{t}] = {pyo.value(instance.x[t])}')


# import pandas as pd

# def breaking_dataframe(df,horizon,saved_position):
#     list_split = []
#     for i in range(0,len(df),saved_position):
#         # print('\n------------'+str(i))
#         if i == 0:
#             list_split.append(df.iloc[i:i + horizon])
#         else:
#             list_split.append(df.iloc[i-1:i+horizon])
#         # print(list_split[-1])
#     return list_split

# horizon = 20
# saved_position = 3
# path_input = './input/'
# path_output = './output/teste/'

# df = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'series', nrows= 30)

# list_split = breaking_dataframe(df,horizon,saved_position)

# for i,df in enumerate(list_split):
#     df.to_excel(path_output + 'df_split' + str(i) +'.xlsx')
