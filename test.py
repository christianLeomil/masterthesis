import pyomo.environ as pyo

# Create an abstract model
model = pyo.AbstractModel()

# Define a set of time steps
model.time = pyo.Set(initialize=[1, 2, 3, 4, 5, 6, 7])

# Define a variable
model.x = pyo.Var(model.time, within=pyo.NonNegativeReals)

# Define an objective (just for demonstration purposes)
def objective_rule(model):
    return sum(model.x[t] for t in model.time)

model.obj = pyo.Objective(rule=objective_rule, sense=pyo.minimize)

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
