from pyomo.environ import *

# Create an MILP model
model = ConcreteModel()

# Define decision variables
model.x1 = Var(within=Reals)
model.x2 = Var(within=Reals)
model.x3 = Var(within=Reals)
a = 2.0  # Set a constant value for 'a'

# Define binary variables for piecewise approximation
N = 10  # Number of pieces in the approximation
model.y = Var(range(N), within=Binary)

# Define variable bounds for piecewise approximation
x3_min = 0.001  # Minimum allowed value for x3
x3_max = 10.0   # Maximum allowed value for x3

# Define the objective (for demonstration purposes)
model.obj = Objective(expr=model.x1 + model.x2, sense=minimize)

# Define the piecewise linear approximation constraints
def piecewise_linearization_rule(model, i):
    x3_range = (x3_max - x3_min) / N
    x3_i_min = x3_min + i * x3_range
    x3_i_max = x3_min + (i + 1) * x3_range
    
    return model.x1 <= (model.x2 / (x3_i_min * a)) + (model.y[i] * (x3_i_max * a - x3_i_min * a))
model.piecewise_linearization_constraint = Constraint(range(N), rule=piecewise_linearization_rule)

# Solve the MILP
solver = SolverFactory('cplex')
solver.solve(model)

# Display the results
print("Solution:")
print("x1 =", value(model.x1))
print("x2 =", value(model.x2))
print("x3 =", value(model.x3))
print("y =", [value(model.y[i]) for i in range(N)])
