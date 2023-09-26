from pyomo.environ import *

# Create a Concrete Pyomo Model
model = ConcreteModel()

# Define Variables
model.x = Var(within = Reals)
model.y = Var(within = Reals)
model.w = Var(within = Reals)

# # Define the Nonlinear Constraint to Linearize (e.g., xy <= z)
# def nonlinear_constraint_rule(model):
#     return model.z <= model.x * model.y
# model.nonlinear_constraint = Constraint(rule=nonlinear_constraint_rule)

# Create auxiliary variables for the McCormick envelopes
model.mccormick_x_lower = Param(initialize = 0)
model.mccormick_x_upper = Param(initialize = 10)
model.mccormick_y_lower = Param(initialize = 0)
model.mccormick_y_upper = Param(initialize = 3)


# Linearize the bilinear term xy <= z using McCormick envelopes
def mccormick_rule1(model):
    return model.x >= model.mccormick_x_lower
model.mccormick_rule_1 = Constraint(rule = mccormick_rule1)

def mccormick_rule2(model):
    return model.x <= model.mccormick_x_upper
model.mccormick_rule_2 = Constraint(rule = mccormick_rule2)

def mccormick_rule3(model):
    return model.y >= model.mccormick_y_lower
model.mccormick_rule_3 = Constraint(rule = mccormick_rule3)

def mccormick_rule4(model):
    return model.y <= model.mccormick_y_upper
model.mccormick_rule_4 = Constraint(rule = mccormick_rule4)



def mccormick_rule5(model):
    return model.w >= model.mccormick_x_lower * model.y + model.x * model.mccormick_y_lower - model.mccormick_x_lower * model.mccormick_y_lower
model.mccormick_rule_5 = Constraint(rule = mccormick_rule5)

def mccormick_rule6(model):
    return model.w >= model.mccormick_x_upper * model.y + model.x * model.mccormick_y_upper - model.mccormick_x_upper * model.mccormick_y_upper
model.mccormick_rule_6 = Constraint(rule = mccormick_rule6)

def mccormick_rule7(model):
    return model.w <= model.mccormick_x_upper * model.y + model.x * model.mccormick_y_lower - model.mccormick_x_upper * model.mccormick_y_lower
model.mccormick_rule_7 = Constraint(rule = mccormick_rule7)

def mccormick_rule8(model):
    return model.w <= model.x * model.mccormick_y_upper + model.y * model.mccormick_x_lower - model.mccormick_x_lower * model.mccormick_y_upper
model.mccormick_rule_8 = Constraint(rule = mccormick_rule8)



def last_rule(model):
    return model.w >= 18
model.last_rule = Constraint(rule = last_rule)



def objective_rule(model):
    return model.w + 6 * model.x - 2 * model.y
model.objective_rule = Objective(rule = objective_rule, sense = minimize)


# Create a solver and solve the MINLP problem
solver = SolverFactory('cplex')
results = solver.solve(model)

# Display the results
model.display()

# Access the solution values
x_value = value(model.x)
y_value = value(model.y)
w_value = value(model.w)

print(f"x = {x_value}")
print(f"y = {y_value}")
print(f"w = {w_value}")
print(f"objective_value = {value(model.objective_rule)}")