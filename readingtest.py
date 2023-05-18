from pyomo.environ import *

model = AbstractModel()

# Sets
model.HOURS = Set()

# Parameters
model.demand = Param(model.HOURS)
model.costA = Param(initialize = 20)
model.costB = Param(initialize = 15)
model.workCapA = Param(model.HOURS)
model.workCapB = Param(model.HOURS)

# Variables
model.amountA = Var(model.HOURS, within=NonNegativeIntegers)
model.amountB = Var(model.HOURS, within=NonNegativeIntegers)

# Constraints
def demand_rule(model, t):
    return model.amountA[t] * model.workCapA[t] + model.amountB[t] * model.workCapB[t] >= model.demand[t]
model.demand_rule = Constraint(model.HOURS, rule=demand_rule)

# Objective function
def objective_rule(model):
    return sum(model.amountA[t] * model.costA + model.amountB[t] * model.costB for t in model.HOURS)
model.objective_function = Objective(rule=objective_rule, sense=minimize)

# Reading data
data = DataPortal()
data.load(filename='HOURS.csv', set='HOURS')
data.load(filename='demand.csv', param='demand',index = 'HOURS')
data.load(filename='workCapA.csv', param='workCapA',index = 'HOURS')
data.load(filename='workCapB.csv', param='workCapB',index = 'HOURS')

instance = model.create_instance(data)

# Solving the model
optimizer = SolverFactory('cplex')
results = optimizer.solve(instance)

# Displaying the results
instance.pprint()
instance.display()
