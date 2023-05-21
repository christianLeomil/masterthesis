from pyomo.environ import *
import pandas as pd

path_output = './masterthesis/output/'
path_parameters = './masterthesis/input/old/'

model = AbstractModel()

# Sets
model.HOURS = Set()

# Parameters
model.demand = Param(model.HOURS)
model.costA = Param(initialize=20)
model.costB = Param(initialize=15)
model.workCapA = Param(model.HOURS)
model.workCapB = Param(model.HOURS)

# Variables
model.amountA = Var(model.HOURS, within=NonNegativeIntegers, initialize=10)
model.amountB = Var(model.HOURS, within=NonNegativeIntegers)

# Constraints
def demand_rule(model, t):
    return model.amountA[t] * model.workCapA[t] + model.amountB[t] * model.workCapB[t] >= model.demand[t]
model.demandRule = Constraint(model.HOURS, rule=demand_rule)

def time_step_rule(model, t):
    if t == 1:
        return model.amountA[t] == 10
    else: 
        return model.amountA[t] <= 1.5 * model.amountA[t-1]
model.timeStepRule = Constraint(model.HOURS, rule=time_step_rule)

# Objective function
def objective_rule(model):
    return sum(model.amountA[t] * model.costA + model.amountB[t] * model.costB for t in model.HOURS)
model.objective_function = Objective(rule=objective_rule, sense=minimize)

# Reading data
data = DataPortal()
data.load(filename=path_parameters+'HOURS.csv', set='HOURS')
data.load(filename=path_parameters+'demand.csv', param='demand', index='HOURS')
data.load(filename=path_parameters+'workCapA.csv', param='workCapA', index='HOURS')
data.load(filename=path_parameters+'workCapB.csv', param='workCapB', index='HOURS')

instance = model.create_instance(data)

# instance.timeStepRule.deactivate()

# Solving the model
optimizer = SolverFactory('cplex')
results = optimizer.solve(instance)

# Displaying the results
instance.pprint()
instance.display()

# Exporting variable values to Excel
variable_values = {
    'Hours': [],
    'AmountA': [],
    'AmountB': []
}

for t in instance.HOURS:
    variable_values['Hours'].append(str(t))
    variable_values['AmountA'].append(instance.amountA[t].value)
    variable_values['AmountB'].append(instance.amountB[t].value)

df = pd.DataFrame(variable_values)
df.to_excel(path_output + 'variable_values.xlsx', index=False)
