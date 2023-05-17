import pandas as pd
from pyomo.environ import *

model = AbstractModel()

# Sets
model.MONTHS = Set()

# Parameters
model.energySellingPrice = Param(model.MONTHS)
model.energyBuyingPrice = Param(model.MONTHS)
model.energyConsumption = Param(model.MONTHS)
model.E_pv = Param(model.MONTHS)
model.E_battery_max = Param(initialize=100)  # Set the desired constant value

# Decision variables
model.E_buy = Var(model.MONTHS, within=NonNegativeReals)
model.E_sell = Var(model.MONTHS, within=NonNegativeReals)
model.E_pv_use = Var(model.MONTHS, within=NonNegativeReals)
model.E_battery = Var(model.MONTHS, within=NonNegativeReals)
model.E_in_out = Var(model.MONTHS)

# Constraints
def consumptionRule(model, t):
    return model.energyConsumption[t] == model.E_buy[t] + model.E_pv_use[t]

model.consumption_rule = Constraint(model.MONTHS, rule=consumptionRule)

def batteryRule(model, t):
    if t == 1:
        return model.E_battery[t] == 0
    else:
        return model.E_battery[t] == model.E_battery[t-1] + model.E_in_out[t]
model.battery_rule = Constraint(model.MONTHS, rule=batteryRule)

def maxBatteryRule(model, t):
    return model.E_battery[t] <= model.E_battery_max

model.max_battery_rule = Constraint(model.MONTHS, rule=maxBatteryRule)

def energyInOut(model, t):
    return model.E_in_out[t] == (model.E_pv[t] - model.E_pv_use[t] - model.E_sell[t])

model.energy_In_Out = Constraint(model.MONTHS, rule=energyInOut)

# Objective function
def objectiveRule(model):
    return sum(model.E_buy[t] * model.energyBuyingPrice[t] - model.E_sell[t] * model.energySellingPrice[t] for t in model.MONTHS)

model.objective_function = Objective(rule=objectiveRule, sense=minimize)

data = DataPortal()
data.load(filename='optimizationData.dat')
instance = model.create_instance(data)

optimizer = SolverFactory('cplex')
optimizer.solve(instance)
instance.display()
