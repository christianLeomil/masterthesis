import pandas as pd
from pyomo.environ import *

model=AbstractModel()

#Sets
model.MONTHS = Set()

#parameters
model.energySellingPrice = Param(model.MONTHS)
model.energyBuyingPrice = Param(model.MONTHS)
model.energyConsumption = Param(model.MONTHS)
model.E_pv = Param(model.MONTHS)

#decision variables
model.E_buy = Var(model.MONTHS,within = NonNegativeReals)
model.E_sell = Var(model.MONTHS,within = NonNegativeReals)
model.E_pv_use = Var(model.MONTHS,within = NonNegativeReals)


#constraints
def consumptionRule(model,t):
    return model.energyConsumption[t] == model.E_buy[t] + model.E_pv_use[t]
model.consumption_rule = Constraint(model.MONTHS,rule = consumptionRule)

def energySellRule(model,t):
    return model.E_sell[t] == model.E_pv[t] - model.E_pv_use[t] 
model.energy_sell = Constraint(model.MONTHS,rule = energySellRule)


#objective function
def objectiveRule(model):
    return  sum(model.E_buy[t]*model.energyBuyingPrice[t]-model.E_sell[t]*model.energySellingPrice[t] for t in model.MONTHS)
model.objective_function = Objective(rule = objectiveRule,sense = minimize)

data = DataPortal()
data.load(filename='optimizationData.dat')
instance = model.create_instance(data)

optimizer = SolverFactory('cplex')
optimizer.solve(instance)
instance.display()