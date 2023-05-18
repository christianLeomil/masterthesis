from pyomo.environ import *
import pandas as pd

path_output = './masterthesis/output/'
path_input = './masterthesis/input/'

#model
model=AbstractModel()

#sets
model.HOURS=Set()

#parameters
model.starting_E_bat = Param(initialize = 50)
model.E_bat_max = Param(initialize = 100)
model.pv_eff = Param(initialize = 0.1)
model.E_solar = Param(model.HOURS)
model.E_demand = Param(model.HOURS)
model.costBuy = Param(model.HOURS)
model.costSell = Param(model.HOURS)

#variables
model.E_buy_demand = Var(model.HOURS,within=NonNegativeReals)
model.E_buy_bat = Var(model.HOURS,within=NonNegativeReals)

model.E_sell = Var(model.HOURS,within=NonNegativeReals)

model.E_bat = Var(model.HOURS,within = NonNegativeReals)
model.E_bat_demand = Var(model.HOURS, within = NonNegativeReals)
model.E_bat_sell = Var(model.HOURS, within = NonNegativeReals)

model.E_pv_bat = Var(model.HOURS,within = NonNegativeReals)
model.E_pv_demand = Var(model.HOURS,within = NonNegativeReals)
model.E_pv_sell = Var(model.HOURS,within = NonNegativeReals)

#objective function
def objective_rule(model):
    return sum( model.E_bat[t] * model.costBuy[t] - model.E_sell[t] * model.costSell[t] for t in model.HOURS)
model.objectiveRule = Objective(rule = objective_rule , sense = minimize)

#contraints
def battery_capacity_rule(model,t):
    return model.E_bat[t] <= model.E_bat_max
model.bateryCapacityRule = Constraint(model.HOURS, rule = battery_capacity_rule)

def demand_rule(model,t):
    return (model.E_pv_demand[t] + model.E_buy_demand[t] + model.E_bat_demand[t]) == model.E_demand[t]
model.demandRule = Constraint(model.HOURS,rule = demand_rule)

def battery_charge_rule(model,t):
    if t == 1:
        return model.E_bat[t] == model.starting_E_bat
    else:
        return model.E_bat[t] == (model.E_bat[t-1] + model.E_pv_bat[t] + model.E_buy_bat[t]
                                  - model.E_bat_demand[t] - model.E_bat_sell[t])
model.batteryChargeRule = Constraint(model.HOURS, rule = battery_charge_rule)

def solar_energy_rule(model,t):
    return model.E_solar[t] * model.pv_eff == model.E_pv_bat[t] + model.E_pv_demand[t] + model.E_pv_sell[t]
model.solarEnergyRule = Constraint(model.HOURS,rule=solar_energy_rule)

def sell_rule(model,t):
    return model.E_sell[t] == model.E_pv_sell[t] + model.E_bat_sell[t]
model.sellRule = Constraint(model.HOURS,rule=sell_rule)

def buy_rule(model,t):
    return model.E_buy[t] == model.E_buy_demand[t] + model.E_buy_bat[t]
model.sellRule = Constraint(model.HOURS,rule=sell_rule)

#reading data
data = DataPortal()
data.load(filename = path_input + 'HOURS.csv', set = 'HOURS')
data.load(filename = path_input + 'E_solar.csv', param='E_solar',index='HOURS')
data.load(filename = path_input + 'E_demand.csv', param='E_demand',index='HOURS')
data.load(filename = path_input + 'costBuy.csv', param='costBuy',index='HOURS')
data.load(filename = path_input + 'costSell.csv', param='costSell',index='HOURS')

#generating instance
instance = model.create_instance(data)

#solving the model
optimizer = SolverFactory('cplex')
results = optimizer.solve(instance)

# Displaying the results
instance.pprint()
instance.display()

#exporting data
df = {'HOURS':[],
      'E_buy_demand':[],
      'E_buy_bat':[],
      'E_sell':[],
      'E_bat':[],
      'E_bat_demand':[],
      'E_bat_sell':[],
      'E_pv_bat':[],
      'E_pv_demand':[],
      'E_pv_sell':[]
      }

for t in instance.HOURS:
    df['HOURS'].append(str(t))
    df['E_buy_demand'].append(instance.E_buy_demand[t].value)
    df['E_buy_bat'].append(instance.E_buy_bat[t].value)
    df['E_sell'].append(instance.E_sell[t].value)
    df['E_bat'].append(instance.E_bat[t].value)
    df['E_bat_demand'].append(instance.E_bat_demand[t].value)
    df['E_bat_sell'].append(instance.E_bat_sell[t].value)
    df['E_pv_bat'].append(instance.E_pv_bat[t].value)
    df['E_pv_demand'].append(instance.E_pv_demand[t].value)
    df['E_pv_sell'].append(instance.E_pv_sell[t].value)

df = pd.DataFrame(df)
df.to_excel(path_output + 'df_results.xlsx', index=False)
