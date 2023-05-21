import pyomo.environ as pyo
import pandas as pd

path_output = './masterthesis/output/'
path_input = './masterthesis/input/'

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#parameters
model.starting_E_bat = pyo.Param(initialize = 50)
model.E_bat_max = pyo.Param(initialize = 100)
model.pv_eff = pyo.Param(initialize = 0.1)
model.E_solar = pyo.Param(model.HOURS)
model.E_demand = pyo.Param(model.HOURS)
model.costBuy = pyo.Param(model.HOURS)
model.costSell = pyo.Param(model.HOURS)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#variables
model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.E_pv_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_pv_bat = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_pv_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.E_bat = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_bat_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#constraints
def demand_rule(model,t):
    return model.E_demand[t] == model.E_pv_demand[t] + model.E_buy[t] + model.E_bat_demand[t]
model.demandRule = pyo.Constraint(model.HOURS,rule = demand_rule)

def pv_rule(model,t):
    return model.E_solar[t] + model.pv_eff == model.E_pv_bat[t] + model.E_sell[t] + model.E_pv_demand[t]
model.pvRule = pyo.Constraint(model.HOURS,rule = pv_rule)

def battery_rule(model,t):
        if t ==1:
             return model.E_bat[t] == model.starting_E_bat
        else:
            return model.E_bat[t] == model.E_bat[t-1] + model.E_pv_bat[t-1] - model.E_bat_demand[t-1]
model.batteryRule = pyo.Constraint(model.HOURS,rule = battery_rule)

def sell_rule(model,t):
    return model.E_sell[t] == model.E_pv_sell[t]
model.sellRule = pyo.Constraint(model.HOURS, rule = sell_rule)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#objective function
def objective_rule(model,t):
     return sum(model.E_buy[t] * model.costBuy[t] - model.E_sell[t] * model.costSell[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule = objective_rule,sense= pyo.minimize)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#reading data
data = pyo.DataPortal()
data.load(filename = path_input + 'HOURS.csv', set = 'HOURS')
data.load(filename = path_input + 'E_solar.csv', param='E_solar',index='HOURS')
data.load(filename = path_input + 'E_demand.csv', param='E_demand',index='HOURS')
data.load(filename = path_input + 'costBuy.csv', param='costBuy',index='HOURS')
data.load(filename = path_input + 'costSell.csv', param='costSell',index='HOURS')

#generating instance
instance = model.create_instance(data)

# instance.bateryCapacityRule.deactivate()

#solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
instance.pprint()
instance.display()

#exporting data
df = {'HOURS':[],
      'E_buy':[],
      'E_sell':[],
      'E_sell':[],
      'E_bat':[],
      'E_bat_demand':[],
      'E_pv_bat':[],
      'E_pv_demand':[],
      'E_pv_sell':[]
      }

for t in instance.HOURS:
    df['HOURS'].append(str(t))
    df['E_buy'].append(instance.E_buy[t].value)
    df['E_sell'].append(instance.E_sell[t].value)
    df['E_sell'].append(instance.E_sell[t].value)
    df['E_bat'].append(instance.E_bat[t].value)
    df['E_bat_demand'].append(instance.E_bat_demand[t].value)
    df['E_pv_bat'].append(instance.E_pv_bat[t].value)
    df['E_pv_demand'].append(instance.E_pv_demand[t].value)
    df['E_pv_sell'].append(instance.E_pv_sell[t].value)

print(len(df['E_buy']))

df = pd.DataFrame(df)
df.to_excel(path_output + 'df_results.xlsx', index=False)