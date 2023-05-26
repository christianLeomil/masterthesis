import pyomo.environ as pyo
import pandas as pd

path_output = './output/'
path_input = './input/'

df_input_series = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'series')
df_input_other = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'other')

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#parameters other
model.time_step = pyo.Param()

model.pv_eff = pyo.Param()

model.starting_SOC = pyo.Param() #starting SOC of battery
model.E_bat_max = pyo.Param() #capacity of battery
model.bat_ch_eff = pyo.Param() #charging efficiency of battery
model.bat_dis_eff = pyo.Param() #discharging efficiency of battery
model.c_rate_ch = pyo.Param() #maximal charging power (max charging power = c_rate * E_bat_max)
model.c_rate_dis = pyo.Param() #maximal discharging power (max discharging power = c_rate * E_bat_max)

#parameters series
model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
model.P_demand = pyo.Param(model.HOURS) #time series with solar energy
model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#variables
model.P_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.P_pv_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_pv_bat = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_pv_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.SOC = pyo.Var(model.HOURS, within = pyo.NonNegativeReals, bounds=(0, 1))
model.P_bat_ch = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
model.P_bat_dis = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_bat_sell = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
model.P_bat_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.K_ch = pyo.Var(model.HOURS, domain = pyo.Binary)
model.K_dis = pyo.Var(model.HOURS, domain = pyo.Binary)

model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constraints

#constraints

def demand_rule(model,t):
    return model.P_demand[t] == model.P_pv_demand[t] + model.P_buy[t] + model.P_bat_demand[t]
model.demandRule = pyo.Constraint(model.HOURS,rule = demand_rule)


def pv_rule(model,t):
    return model.P_solar[t] * model.pv_eff == model.P_pv_bat[t] + model.P_pv_sell[t] + model.P_pv_demand[t]
model.pvRule = pyo.Constraint(model.HOURS,rule = pv_rule)


def battery_rule(model,t):
        if t ==1:
             return model.SOC[t] == model.starting_SOC
        else:
            return model.SOC[t] == model.SOC[t-1] + (model.P_bat_ch[t-1] * model.bat_ch_eff -
                                                      model.P_bat_dis[t-1]/model.bat_dis_eff) * model.time_step / model.E_bat_max
model.batteryRule = pyo.Constraint(model.HOURS,rule = battery_rule)


def charge_rule(model,t):
     return model.P_bat_ch[t] == model.P_pv_bat[t]
model.chargeRule = pyo.Constraint(model.HOURS,rule = charge_rule)


def discharge_rule(model,t):
     return model.P_bat_dis[t] == model.P_bat_sell[t] + model.P_bat_demand[t]
model.dischargeRule = pyo.Constraint(model.HOURS,rule = discharge_rule)


def charge_limit(model,t):
     return model.P_bat_ch[t] <= model.E_bat_max * model.c_rate_ch * model.K_ch[t]
model.chargeLimit = pyo.Constraint(model.HOURS,rule = charge_limit)

def discharge_limit(model,t):
     return model.P_bat_dis[t] <= model.E_bat_max * model.c_rate_dis * model.K_dis[t]
model.dischargeLimit = pyo.Constraint(model.HOURS,rule = discharge_limit)


def keys_rule(model,t):
     return model.K_ch[t] + model.K_dis[t] <= 1
model.keysRule = pyo.Constraint(model.HOURS, rule = keys_rule)


def sell_rule(model,t):
     return model.P_sell[t] == model.P_pv_sell[t] + model.P_bat_sell[t]
model.sellRule = pyo.Constraint(model.HOURS,rule = sell_rule)


def sell_energy(model,t):
    return model.E_sell[t] == model.P_sell[t] * model.time_step
model.sellEnergy = pyo.Constraint(model.HOURS, rule = sell_energy)


def buy_rule(model,t):
     return model.P_buy[t] == model.P_demand[t] - model.P_pv_demand[t] - model.P_bat_demand[t]
model.buyRule = pyo.Constraint(model.HOURS, rule = buy_rule)


def buy_energy(model,t):
     return model.E_buy[t] == model.P_buy[t] * model.time_step
model.buyEnergy = pyo.Constraint(model.HOURS, rule = buy_energy)


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
data['HOURS'] = df_input_series['HOURS'].tolist()
data['P_solar'] = df_input_series.set_index('HOURS')['P_solar'].to_dict()
data['P_demand'] = df_input_series.set_index('HOURS')['P_demand'].to_dict()
data['costBuy'] = df_input_series.set_index('HOURS')['costBuy'].to_dict()
data['costSell'] = df_input_series.set_index('HOURS')['costSell'].to_dict()

data['Q_demand'] = df_input_series.set_index('HOURS')['Q_demand'].to_dict()

data['E_bat_max'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'E_bat_max', 'Value'].values[0]}
data['starting_SOC'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'starting_SOC', 'Value'].values[0]}
data['pv_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'pv_eff', 'Value'].values[0]}
data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}
data['c_rate_ch'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'c_rate_ch', 'Value'].values[0]}
data['c_rate_dis'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'c_rate_dis', 'Value'].values[0]}
data['bat_ch_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat_ch_eff', 'Value'].values[0]}
data['bat_dis_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat_dis_eff', 'Value'].values[0]}

#generating instance
instance = model.create_instance(data)

#solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
instance.display()

#exporting data
df = {'HOURS':[],
      'P_demand':[],
      'P_solar':[],
      'P_buy':[],
      'P_sell':[],
      'P_bat_demand':[],
      'P_bat_ch':[],
      'P_bat_dis':[],
      'P_bat_sell':[],
      'P_pv_bat':[],
      'P_pv_demand':[],
      'P_pv_sell':[],
      'E_sell':[],
      'E_buy':[],
      'SOC':[],
      'K_ch':[],
      'K_dis':[]
      }

for t in instance.HOURS:
    df['HOURS'].append(str(t))
    df['P_demand'].append(instance.P_demand[t])
    df['P_solar'].append(instance.P_solar[t])
    df['P_buy'].append(instance.P_buy[t].value)
    df['P_sell'].append(instance.P_sell[t].value)
    df['P_bat_demand'].append(instance.P_bat_demand[t].value)
    df['P_bat_ch'].append(instance.P_bat_ch[t].value)
    df['P_bat_dis'].append(instance.P_bat_dis[t].value)
    df['P_bat_sell'].append(instance.P_bat_sell[t].value)
    df['P_pv_bat'].append(instance.P_pv_bat[t].value)
    df['P_pv_demand'].append(instance.P_pv_demand[t].value)
    df['P_pv_sell'].append(instance.P_pv_sell[t].value)
    df['E_sell'].append(instance.E_sell[t].value)
    df['E_buy'].append(instance.E_buy[t].value)
    df['SOC'].append(instance.SOC[t].value)
    df['K_ch'].append(instance.K_ch[t].value)
    df['K_dis'].append(instance.K_dis[t].value)

column_keys = df.keys()

df = pd.DataFrame(df)
df.to_excel(path_output + 'df_results.xlsx', index=False)

#endregion
