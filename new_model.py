import pyomo.environ as pyo
import pandas as pd
import classes
import functions

# ---------------------------------------------------------------------------------------------------------------------
# region reading data
path_input = './input/'
path_output = './output/'
name_file = 'df_input.xlsx'

df_input_series = pd.read_excel(path_input +name_file,sheet_name='series')
df_input_other = pd.read_excel(path_input + name_file, sheet_name = 'other')
df_input_pv = pd.read_excel(path_input + name_file,sheet_name='pv')
df_input_bat = pd.read_excel(path_input + name_file,sheet_name='bat')
df_aux = pd.read_excel(path_input + name_file,sheet_name='aux')

df_conect = pd.read_excel(path_input + name_file,sheet_name='conect',index_col=0)
df_conect.index.name = None

[df_conect, df_variables] = functions.connection_creator(df_conect,df_aux,path_output)
list_expressions = functions.exp_creator(df_conect,df_variables)
print(list_expressions)

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating abstract model

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()
model.PV = pyo.Set()
model.BAT = pyo.Set()

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating parameters

#parameter scalar
model.time_step = pyo.Param()
model.pv_eff = pyo.Param(model.PV)

model.starting_SOC = pyo.Param(model.BAT)
model.E_bat_max = pyo.Param(model.BAT)

#parameters series
model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
model.P_to_demand = pyo.Param(model.HOURS) #time series with solar energy
model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating variables

model.P_to_net = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_from_net = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_net_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_net_bat = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals)

model.P_from_pv = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
model.P_pv_net = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
model.P_pv_bat = pyo.Var(model.HOURS,model.PV, model.BAT, within = pyo.NonNegativeReals)
model.P_pv_demand = pyo.Var(model.HOURS, model.PV,within = pyo.NonNegativeReals)

model.SOC = pyo.Var(model.HOURS, model.BAT,within = pyo.NonNegativeReals, bounds=(0, 1))
model.P_to_bat = pyo.Var(model.HOURS,model.BAT, within = pyo.NonNegativeReals)
model.P_from_bat = pyo.Var(model.HOURS,model.BAT, within = pyo.NonNegativeReals)
model.P_bat_net = pyo.Var(model.HOURS,model.BAT, within = pyo.NonNegativeReals)
model.P_bat_demand = pyo.Var(model.HOURS,model.BAT, within = pyo.NonNegativeReals)

model.K_ch = pyo.Var(model.HOURS,model.BAT, domain = pyo.Binary)
model.K_dis = pyo.Var(model.HOURS,model.BAT, domain = pyo.Binary)

model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating constraints from classes

list_elements = ['pv','bat','general']

for i in list_elements:
     constraint_class = getattr(classes,i)
     constraint_methods = [method_name for method_name in dir(constraint_class) if callable(getattr(constraint_class, method_name))]
     # print('--------------Constraint Methods-----------')
     # print(constraint_methods)
     for name in constraint_methods:
          if name.startswith('__'):
               pass
          elif i !='general':
               method = getattr(constraint_class, name)
               name_set = getattr(model,i.upper())
               model.add_component(name, pyo.Constraint(model.HOURS, name_set, rule = method))
          else:
               method = getattr(constraint_class,name)
               model.add_component(name,pyo.Constraint(model.HOURS,rule = method))

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region connection constraints

class myClass:
    pass

constraint_number = 1
for i in list_expressions:
    def dynamic_method(model,t,n,m,expr):
        return eval(expr, globals(), locals())
    method_name = 'Constraint_' + str(constraint_number)

    def method_wrapper(self, model,t,n,m,expr=i):
        return dynamic_method(model,t,n,m,expr)
    setattr(myClass, method_name, method_wrapper)
    my_obj = myClass()
    setattr(model, 'Constraint_' + str(constraint_number), pyo.Constraint(model.HOURS,model.PV, model.BAT,
                                                                             rule=getattr(my_obj, method_name)))
    constraint_number = constraint_number + 1

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region other constraints

# #demand rule
# def demand_rule(model,t,n,m):
#     return model.P_to_demand[t] == (sum(model.P_pv_demand[t,n] for n in model.PV) + model.P_net_demand[t] + 
#                                     sum(model.P_bat_demand[t,m] for m in model.BAT))
# model.demandRule = pyo.Constraint(model.HOURS,model.PV,model.BAT,rule = demand_rule)

# #net_rule
# def net_rule(model,t,m):
#      return model.P_from_net[t] == model.P_to_demand[t] + sum(model.P_net_bat[t,m] for m in model.BAT)
# model.netRule = pyo.Constraint(model.HOURS, model.BAT, rule = net_rule)

# # pv rules
# def solar_rule(model,t,n):
#     return model.P_from_pv[t,n] == model.P_solar[t] * model.pv_eff[n]
# model.solarRule = pyo.Constraint(model.HOURS,model.PV,rule = solar_rule)

# def pv_rule(model,t,n,m):
#     return model.P_from_pv[t,n] == model.P_pv_bat[t,n,m] + model.P_pv_net[t,n] + model.P_pv_demand[t,n]
# model.pvRule = pyo.Constraint(model.HOURS , model.PV, model.BAT, rule = pv_rule)

# battery rule
# def battery_rule(model,t,m):
#         if t ==1:
#              return model.SOC[t,m] == model.starting_SOC[m]
#         else:
#             return model.SOC[t,m] == model.SOC[t-1,m] + (model.P_to_bat[t-1,m] - model.P_from_bat[t-1,m]) * model.time_step / model.E_bat_max[m]
# model.batteryRule = pyo.Constraint(model.HOURS,model.BAT,rule = battery_rule)

# def charge_rule(model,t,n,m):
#      return model.P_to_bat[t,m] == model.P_pv_bat[t,n,m] + model.P_net_bat[t,m]
# model.chargeRule = pyo.Constraint(model.HOURS,model.PV,model.BAT,rule = charge_rule)

# def discharge_rule(model,t,m):
#      return model.P_from_bat[t,m] == model.P_bat_net[t,m] + model.P_bat_demand[t,m]
# model.dischargeRule = pyo.Constraint(model.HOURS,model.BAT,rule = discharge_rule)

# def charge_limit(model,t,m):
#      return model.P_to_bat[t,m] <= model.E_bat_max[m] * 1 * model.K_ch[t,m]
# model.chargeLimit = pyo.Constraint(model.HOURS,model.BAT,rule = charge_limit)

# def discharge_limit(model,t,m):
#      return model.P_from_bat[t,m] <= model.E_bat_max[m] * 1 * model.K_dis[t,m]
# model.dischargeLimit = pyo.Constraint(model.HOURS,model.BAT,rule = discharge_limit)

# def keys_rule(model,t,m):
#      return model.K_ch[t,m] + model.K_dis[t,m] <= 1
# model.keysRule = pyo.Constraint(model.HOURS,model.BAT, rule = keys_rule)

# #related to objective rule
# def sell_rule(model,t,n,m):
#      return model.P_to_net[t] == sum(model.P_pv_net[t,n] for n in model.PV) + sum(model.P_bat_net[t,m] for m in model.BAT)
# model.sellRule = pyo.Constraint(model.HOURS,model.PV,model.BAT,rule = sell_rule)

# def sell_energy(model,t):
#     return model.E_sell[t] == model.P_to_net[t] * model.time_step
# model.sellEnergy = pyo.Constraint(model.HOURS, rule = sell_energy)

# def buy_energy(model,t):
#      return model.E_buy[t] == model.P_from_net[t] * model.time_step  
# model.buyEnergy = pyo.Constraint(model.HOURS, rule = buy_energy)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region objective

#objective function
def objective_rule(model,t):
     return sum(model.E_buy[t] * model.costBuy[t] - model.E_sell[t] * model.costSell[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule = objective_rule,sense= pyo.minimize)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region reading data


#reading data
data = pyo.DataPortal()
data['HOURS'] = df_input_series['HOURS'].tolist()
data['P_solar'] = df_input_series.set_index('HOURS')['P_solar'].to_dict()
data['P_to_demand'] = df_input_series.set_index('HOURS')['P_to_demand'].to_dict()
data['costBuy'] = df_input_series.set_index('HOURS')['costBuy'].to_dict()
data['costSell'] = df_input_series.set_index('HOURS')['costSell'].to_dict()

data['PV'] = df_input_pv['PV'].tolist()
data['pv_eff'] = df_input_pv.set_index('PV')['pv_eff'].to_dict()

data['BAT'] = df_input_bat['BAT'].tolist()
data['E_bat_max'] = df_input_bat.set_index('BAT')['E_bat_max'].to_dict()
data['starting_SOC'] = df_input_bat.set_index('BAT')['starting_SOC'].to_dict()

# data['E_bat_max'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'E_bat_max', 'Value'].values[0]}
# data['starting_SOC'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'starting_SOC', 'Value'].values[0]}
data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region starting model

#generating instance
instance = model.create_instance(data)

#solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
instance.display()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region exporting data

# endregion

