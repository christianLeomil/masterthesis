import pandas as pd
import pyomo.environ as pyo
import classes
import inspect

path_output = './output/'
path_input = './input/'
name_file = 'df_input.xlsx'

df_input_series = pd.read_excel(path_input + name_file, sheet_name='series')
df_input_other = pd.read_excel(path_input + name_file,sheet_name = 'other')

df_elements = pd.read_excel(path_input + name_file, sheet_name = 'elements')
df_conect_to = pd.read_excel(path_input + name_file, sheet_name = 'connect_to',index_col=0)
df_conect_from = pd.read_excel(path_input + name_file, sheet_name = 'connect_from',index_col=0)

df_input_pv = pd.read_excel(path_input + name_file, sheet_name='pv')
df_input_bat = pd.read_excel(path_input + name_file, sheet_name='bat')

#end region
# ---------------------------------------------------------------------------------------------------------------------
# region elements

list_elements = []
for i in df_elements.columns:
     if df_elements[i].iloc[0] != 0:
          list_elements.append(i)

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region model, sets and parameters

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()
model.PV = pyo.Set()
model.BAT = pyo.Set()

#parameter scalar
model.time_step = pyo.Param()
model.pv_eff = pyo.Param(model.PV)

# model.starting_SOC = pyo.Param()
# model.E_bat_max = pyo.Param() #capacity of battery
# model.bat_ch_eff = pyo.Param() #charging efficiency of battery
# model.bat_dis_eff = pyo.Param() #discharging efficiency of battery
# model.c_rate_ch = pyo.Param() #maximal charging power (max charging power = c_rate * E_bat_max)
# model.c_rate_dis = pyo.Param() #maximal discharging power (max discharging power = c_rate * E_bat_max)

model.starting_SOC = pyo.Param(model.BAT) #starting SOC of battery
model.E_bat_max = pyo.Param(model.BAT) #capacity of battery
model.bat_ch_eff = pyo.Param(model.BAT) #charging efficiency of battery
model.bat_dis_eff = pyo.Param(model.BAT) #discharging efficiency of battery
model.c_rate_ch = pyo.Param(model.BAT) #maximal charging power (max charging power = c_rate * E_bat_max)
model.c_rate_dis = pyo.Param(model.BAT) #maximal discharging power (max discharging power = c_rate * E_bat_max)


#parameters series
model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
model.P_to_demand = pyo.Param(model.HOURS) #time series with solar energy
model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region variables

model.P_net_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_to_net = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.P_pv = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
model.P_pv_net = pyo.Var(model.HOURS, model.PV, within = pyo.NonNegativeReals)
model.P_pv_bat = pyo.Var(model.HOURS,model.PV, within = pyo.NonNegativeReals)
model.P_pv_demand = pyo.Var(model.HOURS, model.PV,within = pyo.NonNegativeReals)

# model.SOC = pyo.Var(model.HOURS, within = pyo.NonNegativeReals, bounds=(0, 1))
# model.P_to_bat = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
# model.P_bat_dis = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.P_bat_net = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
# model.P_bat_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.K_ch = pyo.Var(model.HOURS, domain = pyo.Binary)
# model.K_dis = pyo.Var(model.HOURS, domain = pyo.Binary)

model.SOC = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals, bounds=(0, 1))
model.P_to_bat = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals)
model.P_bat_dis = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals)
model.P_bat_net = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals)
model.P_bat_demand = pyo.Var(model.HOURS, model.BAT, within = pyo.NonNegativeReals)
model.K_ch = pyo.Var(model.HOURS, model.BAT, domain = pyo.Binary)
model.K_dis = pyo.Var(model.HOURS, model.BAT, domain = pyo.Binary)

model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)


#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating constraints from classes


for i in list_elements:
     constraint_class = getattr(classes,i)
     constraint_methods = [method_name for method_name in dir(constraint_class) if callable(getattr(constraint_class, method_name))]
     print('--------------Constraint Methods-----------')
     print(constraint_methods)
     for name in constraint_methods:
          if name.startswith('__'):
               pass
          else:
            method = getattr(constraint_class, name)
            name_set = getattr(model,i.upper())
            model.add_component(name, pyo.Constraint(model.HOURS, name_set, rule = method))

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating constraints from conections

# creating list of expressions "to"
df_conect_to.index.name = None

list_energy_from = []
list_energy_to_total = []
for i in range(0, len(df_conect_to)):
    list_energy_to = []
    for j in df_conect_to.columns:
        if df_conect_to[j].iloc[i] == 1:
            list_energy_to.append(j)
    if list_energy_to != []:
        list_energy_to_total.append(list_energy_to)
        list_energy_from.append(df_conect_to.index[i])

df_conect_to = pd.DataFrame({'from': list_energy_from, 'to': list_energy_to_total})
print(df_conect_to)
print('\n')

df_sets = pd.DataFrame({'name':['pv','bat'],
                          'var':['n','m'],
                          'set':['model.PV','model.BAT']})
list_expressions = []
list_sets = []
for i in df_conect_to.index:
    list_expression_conect =[]
    if df_conect_to['from'].iloc[i] in ['demand','net']:
        list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t] == ')
        list_sets.append('')
    else:
        list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t][' + df_sets['var'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0] + '] == ')
        list_sets.append(df_sets['set'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0])
    for index,j in enumerate(df_conect_to['to'].iloc[i]):
        if index != 0:
            list_expression_conect[-1] = list_expression_conect[-1] + ' +'
        if j not in ['demand','net']:
            list_expression_conect[-1] = list_expression_conect[-1] + ' sum(model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t][' + str(df_sets['var'][df_sets['name']==j].iloc[0]) +'] for ' + str(df_sets['var'][df_sets['name']==j].iloc[0]) + ' in ' + str(df_sets['set'][df_sets['name']==j].iloc[0]) + ')'
        else:
            list_expression_conect[-1] = list_expression_conect[-1] + ' model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t]'
    list_expression_conect = list_expression_conect[-1]
    list_expressions.append(list_expression_conect)

print('---------------------------------------List Expressions:')
print(list_expressions)
print('---------------------------------------List Sets')
print(list_sets)

#creating constraints from connections
class myClass:
    pass

constraint_number = 1
for i in list_expressions:
    def dynamic_method(model,t,n,m,expr):
        return eval(expr, globals(), locals())

    method_name = 'constraint_conect' + str(constraint_number)
#     print('-----------------------method name-----------------------')
#     print(method_name)
    def method_wrapper(self, model, t, n,m,expr=i):
        return dynamic_method(model,t,n,m,expr)

    setattr(myClass, method_name, method_wrapper)

    my_obj = myClass()

    setattr(model, 'constraint_conect' + str(constraint_number), pyo.Constraint(model.HOURS,model.PV,
                                                                                model.BAT,rule=getattr(my_obj, method_name)))
    constraint_number = constraint_number + 1

#endregion
# ---------------------------------------------------------------------------------------------------------------------
# region continuing other constraints


def pv_rule(model,t,n,m):
    return model.P_pv[t,n] == sum(model.P_pv_bat[t,n,m] for m in model.BAT) + model.P_pv_net[t,n] + model.P_pv_demand[t,n]
model.pvRule = pyo.Constraint(model.HOURS, model.PV,model.BAT,rule = pv_rule)

#created through conections "to" 
# def demand_rule(model,t):
#     return model.P_to_demand[t] == sum(model.P_pv_demand[t,n] for n in model.PV) + model.P_net_demand[t] + model.P_bat_demand[t]
# model.demandRule = pyo.Constraint(model.HOURS,rule = demand_rule)


# created through class
# def battery_rule(model,t):
#         if t ==1:
#              return model.SOC[t] == model.starting_SOC
#         else:
#             return model.SOC[t] == model.SOC[t-1] + (model.P_to_bat[t-1] * model.bat_ch_eff -
#                                                       model.P_bat_dis[t-1]/model.bat_dis_eff) * model.time_step / model.E_bat_max
# model.batteryRule = pyo.Constraint(model.HOURS,rule = battery_rule)

# created through connections "to"
# def charge_rule(model,t):
#      return model.P_to_bat[t] == sum(model.P_pv_bat[t,n] for n in model.PV)
# model.chargeRule = pyo.Constraint(model.HOURS,rule = charge_rule)


def discharge_rule(model,t,m):
     return model.P_bat_dis[t,m] == model.P_bat_net[t,m] + model.P_bat_demand[t,m]
model.dischargeRule = pyo.Constraint(model.HOURS,model.BAT,rule = discharge_rule)


# created through classes
# def charge_limit(model,t):
#      return model.P_to_bat[t] <= model.E_bat_max * model.c_rate_ch * model.K_ch[t]
# model.chargeLimit = pyo.Constraint(model.HOURS,rule = charge_limit)


# created through class
# def discharge_limit(model,t):
#      return model.P_bat_dis[t] <= model.E_bat_max * model.c_rate_dis * model.K_dis[t]
# model.dischargeLimit = pyo.Constraint(model.HOURS,rule = discharge_limit)


# created through class
# def keys_rule(model,t):
#      return model.K_ch[t] + model.K_dis[t] <= 1
# model.keysRule = pyo.Constraint(model.HOURS, rule = keys_rule)

# created through connections
# def sell_rule(model,t):
#      return model.P_to_net[t] == sum(model.P_pv_net[t,n] for n in model.PV) + model.P_bat_net[t]
# model.sellRule = pyo.Constraint(model.HOURS,rule = sell_rule)


# created through classes
# def sell_energy(model,t):
#     return model.E_sell[t] == model.P_to_net[t] * model.time_step
# model.sellEnergy = pyo.Constraint(model.HOURS, rule = sell_energy)


# created through classes
# def buy_energy(model,t):
#      return model.E_buy[t] == model.P_net_demand[t] * model.time_step
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

data['PV'] = df_input_pv['pv'].tolist()
data['pv_eff'] = df_input_pv.set_index('pv')['pv_eff'].to_dict()

data['BAT'] = df_input_bat['bat'].tolist()
data['starting_SOC'] = df_input_bat.set_index('bat')['starting_SOC'].to_dict()
data['E_bat_max'] = df_input_bat.set_index('bat')['E_bat_max'].to_dict()
data['bat_ch_eff'] = df_input_bat.set_index('bat')['bat_ch_eff'].to_dict()
data['bat_dis_eff'] = df_input_bat.set_index('bat')['bat_dis_eff'].to_dict()
data['c_rate_ch'] = df_input_bat.set_index('bat')['c_rate_ch'].to_dict()
data['c_rate_dis'] = df_input_bat.set_index('bat')['c_rate_dis'].to_dict()

# data['E_bat_max'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'E_bat_max', 'Value'].values[0]}
# data['starting_SOC'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'starting_SOC', 'Value'].values[0]}
# data['c_rate_ch'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'c_rate_ch', 'Value'].values[0]}
# data['c_rate_dis'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'c_rate_dis', 'Value'].values[0]}
# data['bat_ch_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat_ch_eff', 'Value'].values[0]}
# data['bat_dis_eff'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'bat_dis_eff', 'Value'].values[0]}

data['time_step'] = {None:df_input_other.loc[df_input_other['Parameter'] == 'time_step', 'Value'].values[0]}

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region starting model

#generating instance
instance = model.create_instance(data)

print("Constraint Expressions:")
for constraint in instance.component_objects(pyo.Constraint):
    for index in constraint:
        print(f"{constraint}[{index}]: {constraint[index].body}")

#solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
instance.display()

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region export data

# # Extracting variable data from the Pyomo model
# variable_data = {
#     'P_net_demand': [instance.P_net_demand[t].value for t in instance.HOURS],
#     'P_sell': [instance.P_sell[t].value for t in instance.HOURS],
#     'P_pv': {n: [instance.P_pv[t, n].value for t in instance.HOURS] for n in instance.PV},
#     'P_pv_net': {n: [instance.P_pv_net[t, n].value for t in instance.HOURS] for n in instance.PV},
#     'P_pv_bat': {n: [instance.P_pv_bat[t, n].value for t in instance.HOURS] for n in instance.PV},
#     'P_pv_demand': {n: [instance.P_pv_demand[t, n].value for t in instance.HOURS] for n in instance.PV},
#     'SOC': [instance.SOC[t].value for t in instance.HOURS],
#     'P_bat_ch': [instance.P_bat_ch[t].value for t in instance.HOURS],
#     'P_bat_dis': [instance.P_bat_dis[t].value for t in instance.HOURS],
#     'P_bat_net': [instance.P_bat_net[t].value for t in instance.HOURS],
#     'P_bat_demand': [instance.P_bat_demand[t].value for t in instance.HOURS],
#     'K_ch': [instance.K_ch[t].value for t in instance.HOURS],
#     'K_dis': [instance.K_dis[t].value for t in instance.HOURS],
#     'E_sell': [instance.E_sell[t].value for t in instance.HOURS],
#     'E_buy': [instance.E_buy[t].value for t in instance.HOURS]
# }

# # Create a DataFrame from the variable data
# df_variable_data = pd.DataFrame(variable_data, index=instance.HOURS)

# # Define the path and filename for the output Excel file
# output_file_name = 'variable_data.xlsx'

# # Export the DataFrame to Excel
# df_variable_data.to_excel(path_output + output_file_name, index_label='Hours')


#endregion