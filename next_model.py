import pyomo.environ as pyo
import pandas as pd

path_output = './output/'
path_input = './input/'
name_file = 'df_input_test.xlsx'

df_input_series = pd.read_excel(path_input + name_file,sheet_name = 'series')
df_conect = pd.read_excel(path_input + 'c_matrix.xlsx', sheet_name = 'conect')

# ---------------------------------------------------------------------------------------------------------------------
# region start

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

#parameters series
model.time_step = pyo.Param(initialize = 1)

model.pv_eff = pyo.Param(initialize = 0.1)

model.starting_SOC = pyo.Param(initialize = 0.5) #starting SOC of battery
model.E_bat_max = pyo.Param(initialize = 200) #capacity of battery
model.bat_ch_eff = pyo.Param(initialize = 0.98) #charging efficiency of battery
model.bat_dis_eff = pyo.Param(initialize = 0.98) #discharging efficiency of battery
model.c_rate_ch = pyo.Param(initialize = 1) #maximal charging power (max charging power = c_rate * E_bat_max)
model.c_rate_dis = pyo.Param(initialize = 1) #maximal discharging power (max discharging power = c_rate * E_bat_max)

#parameters series
model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
model.P_pv = pyo.Param(model.HOURS)
model.P_demand = pyo.Param(model.HOURS) #time series with solar energy
model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid

#variables
model.P_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.P_pv_net = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_pv_bat = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_pv_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.SOC = pyo.Var(model.HOURS, within = pyo.NonNegativeReals, bounds=(0, 1))
model.P_bat_ch = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
model.P_bat_dis = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.P_bat_net = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
model.P_bat_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

model.K_ch = pyo.Var(model.HOURS, domain = pyo.Binary)
model.K_dis = pyo.Var(model.HOURS, domain = pyo.Binary)

model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

#objective function
def objective_rule(model,t):
     return sum(model.E_buy[t] * model.costBuy[t] - model.E_sell[t] * model.costSell[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule = objective_rule,sense= pyo.minimize)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating strings of constraints

df_conect.set_index(df_conect['from'],inplace=True)
df_conect.index.name = None
df_conect.drop(['from'],axis = 1,inplace=True)

list_variables_conect = []
list_energy_from = []
list_energy_to_total = []
for i in range(0,len(df_conect)):
    list_energy_to = []
    for j in df_conect.columns:
        if df_conect[j].iloc[i] == 1:
            list_energy_to.append(j)
    if list_energy_to != []:
        list_energy_to_total.append(list_energy_to)
        list_energy_from.append(df_conect.index[i])

df_conect = pd.DataFrame({'energy from':list_energy_from,
                          'energy to':list_energy_to_total})

list_string_total = []
for i in range(0,len(df_conect)):
    string_partial = ''
    string_partial ='model.P_' + str(df_conect['energy from'].iloc[i]) + '[t]' '== model.P_' + str(df_conect['energy from'].iloc[i]) + '_' +str(df_conect['energy to'].iloc[i][0])+ '[t]'
    for j in range(1,len(df_conect['energy to'].iloc[i])):
        string_partial = string_partial + '+model.P_' + str(df_conect['energy from'].iloc[i]) + '_' + str(df_conect['energy to'].iloc[i][j])+ '[t]'
    list_string_total.append(string_partial)
print(list_string_total)

# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constriant builder


class ConstraintBuilder:
    def __init__(self, model):
        self.model = model

    def create_constraint(self, constraint_name, expression_str):
        expression = eval(expression_str, globals(), locals())
        setattr(self, constraint_name, expression)
        constraint = pyo.Constraint(expr=expression)
        setattr(self.model, constraint_name, constraint)

builder = ConstraintBuilder(model)
model.constraint_builder = builder

constraint_number = 1
for i in list_string_total:
    model.constraint_builder.create_constraint('Con '+str(constraint_number), i)


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region creating other constraints


# endregion
# ---------------------------------------------------------------------------------------------------------------------
# region constriant builder

# endregion