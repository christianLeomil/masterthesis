import pandas as pd
import pyomo.environ as pyo
import classes 
import inspect
import hyperclass
import importlib

path_output = './output/'
path_input = './input/'
name_file = 'df_input.xlsx'

df_input_series = pd.read_excel(path_input + name_file, sheet_name='series')

#model
model = pyo.AbstractModel()

#sets
model.HOURS = pyo.Set()

# ---------------------------------------------------------------------------------------------------------------------
# region parameters for simulation

#hyperclass
df_hyper = hyperclass.define_hyper_class(path_input,path_output)

#import classes that can be used
dict_classes = pd.read_excel(path_input + name_file,sheet_name='elements')

# create instances that are going to be used (pv1, pv2 ,etc)
module_name = 'classes'
module = importlib.import_module(module_name)
list_constraints = []
list_created_classes = []
list_parameters_and_variables = []
for i in dict_classes:
    for j in range(0,dict_classes[i][0]):
        class_name = i + str(j + 1)
        globals()[class_name] = getattr(module,i)
        list_created_classes.append(class_name)

        for k in dir(globals()[class_name]):
            if k.startswith('__') or k == 'libary':
                pass
            else:
                method = getattr(globals()[class_name],k)
                list_constraints.append(method())
                list_constraints[-1] = list_constraints[-1].replace(i,class_name)
                
                method = getattr(globals()[class_name],'libary')
                libary = [element.replace(i,class_name) for element in method()]
                list_parameters_and_variables.append(libary)
list_parameters_and_variables = [item for sublist in list_parameters_and_variables for item in sublist]

# print(list_constraints)
# print('\n')
# print(list_libary)

#create table with possible connections and eliminate those that do not make sense
list_elements = ['demand','net']
list_elements = list_elements + list_created_classes

# print('\n')
# print(list_elements)

#writing connections
# df_conect = pd.DataFrame(0 , index = list_elements, columns = list_elements)
# df_conect.to_excel(path_output + 'df_conect.xlsx')

#reading connections 
df_conect = pd.read_excel(path_output+'df_conect.xlsx',index_col=0)
print(df_conect)

#Aqui entraria uma tela pra inserir a relacao entre os componentes.

list_power_from = []
list_power_to_total = []
for i in df_conect.index:
    list_power_to = []
    for j in df_conect.columns:
        if df_conect.loc[i,j] == 1:
            list_power_to.append(j)
    if list_power_to != []:
        list_power_from.append(i)
        list_power_to_total.append(list_power_to)

print(list_power_from)
print(list_power_to_total)
df_conect = pd.DataFrame({'power from': list_power_from, 'power to': list_power_to_total})
df_conect.to_excel(path_output+'df_conect_constraints.xlsx')

list_string_total = []
for i in range(0, len(df_conect)):
    string_partial = ''
    string_partial = 'model.P_' + str(df_conect['power from'].iloc[i]) + '[t]' '== model.P_' + str(
        df_conect['power to'].iloc[i][0]) + '_'+ str(df_conect['power from'].iloc[i]) + '[t]'
    for j in range(1, len(df_conect['power to'].iloc[i])):
        string_partial = string_partial + '+ model.P_' + str(df_conect['power to'].iloc[i][j]) + '_' + str(df_conect['power from'].iloc[i]) + '[t]'
    list_string_total.append(string_partial)

print('\n')
print(list_string_total)

# #criar restricoes de P_buy e P_sell, se tiver bateria P_ch e P_dis

# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region create parameters and variables

# #variables
# model.P_Pv1 = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.P_Pv1_net = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
# model.P_Pv1_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.P_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.P_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.E_buy = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.E_sell = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# model.SOC = pyo.Var(model.HOURS, within = pyo.NonNegativeReals, bounds=(0, 1))
# model.P_bat_ch = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
# model.P_bat_dis = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
# model.P_bat_net = pyo.Var(model.HOURS,within = pyo.NonNegativeReals)
# model.P_bat_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# model.K_ch = pyo.Var(model.HOURS, domain = pyo.Binary)
# model.K_dis = pyo.Var(model.HOURS, domain = pyo.Binary)

# #parameters
# model.P_solar = pyo.Param(model.HOURS) #time series with solar energy
# model.P_demand = pyo.Param(model.HOURS) #time series with solar energy
# model.costBuy = pyo.Param(model.HOURS) #time series with costs of buying energy
# model.costSell = pyo.Param(model.HOURS) #time series with price of energy being sold to grid
# model.pv_eff = pyo.Param(initialize = 0.5) #efficiency of Pv cell
# model.timestep = pyo.Param(initialize = 1) #efficiency of Pv cell

# model.starting_SOC = pyo.Param(initialize = 0.5) #starting SOC of battery
# model.E_bat_max = pyo.Param(initialize = 100) #capacity of battery
# model.bat_ch_eff = pyo.Param(initialize = 0.98) #charging efficiency of battery
# model.bat_dis_eff = pyo.Param(initialize = 0.98) #discharging efficiency of battery
# model.c_rate_ch = pyo.Param(initialize = 1) #maximal charging power (max charging power = c_rate * E_bat_max)
# model.c_rate_dis = pyo.Param(initialize = 1) #maximal discharging power (max discharging power = c_rate * E_bat_max)
 
# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region constaints

# #model.add_component acho que da pra usar esse metodo pra criar dinamicamente assim (alguma hora vou ter que fazer assim)

# # endregion 
# # ---------------------------------------------------------------------------------------------------------------------
# # region constaints generator 

# class myClass:
#     pass

# constraint_number = 1
# for i in list_constraints:
#     def dynamic_method(model, t, expr):
#         return eval(expr, globals(), locals())
#     method_name = 'Constraint_' + str(constraint_number)
#     def method_wrapper(self, model, t, expr=i):
#         return dynamic_method(model, t, expr)
#     setattr(myClass, method_name, method_wrapper)
#     my_obj = myClass()
#     setattr(model, 'constraint' + str(constraint_number), 
#             pyo.Constraint(model.HOURS,rule=getattr(my_obj, method_name)))
#     constraint_number = constraint_number + 1

# # endregion 
# # ---------------------------------------------------------------------------------------------------------------------
# # region other constraints

# def demand_rule(model,t):
#     return model.P_demand[t] == model.P_buy[t] + model.P_Pv1_demand[t]
# model.demandRule = pyo.Constraint(model.HOURS, rule = demand_rule)

# def sell_rule(model,t):
#     return model.P_sell[t] == model.P_Pv1_net[t]
# model.sellRule = pyo.Constraint(model.HOURS, rule = sell_rule)

# def con_rule(model,t):
#     return model.P_Pv1[t] == model.P_Pv1_net[t] + model.P_Pv1_demand[t]
# model.conRule = pyo.Constraint(model.HOURS, rule = con_rule)

# def buy_energy_rule(model,t):
#     return model.E_buy[t] == model.P_buy[t] * model.timestep
# model.buyEnergyRule = pyo.Constraint(model.HOURS, rule = buy_energy_rule)

# def sell_energy_rule(model,t):
#     return model.E_sell[t] == model.P_sell[t] * model.timestep
# model.sellEnergyRule = pyo.Constraint(model.HOURS, rule = sell_energy_rule)

# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region objective

# def objective_rule(model,t):
#     return sum(model.costBuy[t] * model.E_buy[t] - model.costSell[t] * model.E_sell[t] for t in model.HOURS)
# model.objectiveRule = pyo.Objective(sense = pyo.minimize,rule = objective_rule)

# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region configuring data reading for model

# data = pyo.DataPortal()
# data['HOURS'] = df_input_series['HOURS'].tolist()
# data['P_solar'] = df_input_series.set_index('HOURS')['P_solar'].to_dict()
# data['P_demand'] = df_input_series.set_index('HOURS')['P_demand'].to_dict()
# data['costBuy'] = df_input_series.set_index('HOURS')['costBuy'].to_dict()
# data['costSell'] = df_input_series.set_index('HOURS')['costSell'].to_dict()

# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region creating instance of model

# #generating instance
# instance = model.create_instance(data)

# #solving the model
# optimizer = pyo.SolverFactory('cplex')
# results = optimizer.solve(instance)

# # # Displaying the results
# # instance.pprint()
# instance.display()

# # endregion
# # ---------------------------------------------------------------------------------------------------------------------
# # region rexporting data

# df = {'HOURS':[],
#       'costBuy':[],
#       'costSell':[],
#       'P_demand':[],
#       'P_buy':[],
#       'P_solar':[],
#       'P_sell':[],
#       'P_Pv1':[],
#       'P_Pv1_demand':[],
#       'P_Pv1_net':[],
#       'E_sell':[],
#       'E_buy':[],
#       }

# for t in instance.HOURS:
#     df['HOURS'].append(str(t))
#     df['costBuy'].append(instance.costBuy[t])
#     df['costSell'].append(instance.costSell[t])
#     df['P_demand'].append(instance.P_demand[t])
#     df['P_buy'].append(instance.P_buy[t].value)
#     df['P_solar'].append(instance.P_solar[t])
#     df['P_sell'].append(instance.P_sell[t].value)
#     df['P_Pv1'].append(instance.P_Pv1[t].value)
#     df['P_Pv1_demand'].append(instance.P_Pv11_demand[t].value)
#     df['P_Pv1_net'].append(instance.P_Pv1_net[t].value)
#     df['E_sell'].append(instance.E_sell[t].value)
#     df['E_buy'].append(instance.E_buy[t].value)

# column_keys = df.keys()

# df = pd.DataFrame(df)
# df.to_excel(path_output + 'df_results_net.xlsx', index=False)

# endregion

#----------------------------------------------------------------------------------------------

# import pandas as pd
# import pyomo.environ as pyo
# import classes 
# import inspect

# path_output = './output/'
# path_input = './input/'
# name_file = 'df_input_test.xlsx'

# df_input_series = pd.read_excel(path_input + name_file, sheet_name='series')

# #model
# model = pyo.AbstractModel()

# #sets
# model.HOURS = pyo.Set()

# #parameters
# model.demand = pyo.Param(model.HOURS)
# model.cost_x = pyo.Param(model.HOURS)
# model.cost_y = pyo.Param(model.HOURS)

# #variables

# list_variables = ['quant_x','quant_y','quant_z']

# for i in list_variables:
#     setattr(model,i,pyo.Var(model.HOURS,within= pyo.NonNegativeReals))

# # model.quant_x = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
# # model.quant_y = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
# # model.quant_z = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)

# #constraints

# # class constraint_class():
# #     def demand_rule(model,t):
# #         return model.demand[t] == model.quant_x[t] + model.quant_y[t]
# #     def second_rule(model,t):
# #         return model.quant_x[t] == model.quant_z[t]

# # class Testing:
# #     def demand_rule(model,t):
# #         return model.demand[t] == model.quant_x[t] + model.quant_y[t]

# # class Subclass(Testing):
# #     def __init__(self,model,t):
# #         super().__init__(model,t)

# #     def second_rule(model,t):
# #         return model.quant_x[t] == model.quant_z[t]

# subclass_testing = classes.Subclass

# constraint_number = 1
# for name,method in inspect.getmembers(subclass_testing,predicate = inspect.isfunction):
#     if name != '__init__':
#         setattr(model,'constraint_'+str(constraint_number),pyo.Constraint(model.HOURS,rule = method))
#         constraint_number += 1

# # model.demandRule = pyo.Constraint(model.HOURS,rule = subclass_testing.demand_rule)
# # model.secondRule = pyo.Constraint(model.HOURS, rule = subclass_testing.second_rule)

# #objective
# def objective_rule(model, t):
#     return sum(model.cost_x[t] * model.quant_x[t] + model.cost_y[t] * model.quant_y[t] + model.cost_x[t] *
#             model.quant_z[t] for t in model.HOURS)

# model.objectiveRule = pyo.Objective(rule=objective_rule, sense=pyo.minimize)

# # reading data
# data = pyo.DataPortal()
# data['HOURS'] = df_input_series['HOURS'].tolist()
# data['cost_x'] = df_input_series.set_index('HOURS')['cost_x'].to_dict()
# data['cost_y'] = df_input_series.set_index('HOURS')['cost_y'].to_dict()
# data['demand'] = df_input_series.set_index('HOURS')['demand'].to_dict()

# # generating instance
# instance = model.create_instance(data)

# print("Constraint Expressions:")
# for constraint in instance.component_objects(pyo.Constraint):
#     for index in constraint:
#         print(f"{constraint}[{index}]: {constraint[index].body}")

# # solving the model
# optimizer = pyo.SolverFactory('cplex')
# results = optimizer.solve(instance)

# # # Displaying the results
# # instance.pprint()
# instance.display()
