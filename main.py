import pandas as pd
import pyomo.environ as pyo
import classes 
import inspect
import hyperclass

path_output = './output/'
path_input = './input/'
name_file = 'df_input_test.xlsx'

df_input_series = pd.read_excel(path_input + name_file, sheet_name='series')

# ---------------------------------------------------------------------------------------------------------------------
# region creating connections

#hyperclass
df_hyper = hyperclass.define_hyper_class(path_input,path_output)

#import classes that can be used
dict_classes = {'pv': 'Pv',
                'demand':'demand',
                'net':'net'}
elements = ['pv','demand','net']

for i in elements:
    pv_1 = type(i)

print(help(pv_1))

#create instances that are going to be used (pv1, pv2 ,etc)

#create table with possible connections and eliminate those that do not make sense

#criar restricoes de P_buy e P_sell, se tiver bateria P_ch e P_dis

# endregion
# ---------------------------------------------------------------------------------------------------------------------

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

# subclass_testing = Subclass

# constraint_number = 1
# for name,method in inspect.getmembers(subclass_testing,predicate=inspect.isfunction):
#     if name != '__init__':
#         setattr(model,'constraint_'+str(constraint_number),pyo.Constraint(model.HOURS,rule = method))
#     constraint_number += 1

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
