# import pandas as pd
# import pyomo.environ as pyo

# #model
# model = pyo.AbstractModel()

# #sets
# model.HOURS = pyo.Set()
# model.PV = pyo.Set()
# model.BAT = pyo.Set()

# #varables
# model.P_pv_demand = pyo.Var(model.HOURS, model.PV,within = pyo.NonNegativeReals)
# model.P_bat_demand = pyo.Var(model.HOURS, model.BAT,within = pyo.NonNegativeReals)
# model.P_net_demand = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

# name_file = 'df_input.xlsx'
# path_output = './output/'
# path_input = './input/'

# df_conect_to  = pd.read_excel(path_input + name_file, sheet_name = 'connect_to',index_col=0)
# df_conect_to.index.name = None

# list_energy_from = []
# list_energy_to_total = []
# for i in range(0, len(df_conect_to)):
#     list_energy_to = []
#     for j in df_conect_to.columns:
#         if df_conect_to[j].iloc[i] == 1:
#             list_energy_to.append(j)
#     if list_energy_to != []:
#         list_energy_to_total.append(list_energy_to)
#         list_energy_from.append(df_conect_to.index[i])

# df_conect_to = pd.DataFrame({'from': list_energy_from, 'to': list_energy_to_total})
# print(df_conect_to)
# print('\n')

# df_sets = pd.DataFrame({'name':['pv','bat'],
#                           'var':['n','m'],
#                           'set':['model.PV','model.BAT']})

# list_expressions = []
# list_sets = []
# for i in df_conect_to.index:
#     list_expression_conect =[]
#     if df_conect_to['from'].iloc[i] in ['demand','net']:
#         list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t] == ')
#         list_sets.append('')
#     else:
#         list_expression_conect.append('model.P_to_' + df_conect_to['from'].iloc[i] + '[t][' + df_sets['var'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0] + '] == ')
#         list_sets.append(df_sets['set'][df_sets['name'] == df_conect_to['from'].iloc[i]].iloc[0])
#     for index,j in enumerate(df_conect_to['to'].iloc[i]):
#         if index != 0:
#             list_expression_conect[-1] = list_expression_conect[-1] + ' +'
#         if j not in ['demand','net']:
#             list_expression_conect[-1] = list_expression_conect[-1] + ' sum(model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t][' + str(df_sets['var'][df_sets['name']==j].iloc[0]) +'] for ' + str(df_sets['var'][df_sets['name']==j].iloc[0]) + ' in ' + str(df_sets['set'][df_sets['name']==j].iloc[0]) + ')'
#         else:
#             list_expression_conect[-1] = list_expression_conect[-1] + ' model.P_' + j + '_' + df_conect_to['from'].iloc[i] +'[t]'
#     list_expressions.append(list_expression_conect)

# print(list_expressions)
# print('\n')
# print(list_sets)


# #creating constraints from connections
# class myClass:
#     pass

# constraint_number = 1
# for i in list_expressions:
#     def dynamic_method(model,t,n,m,expr):
#         return eval(expr, globals(), locals())

#     method_name = 'Constraint_' + str(constraint_number)

#     def method_wrapper(self, model, t, n,m,expr=i):
#         return dynamic_method(model,t,n,m,expr)

#     setattr(myClass, method_name, method_wrapper)

#     my_obj = myClass()

#     setattr(model, 'constraint_conect' + str(constraint_number), pyo.Constraint(model.HOURS,model.PV,
#                                                                                 model.BAT,rule=getattr(my_obj, method_name)))
#     constraint_number = constraint_number + 1

# class MyClass:
#     pass

# expression = 'model.P_to_demand[t] == model.P_pv_demand[t,n] + model.P_bat_demand[t] + model.P_net_demand[t]'

# def demand_rule(model,t,n,m):
#     return model.P_to_demand[t] == model.P_pv_demand[t,n] + model.P_bat_demand[t] + model.P_net_demand[t]

'--------------------------------------------------------------------------------------------------------------'

# import inspect

# class MyClass:
#     pass

# expression = 'model.P_to_demand[t] == model.P_pv_demand[t,n] + model.P_bat_demand[t] + model.P_net_demand[t]'

# def demand_rule(self, model, t, n):
#     return eval(expression)

# setattr(MyClass, 'demand_rule', demand_rule)


# my_object = MyClass()

# methods = dir(my_object)

# for method_name in methods:
#     # Check if the item is a method
#     if callable(getattr(my_object, method_name)):
#         method = getattr(my_object, method_name)
#         if method.__name__ == 'demand_rule':
#             method_source = inspect.getsource(method)
#             print(f"\n{method_name}:")
#             print(method_source)


import pandas as pd

# Create a sample DataFrame
data = {'A': [0, 0, 0, 0, 0],
        'B': [0, 2, 0, 0, 0],
        'C': [0, 1, 0, 0, 0]}
df = pd.DataFrame(data)

# Specify the line number to check (e.g., line 2)
line_number = 7

# Check if the specified line has any component different from zero
line_has_non_zero = (df.loc[line_number] != 0).any()

print(line_has_non_zero)
