# from pyomo.environ import *

# class ExampleClass:
#     def __init__(self, param1, param2):
#         self.param1 = param1
#         self.param2 = param2

#     def func1(self, x):
#         return self.param1 * x

#     def func2(self, x, y):
#         return self.param2 * x + y

# class AbstractModel:
#     def __init__(self, example):
#         self.example = example
#         self.model = AbstractModel()

#         self.model.x = Var(domain=NonNegativeReals)
#         self.model.y = Var(domain=NonNegativeReals)
#         self.model.obj = Objective(rule=self.objective_rule)

#         self.model.constraints = ConstraintList()
#         self.model.constraints.add(self.example.func1(self.model.x) >= 10)
#         self.model.constraints.add(self.example.func2(self.model.x, self.model.y) <= 20)

#     def objective_rule(self, model):
#         return self.example.func1(model.x) + self.example.func2(model.x, model.y)

#     def solve(self):
#         solver = SolverFactory('glpk')
#         instance = self.model.create_instance()
#         results = solver.solve(instance)

#         if results.solver.termination_condition == TerminationCondition.optimal:
#             instance.solutions.load_from(results)
#             instance.display()

# example = ExampleClass(2, 3)
# abstract_model = AbstractModel(example)
# abstract_model.solve()

# #-------------------------------------------------------------------------------------------

# class DynamicClass:
#     def __init__(self, n):
#         self.n = n

#         for i in range(n):
#             def method(self, num):  # Define the method inside the loop
#                 print(f"Method {num} called!")
            
#             # Assign the method to the instance
#             setattr(self, f"method_{i}", method.__get__(self, DynamicClass))

# # Create an instance of DynamicClass
# dynamic_obj = DynamicClass(3)

# # Call the dynamically created methods
# dynamic_obj.method_0(0)
# dynamic_obj.method_1(1)
# dynamic_obj.method_2(2)

# #-------------------------------------------------------------------------------------------

# import pandas as pd

# path = './input/'
# name_file = 'c_matrix.xlsx'

# df = pd.read_excel(path + name_file)
# df.set_index('from', inplace=True)
# df.index.name = None

# class dynamic_class:
#     def __init__(self,df):
#         self.df = df

#         for i in range(0,len(df)):
#             for j in df.columns:
#                 if df[j].iloc[i] == 1:
#                     name = str(df.index[i]) + '_' + str(j)
#                     print(name)
#                     def dynamic_method(self,num):
#                         return print(num*10)       
#                     setattr(self,name,dynamic_method.__get__(self, dynamic_class))

# dynamic_obj = dynamic_class(df)
# dynamic_obj.bat_demand(4)

# method_list = [method for method in dir(dynamic_obj) if method.startswith('__') is False]
# print(method_list)


# print(eval("2 ** 8"))
# eval("pi = 3.1416")

# vector = ['pv','pv_bat','pv_demand']

# class teste1:
#     def teste(vector):
#         globals()[vector[0]] = globals()[vector[1]] +globals()[vector[2]]

# teste2 = teste1

# method_list = [method for method in dir(teste2) if method.startswith('__') is False]
# print(method_list)

# #-------------------------------------------------------------------------------------------

# from pyomo.environ import *

# model = ConcreteModel()

# # Define variables
# model.x = Var(domain=NonNegativeReals)
# model.y = Var(domain=NonNegativeReals)

# # Define a parameter
# param = 5

# # Define a dynamic expression
# # model.dynamic_expr = Expression(expr=model.x + model.y <= param)

# # Define a constraint using the dynamic expression
# model.constraint = Constraint(expr = model.x + model.y <= param)

# model.obj = Objective(expr=model.x + model.y, sense=minimize)

# # Solve the model
# solver = SolverFactory('cplex')
# results = solver.solve(model)

# # Print the results
# print("x =", model.x.value)
# print("y =", model.y.value)


#------------------------------------------------------------------------------------------------------------------------------------

# vetor = ['pv','bat','net','demand']

# def create_equation(vetor):
#     return globals()[vetor[0]] == sum(globals()[i] for i in vetor)

# from pyomo.environ import AbstractModel, Constraint, Var, Objective, Param, minimize, SolverFactory, value

# class ConstraintBuilder:
#     def __init__(self, model):
#         self.model = model

#     def create_constraint(self, constraint_name, expression_str):
#         expression = eval(expression_str, globals(), locals())
#         setattr(self, constraint_name, expression)
#         constraint = Constraint(expr=expression)
#         setattr(self.model, constraint_name, constraint)

# # Create an abstract model
# model = AbstractModel()

# model.Teste = Param(initialize = 100)

# # Define variables
# model.x = Var()
# model.y = Var()

# # Create the constraint builder
# builder = ConstraintBuilder(model)

# # Add constraint building code to the model
# model.constraint_builder = builder

# # Dynamically create a constraint method with a dynamically built expression
# expression_str = "model.x + model.y >= model.Teste"
# model.constraint_builder.create_constraint("constraint1", expression_str)

# # Dynamically create another constraint method with a dynamically built expression
# expression_str = "model.x - model.y <= 2"
# model.constraint_builder.create_constraint("constraint2", expression_str)

# # Define the objective
# model.obj = Objective(expr=model.x + 2 * model.y, sense=minimize)

# # Create an instance of the model
# instance = model.create_instance()

# # Solve the instance
# solver = SolverFactory('cplex')
# solver.solve(instance)

# # Access and check the created constraints
# print('Constraint 1:', instance.constraint1())
# print('Constraint 2:', instance.constraint2())

# # Access and check the objective value
# print('Objective value:', value(instance.obj))


#------------------------------------------------------------------------------------------------------------------------------------

# class myClass:
#     pass

# list_constraints = ['string 1','string 2','string 3']

# constraint_number = 1
# for i in list_constraints:

#     def dynamic_function(x1,x2):
#         return x1 + x2
    
#     method_name = 'Constraint_' +str(constraint_number)
#     method = dynamic_function(model,t)
#     setattr(myClass,method_name,method)

#     my_obj = myClass()

#     constraint_number = constraint_number +1

# print(my_obj.Constraint_1(1,10))

#-------------------------------------------------------------------------------------------------------

# list_constraints = ['string 1', 'string 2', 'string 3']

# class myClass:
#     pass

# constraint_number = 1
# for i in list_constraints:
#     def dynamic_function(x1, x2):
#         return x1 + x2
    
#     method_name = 'Constraint_' + str(constraint_number)
    
#     def method_wrapper(self,x1, x2):
#         return dynamic_function(x1, x2)
    
#     setattr(myClass, method_name, method_wrapper)

#     my_obj = myClass()

#     constraint_number += 1

# print(my_obj.Constraint_3('a', 'b'))

# method_list = [method for method in dir(myClass) if method.startswith('__') is False]
# print(method_list)


#-------------------------------------------------------------------------------------------------------

# import pandas as pd
# import pyomo.environ as pyo

# class MyModel:
#     def __init__(self):
#         self.model = pyo.AbstractModel()

#     def define_sets(self):
#         self.model.HOURS = pyo.Set()

#     def define_parameters(self):
#         self.model.time_step = pyo.Param()
#         self.model.starting_SOC = pyo.Param()
#         self.model.E_bat_max = pyo.Param()
#         self.model.bat_ch_eff = pyo.Param()
#         self.model.bat_dis_eff = pyo.Param()

#     def define_variables(self):
#         self.model.SOC = pyo.Var(self.model.HOURS, within=pyo.NonNegativeReals, bounds=(0, 1))
#         self.model.P_bat_ch = pyo.Var(self.model.HOURS, within=pyo.NonNegativeReals)
#         self.model.P_bat_dis = pyo.Var(self.model.HOURS, within=pyo.NonNegativeReals)

#     def define_constraints(self):
#         def battery_rule(model, t):
#             if t == 1:
#                 return model.SOC[t] == model.starting_SOC
#             else:
#                 return model.SOC[t] == model.SOC[t-1] + (model.P_bat_ch[t-1] * model.bat_ch_eff -
#                                                          model.P_bat_dis[t-1] / model.bat_dis_eff) * \
#                        model.time_step / model.E_bat_max

#         self.model.batteryRule = pyo.Constraint(self.model.HOURS, rule=battery_rule)

#     def solve_model(self, df_input_series, df_input_other, df_power_constraints):
#         # Generate the model and solve it
#         self.define_sets()
#         self.define_parameters()
#         self.define_variables()
#         self.define_constraints()

#         # ... (rest of the code to read data, create instance, and solve the model)
#         # ... (including the code to export the model results)

# # Example usage
# model = MyModel()
# model.solve_model(df_input_series, df_input_other, df_power_constraints)


#-------------------------------------------------------------------------------------------------------

# import pandas as pd
# import pyomo.environ as pyo
# from classes import Subclass
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

# class Testing:
#     def demand_rule(model,t):
#         return model.demand[t] == model.quant_x[t] + model.quant_y[t]

# class Subclass(Testing):
#     def __init__(self,model,t):
#         super().__init__(model,t)

#     def second_rule(model,t):
#         return model.quant_x[t] == model.quant_z[t]

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

#-------------------------------------------------------------------------------------------------------
# class MyClass:
#     def __init__(self, value):
#         self.value = value

# # Define a template expression with custom variables
# template_expression = "self.value * {factor} + {constant}"

# # Define a function to create a new method based on the template expression
# def create_dynamic_method(factor, constant):
#     def dynamic_method(self):
#         expression = template_expression.format(factor=factor, constant=constant)
#         result = eval(expression)
#         print("Dynamic method:", result)
#     return dynamic_method

# # Create new methods dynamically based on the template expression
# factors = [2, 3, 4]
# constants = [5, 10, 15]

# for factor, constant in zip(factors, constants):
#     method_name = f"dynamic_method_{factor}_{constant}"
#     new_method = create_dynamic_method(factor, constant)
#     setattr(MyClass, method_name, new_method)

# # Create an instance of the class
# my_instance = MyClass(7)

# # Call the dynamically created methods
# my_instance.dynamic_method_2_5()
# my_instance.dynamic_method_3_10()
# my_instance.dynamic_method_4_15()

#-------------------------------------------------------------------------------------------------------
original_list = ['apple', 'banana', 'orange', 'kiwi']
substring_list = ['an', 'ra']

filtered_list = [item for item in original_list if not any(substring in item for substring in substring_list)]

print(filtered_list)

