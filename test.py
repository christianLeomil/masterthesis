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


#-------------------------------------------------------------------------------------------

from pyomo.environ import *

model = ConcreteModel()

# Define variables
model.x = Var(domain=NonNegativeReals)
model.y = Var(domain=NonNegativeReals)

# Define a parameter
param = 5

# Define a dynamic expression
model.dynamic_expr = Expression(expr=model.x + model.y <= param)

# Define a constraint using the dynamic expression
model.constraint = Constraint(expr=model.dynamic_expr)

# Solve the model
solver = SolverFactory('glpk')
results = solver.solve(model)

# Print the results
print("x =", model.x.value)
print("y =", model.y.value)

