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

import pandas as pd
import pyomo.environ as pyo

path_output = './output/'
path_input = './input/'
name_file = 'df_input_test.xlsx'

df_input_series = pd.read_excel(path_input + name_file,sheet_name = 'series')
df_conect = pd.read_excel(path_input + 'c_matrix.xlsx', sheet_name = 'test')

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
    string_partial ='model.' + str(df_conect['energy from'].iloc[i]) + '[t]' '== model.' +str(df_conect['energy to'].iloc[i][0]) + '[t]'
    for j in range(1,len(df_conect['energy to'].iloc[i])):
        string_partial = string_partial + '+ model.' + str(df_conect['energy to'].iloc[i][j])+ '[t]'
    list_string_total.append(string_partial)
print(list_string_total)

model = pyo.AbstractModel()

model.HOURS = pyo.Set()

model.demand = pyo.Param(model.HOURS)
model.cost_x = pyo.Param(model.HOURS)
model.cost_y = pyo.Param(model.HOURS)

model.quant_x = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.quant_y = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)
model.quant_z = pyo.Var(model.HOURS, within = pyo.NonNegativeReals)

constraint_number = 1
for i in list_string_total:

    def dynamic_method(model,t):
        return eval(i, globals(), locals())
    
    setattr(model, 'constraint' + str(constraint_number), pyo.Constraint(model.HOURS,rule = dynamic_method))
    constraint_number = constraint_number + 1


# def test_function(model,t):
#     return model.demand[t] == model.quant_x[t] + model.quant_y[t]
# model.testFunction = pyo.Constraint(model.HOURS,rule = test_function)

# def constraint_other(model,t):
#     return model.quant_z[t] == model.quant_x[t]
# model.constraintOther = pyo.Constraint(model.HOURS,rule = constraint_other)

def objective_rule(model,t):
    return sum(model.cost_x[t] * model.quant_x[t] + model.cost_y[t] * model.quant_y[t] + model.cost_x[t] * model.quant_z[t] for t in model.HOURS)
model.objectiveRule = pyo.Objective(rule = objective_rule, sense = pyo.minimize)

#reading data
data = pyo.DataPortal()
data['HOURS'] = df_input_series['HOURS'].tolist()
data['cost_x'] = df_input_series.set_index('HOURS')['cost_x'].to_dict()
data['cost_y'] = df_input_series.set_index('HOURS')['cost_y'].to_dict()
data['demand'] = df_input_series.set_index('HOURS')['demand'].to_dict()

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
      'demand':[],
      'cost_x':[],
      'cost_y':[],
      'quant_x':[],
      'quant_y':[],
      'quant_z':[]
      }

for t in instance.HOURS:
    df['HOURS'].append(str(t))
    df['demand'].append(instance.demand[t])
    df['cost_x'].append(instance.cost_x[t])
    df['cost_y'].append(instance.cost_y[t])
    df['quant_x'].append(instance.quant_x[t].value)
    df['quant_y'].append(instance.quant_y[t].value)
    df['quant_z'].append(instance.quant_z[t].value)

df = pd.DataFrame(df)
df.to_excel(path_output + 'df_results_test.xlsx', index=False)


from pyomo.core import TransformationFactory

trans = TransformationFactory('core.relax_integrality')
trans.apply_to(instance)

print(instance.constraint1.get_body_expr())
print(instance.constraint2.get_body_expr())
