import inspect
import pyomo.environ as pyo
import pandas as pd

#-------------------------------------------------------------------------------------------------------------
#region creating string for constraints

class MyClass:
    def method_1(self, model, t):
        return model.demand[t] == model.quant_x[t] + model.quant_z[t]

myObj = MyClass()

# Get the original method
original_method = myObj.method_1

# print(original_method)

# Get the source code of the method
source_code = inspect.getsource(original_method)

print(source_code)

# Replace the parameter name
modified_source_code = source_code.replace("quant_z", "quant_y")

print(modified_source_code)

# modified_source_code = 'def method_1(self, model, t): \n return model.demand[t] == model.quant_x[t] + model.quant_y[t]'

return_index = modified_source_code.find("return")
modified_source_code = modified_source_code[return_index + len("return"):].strip()

print(modified_source_code)

list_expressions =[modified_source_code]

# endregion
# -------------------------------------------------------------------------------------------------------------
# region creating model

model = pyo.AbstractModel()

model.HOURS = pyo.Set()

model.demand = pyo.Param(model.HOURS)
model.cost_x = pyo.Param(model.HOURS)
model.cost_y = pyo.Param(model.HOURS)

model.quant_x = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
model.quant_y = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)
model.quant_z = pyo.Var(model.HOURS, within=pyo.NonNegativeReals)

#endregion
#-------------------------------------------------------------------------------------------------------------
#region creating constriant

class myClass:
    pass

constraint_number = 1
for i in list_expressions:
    def dynamic_method(model,t,expr):
        return eval(expr, globals(), locals())
    method_name = 'Constraint_' + str(constraint_number)

    def method_wrapper(self, model,t,expr=i):
        return dynamic_method(model,t,expr)
    setattr(myClass, method_name, method_wrapper)
    my_obj = myClass()
    setattr(model, 'Constraint_' + str(constraint_number), pyo.Constraint(model.HOURS,
                                                                             rule=getattr(my_obj, method_name)))
    constraint_number = constraint_number + 1

#endregion
#-------------------------------------------------------------------------------------------------------------
#region objetive function

def objective_rule(model,t):
    return sum(model.quant_x[t] * model.cost_x[t] + model.quant_y[t] * model.cost_y[t] for t in model.HOURS)
model.objectiveRule =pyo.Objective(rule = objective_rule, sense = pyo.minimize)

#endregion
#-------------------------------------------------------------------------------------------------------------
#region executing the model

path_input = './input/'
path_output = './output/'
name_file = 'df_input_test.xlsx'

df_input_series = pd.read_excel(path_input + name_file, sheet_name = 'series')

# reading data
data = pyo.DataPortal()
data['HOURS'] = df_input_series['HOURS'].tolist()
data['cost_x'] = df_input_series.set_index('HOURS')['cost_x'].to_dict()
data['cost_y'] = df_input_series.set_index('HOURS')['cost_y'].to_dict()
data['demand'] = df_input_series.set_index('HOURS')['demand'].to_dict()

# generating instance
instance = model.create_instance(data)

print("Constraint Expressions:")
for constraint in instance.component_objects(pyo.Constraint):
    for index in constraint:
        print(f"{constraint}[{index}]: {constraint[index].body}")

# solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
instance.display()

#endregion