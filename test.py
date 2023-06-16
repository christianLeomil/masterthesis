import inspect
import pyomo.environ as pyo
import pandas as pd

#-------------------------------------------------------------------------------------------------------------
#region creating string for constraints

class MyClass:
    def method_1(self, model, t):
        return model.P_pv[t] == model.P_pv_bat[t] + model.P_pv_bat[t]

myObj = MyClass()

# Get the original method
original_method = myObj.method_1

# print(original_method)

# Get the source code of the method
source_code = inspect.getsource(original_method)

print(source_code)

# Replace the parameter name
modified_source_code = source_code.replace("P_pv", "P_pv1")

print(modified_source_code)

modified_source_code = 'def method_1(self, model, t): \n return model.demand[t] == model.quant_x[t] + model.quant_y[t]'

# Define the new class with the modified method
class NewClass:
    exec(modified_source_code)

# Create an instance of the new class
newObj = NewClass()

print(inspect.getsource(NewClass))

#endregion
#-------------------------------------------------------------------------------------------------------------
#region creating model

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

method = getattr(NewClass,'method_1')
name = 'constraint_test'
model.add_component(name, pyo.Constraint(model.HOURS, rule = method))

def objective_rule(model,t):
    return sum(model.quant_x[t] * model.cost_x + model.quant_y[t] * model.cost_y[t] for t in model.HOURS)
model.objectiveRule =  pyo.Objective(rule=objective_rule,sense = pyo.minimize)

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