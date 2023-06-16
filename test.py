import inspect
import pyomo.environ as pyo
import pandas as pd

#-------------------------------------------------------------------------------------------------------------
#region creating string for constraints
    
import inspect

class MyClass:
    pass

# Define the lambda function
# method_1 = lambda model, t: model.demand[t] == model.quant_x[t] + model.quant_z[t]/2 if t < 10 else model.demand[t] == model.quant_x[t] + model.quant_z[t]
method_1 = lambda model, t: model.demand[t] == model.quant_x[t] + model.quant_z[t]

# Assign the lambda function to method_1 in MyClass
setattr(MyClass, "method_1", method_1)

myObj = MyClass()

# Get the original method
original_method = myObj.method_1

# Get the source code of the method
source_code = inspect.getsource(original_method)

# Replace the parameter name
modified_source_code = source_code.replace("model.quant_z[t]", "model.quant_y[t] + model.quant_z[t]")

print(modified_source_code)

# Compile the modified source code
compiled_code = compile(modified_source_code, "<string>", "exec")

# Create a namespace dictionary for execution
namespace = {}

# Execute the compiled code in the namespace
exec(compiled_code, namespace)

# Get the modified method from the namespace
modified_method = namespace["method_1"]

# Set the modified method as the new method_1
setattr(MyClass, "method_1", modified_method)

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

model.add_component('Constraint1', pyo.Constraint(model.HOURS, rule = MyClass.method_1))

#endregion
#-------------------------------------------------------------------------------------------------------------
#region objetive function

def objective_rule(model,t):
    return sum(model.quant_x[t] * model.cost_x[t] + model.quant_y[t] * model.cost_y[t] + model.quant_z[t] * model.cost_x[t] for t in model.HOURS)
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