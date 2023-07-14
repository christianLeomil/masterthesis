import inspect
import pyomo.environ as pyo
import pandas as pd
import textwrap
#-------------------------------------------------------------------------------------------------------------
#region creating string for constraints
    
import inspect

# class MyClass:
#       def extra_rule(model,t):
#             if t <= 10:
#                   return model.quant_z[t] <= 200
#             else:
#                   return model.quant_z[t] <= 100
      
#       def first_rule(model,t):
#             return model.demand[t] == model.quant_x[t] + model.quant_z[t]

class MyClass:
      def first_constraint(model,t,var_1,var_2,var_3):
            if t == 10:
                  return model.globals()[var_1][t] <= 200
            else:
                  return model.globals()[var_1][t] <= 100
            
      def second_rule(model,t,var_1,var_2,var_3):
            return model.globals()[var_1] == model.globals()[var_2] + model.globals()[var_3]
      
myObj = MyClass()



myObj = MyClass()

for i in dir(myObj):
        if i.startswith('__'):
             pass
        else:
                # Get the original method
                original_method = getattr(myObj,i)
                print(original_method)
                # Get the source code of the method
                source_code = inspect.getsource(original_method)
                source_code = textwrap.dedent(source_code)
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
                modified_method = namespace[i]
                # Set the modified method as the new method_1
                setattr(MyClass, i, modified_method)

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

model.test = pyo.Var(model.HOURS, within = pyo.PositiveIntegers)

#endregion
#-------------------------------------------------------------------------------------------------------------
#region creating constriant

constraint_num = 1
for i in dir(MyClass):
      if i.startswith('__'):
            pass
      else:
           method = getattr(MyClass,i)
           model.add_component('Constraint'+ str(constraint_num), pyo.Constraint(model.HOURS, rule = method))
           constraint_num += 1

#endregion
#-------------------------------------------------------------------------------------------------------------
#region objetive function
class contains_objective:
      def objective_rule(model,t):
            return sum(model.quant_x[t] * model.cost_x[t] + model.quant_y[t] * model.cost_y[t] + model.quant_z[t] * model.cost_x[t] for t in model.HOURS)
model.objectiveRule =pyo.Objective(rule = contains_objective.objective_rule, sense = pyo.minimize)

#endregion
#-------------------------------------------------------------------------------------------------------------
#region executing the model

path_input = './input/'
path_output = './output/'
name_file = 'df_input_new_model.xlsx'

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

# #parameter list
# parameter_list = list(model.component_data_objects(pyo.Param, descend_into=True))

# print('\n')
# # Print the parameter names
# for param in parameter_list:
#     print(param.name)
# print('\n')

# solving the model
optimizer = pyo.SolverFactory('cplex')
results = optimizer.solve(instance)

# # Displaying the results
# instance.pprint()
instance.display()

#endregion