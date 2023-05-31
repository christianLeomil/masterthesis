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

class myClass:
    pass

constraint_number = 1
for i in list_string_total:

    def dynamic_method(model,t):
        return eval(i, globals(), locals())
    
    method_name = 'Constraint_' +str(constraint_number)

    def method_wrapper(self,model,t):
        return dynamic_method(model,t)

    setattr(myClass,method_name,method_wrapper)
    
    my_obj = myClass()
    
    setattr(model, 'constraint' + str(constraint_number), pyo.Constraint(model.HOURS,rule = getattr(my_obj,method_name)))
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

print("Constraint Expressions:")
for constraint in instance.component_objects(pyo.Constraint):
    for index in constraint:
        print(f"{constraint}[{index}]: {constraint[index].body}")

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

method_list = [method for method in dir(myClass) if method.startswith('__') is False]
print(method_list)