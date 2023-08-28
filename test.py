import pandas as pd
import sys
import inspect
import classes

path_input = './input/'
control = getattr(classes, 'control')(path_input)
print(control.df)

class PV:
    def __init__(self,name_of_instance):
        self.name_of_instance = name_of_instance

        self.list_var = ['pv_op_cost','pv_emissions','pv_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_param = ['pv_eff','pv_area','pv_spec_op_cost','pv_spec_em','pv_kWp_per_area',
                           'pv_inv_per_kWp']
        
        self.list_text_param = ['','','','','','']

        self.list_series = ['E_pv_solar']
        self.list_text_series =['model.HOURS']

        #default values in case of no input
        self.pv_eff = 0.15 # aproximate overall efficiency of pv cells 
        self.pv_area = 100 # m^2
        self.pv_spec_op_cost = 0.01 # cost per hour per m^2 area of pv installed
        self.pv_kWp_per_area  = 0.12 # kWP per m2 of PV
        self.pv_inv_per_kWp = 1000 # EURO per kWp
        self.pv_life_time = 30 #lifetime of panels in years
        self.pv_spec_em = 0.50 #kgCO2eq/kWh generated

        #default series
        self.E_pv_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        self.super_class = 'generator'

    def constraint_generation_rule(model,t):
        return model.P_from_pv[t] == (model.E_pv_solar[t] / model.time_step) * model.pv_eff * model.pv_area 
    
    def constraint_operation_costs(model,t):
        return model.pv_op_cost[t] == model.pv_area * model.pv_spec_op_cost
    
    def constraint_emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.pv_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pv_inv_cost[t] == model.pv_area * model.pv_kWp_per_area * model.pv_inv_per_kWp
        else:
            return model.pv_inv_cost[t] == 0

class Solar_Th:
    def __init__(self,name_of_instance):
        self.name_of_instance = name_of_instance

        self.list_var = ['solar_th_op_cost','solar_th_emissions','solar_th_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_param = ['solar_th_eff','solar_th_area','solar_th_spec_op_cost','solar_th_spec_em',
                           'solar_th_inv_per_area']
        self.list_text_param = ['','','','','']

        self.list_series = ['E_solar_th_solar']
        self.list_text_series =['model.HOURS']

        #default values in case of no input
        self.solar_th_eff = 0.20 # aproximate overall efficiency of pv cells 
        self.solar_th_area = 50 # m^2
        self.solar_th_spec_op_cost = 0.02 # cost per hour per m^2 area of pv installed
        self.solar_th_spec_em = 0.50 #kgCO2eq/kWh, same value assumed as for PV
        self.pv_life_time = 15 #lifetime of panels in years
        self.solar_th_inv_per_area = 700 #EURO per m2 aperture

        #default series
        self.E_solar_th_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'no',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
        
    def constraint_generation_rule(model,t):
        return model.Q_from_solar_th[t] == (model.E_solar_th_solar[t] * model.time_step) * model.solar_th_area * model.solar_th_eff
        
    def constraint_operation_cost(model,t):
        return model.solar_th_op_cost[t] == model.solar_th_area * model.solar_th_spec_op_cost
        
    def constraint_emission(model,t):
        return model.solar_th_emissions[t] == model.Q_from_solar_th[t] * model.solar_th_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.solar_th_inv_cost[t] == model.solar_th_area * model.solar_th_inv_per_area
        else:
            return model.solar_th_inv_cost[t] == 0

path_input = './input/'

pv = PV('pv')
solar_th = Solar_Th('solar_th')

list_elements = ['solar_th','pv']

list_altered = []
if control.df.loc['size_optimization','value'] == 'yes':
    for i in list_elements:
        methods = inspect.getmembers(globals()[i])
        for method_name, method_value in methods:
            if method_name == 'list_param':
                list_altered = list_altered + method_value

df = pd.DataFrame({'list_altered':list_altered})
df['choice'] = 0
df['upper value'] = 0
df['lower value'] = 0

with pd.ExcelWriter(path_input + 'df_input.xlsx',mode = 'a', engine = 'openpyxl',if_sheet_exists= 'replace') as writer:
    df.to_excel(writer,sheet_name = 'parameters to variables',index = False)
input('Please select components that are going to be otpimized and press enter...')

df = pd.read_excel(path_input + 'df_input.xlsx',sheet_name = 'parameters to variables',index_col = 0)
df.index.name = None
print('\n')
print(df)

list_altered_variables = []
list_upper_value = []
list_lower_value = []

for i in range(0,len(df)):
    if df['choice'].iloc[i] == 1:
        list_altered_variables.append(df.index[i])
        list_upper_value.append(df['upper value'].iloc[i])
        list_lower_value.append(df['lower value'].iloc[i])

for i in list_elements:
    list_original_param = getattr(globals()[i],'list_param')
    list_original_text_param = getattr(globals()[i],'list_text_param')
    list_original_var = getattr(globals()[i],'list_var')
    list_original_text_var = getattr(globals()[i],'list_text_var')

    print('\n')
    print(list_original_param)
    print(list_altered_variables)
    for ind,j in enumerate(list_altered_variables):
        try:
            index = list_original_param.index(j)
            list_original_param.pop(index)
            list_original_text_param.pop(index)
            list_original_var.append(j)

            text ='within = NonNegativeReals, bounds = ('+str(list_lower_value[ind])+','+str(list_upper_value[ind])+')'
            list_original_text_var.append(text) 
        except Exception:
            pass
    print('list_original_parameters')
    print(list_original_param)
    setattr(globals()[i],'list_param',list_original_param)
    setattr(globals()[i],'list_text_param',list_original_text_param)
    setattr(globals()[i],'list_var',list_original_var)
    setattr(globals()[i],'list_text_var',list_original_text_var)

for i in list_elements:
    print('\n '+ i)
    methods = inspect.getmembers(globals()[i])
    for method_name, method_value in methods:
        if method_name in ['list_param','list_text_param','list_var','list_text_var']:
            print(method_name)
            print(method_value)






