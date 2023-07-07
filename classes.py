# class Generator:
#     def __init__(self,type_, id,eff,E_in,op_cost,inv_cost,emission):
#         self.class_type = 'generator'
#         self.type_ = type_
#         self.id = id
#         self.eff = eff
#         self.E_in= E_in
#         self.inv_cost = inv_cost
#         self.op_cost = op_cost
#         self.emission = emission

class pv():
    def __init__(self):
        self.list_var = ['pv_op_cost','pv_emissions'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']
        self.list_param = ['pv_eff','pv_area','pv_spec_op_cost','pv_spec_em']
        self.list_text_param = ['','','','']
        self.list_series = ['P_pv_solar']
        self.list_text_series =['model.HOURS']

        #default values in case of no input
        self.pv_eff = 0.15 # aproximate overall efficiency of pv cells 
        self.pv_area = 100 # m^2
        self.pv_spec_op_cost = 0.01 # cost per hour per m^2 area of pv installed
        self.P_pv_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given
        self.pv_spec_em = 0 #There is no CO2 emission from generation energy with PV


        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        self.super_class = 'generator'

    def solar_rule(model,t):
        return model.P_from_pv[t] == (model.P_pv_solar[t] / model.time_step) * model.pv_eff * model.pv_area 
    
    def operation_costs(model,t):
        return model.pv_op_cost[t] == model.pv_area * model.pv_spec_op_cost
    
    def emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.pv_spec_em
    
class bat:
    def __init__(self):
        self.list_var = ['bat_SOC','bat_K_ch','bat_K_dis','bat_op_cost','bat_emissions'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',
                              'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']
        self.list_param = ['bat_starting_SOC','bat_ch_eff','bat_dis_eff','bat_E_max',
                           'bat_c_rate_ch','bat_c_rate_dis','bat_spec_op_cost',
                           'bat_spec_em']
        self.list_text_param = ['','','','','','','','']
        self.list_series = []
        self.list_text_series = []

        #default values in case of no input
        self.bat_starting_SOC = 0.5
        self.bat_ch_eff = 0.95
        self.bat_dis_eff = 0.95
        self.bat_E_max = 100
        self.bat_c_rate_ch = 1
        self.bat_c_rate_dis = 1
        self.bat_spec_op_cost = 0.01
        self.bat_spec_em = 0 # no emission for operating batteries

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        
        self.super_class = 'transformer'

    def function_rule(model,t):
            if t == 1:
                return model.bat_SOC[t] == model.bat_starting_SOC + (model.P_to_bat[t] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.bat_dis_eff) * model.time_step / model.bat_E_max
            else:
                return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_to_bat[t] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.bat_dis_eff) * model.time_step / model.bat_E_max

    def charge_limit(model,t):
        return model.P_to_bat[t] <= model.bat_E_max * model.bat_K_ch[t] * model.bat_c_rate_ch
    
    def discharge_limit(model,t):
        return model.P_from_bat[t] <= model.bat_E_max * model.bat_K_dis[t] * model.bat_c_rate_dis
    
    def keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1
    
    def operation_costs(model,t):
        return model.bat_op_cost[t] == model.bat_E_max * model.bat_spec_op_cost
    
    def emissions(model,t):
        return model.bat_emissions[t] == (model.P_from_bat[t] + model.P_to_bat[t]) * model.bat_spec_em
    
class demand:
    def __init__(self):
        self.list_var = [] #no powers
        self.list_text_var = []
        self.list_param = []
        self.list_text_param = []
        self.list_series = ['P_to_demand','Q_to_demand']
        self.list_text_series =['model.HOURS','model.HOURS']

        #default values in case of no input
        self.P_to_demand = 500 # this is tranformed into a series in case user does not give input
        self.Q_to_demand = 1000 # this is tranformed into a series in case user does not give input

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'demand'

class net:
    def __init__(self):
        self.list_var = ['net_sell_electric','net_buy_electric','net_sell_thermal','net_buy_thermal'
                         ,'net_emissions'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals']
        self.list_param = ['net_spec_em_P','net_spec_em_Q']
        self.list_text_param = ['','']
        self.list_series = ['net_cost_buy_electric','net_cost_sell_electric',
                            'net_cost_buy_thermal','net_cost_sell_thermal']
        self.list_text_series =['model.HOURS','model.HOURS',
                                'model.HOURS','model.HOURS']

        #default values in case of no input
        self.net_cost_buy_electric = 0.4
        self.net_cost_sell_electric = 0.3
        self.net_cost_buy_thermal = 0.2
        self.net_cost_sell_thermal = 0.1
        self.net_spec_em_P = 0.56 # kg of CO2 per kWh
        self.net_spec_em_Q = 0.24 # kg of CO2 per kWh

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'external net'
        
    def sell_energy_electric(model,t):
        return model.net_sell_electric[t] == model.P_to_net[t] * model.time_step * model.net_cost_sell_electric[t]
    
    def buy_energy_electric(model,t):
        return model.net_buy_electric[t] == model.P_from_net[t] * model.time_step * model.net_cost_buy_electric[t]
    
    def sell_energy_thermal(model,t):
        return model.net_sell_thermal[t] == model.Q_to_net[t] * model.time_step * model.net_cost_sell_thermal[t]
    
    def buy_energy_thermal(model,t):
        return model.net_buy_thermal[t] == model.Q_from_net[t] * model.time_step * model.net_cost_buy_thermal[t]
    
    def emissions(model,t):
        return model.net_emissions[t] == model.P_from_net[t] * model.net_spec_em_P + model.Q_from_net[t] * model.net_spec_em_Q 

class CHP:
    def __init__(self):
        self.list_var = ['CHP_fuel_cons','CHP_op_cost','CHP_emissions'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']
        self.list_param = ['P_CHP_max','P_CHP_min','CHP_P_to_Q_ratio','CHP_fuel_cons_ratio','CHP_fuel_price',
                           'CHP_spec_em']
        self.list_text_param = ['','','','','','']
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input
        self.P_CHP_max = 2000 #W electric
        self.P_CHP_min = 0
        self.CHP_P_to_Q_ratio = 0.5 
        self.CHP_fuel_cons_ratio = 0.105 #dm3 per kWh of P_from_CHP
        self.CHP_fuel_price = 5 # EUROS/dm3 of fuel 
        self.CHP_spec_em = 2.3 # kg of CO2 emitted per dm3 of fuel (gasoline)

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
    
    def min_generation(model,t):
        return model.P_from_CHP[t] >= model.P_CHP_min

    def max_generation(model,t):
        return model.P_from_CHP[t] <= model.P_CHP_max 

    def generation(model,t):
        return model.P_from_CHP[t] == model.Q_from_CHP[t] * model.CHP_P_to_Q_ratio

    def fuel_consumption(model,t):
        return model.CHP_fuel_cons[t] == model.P_from_CHP[t] * model.time_step * model.CHP_fuel_cons_ratio 

    def operation_costs(model,t):
        return model.CHP_op_cost[t] == model.CHP_fuel_cons[t] * model.CHP_fuel_price
    
    def emissions(model,t):
        return model.CHP_emissions[t] == model.CHP_fuel_cons[t] * model.CHP_spec_em
    
class objective:
    def __init__(self):
        self.list_var = []
        self.list_text_var = []
        self.list_param = []
        self.list_text_param = []
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input

        #defining energy type to build connections with other componets correctly
        self.super_class = 'objective'