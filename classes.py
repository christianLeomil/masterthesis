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
        self.list_var = [] #no powers
        self.list_text_var = []
        self.list_param = ['pv_eff','pv_area']
        self.list_text_param = ['','']
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input
        self.pv_eff = 0.7
        self.pv_area = 1

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}

    def solar_rule(model,t):
        return model.P_from_pv[t] == model.P_solar[t] * model.pv_eff * model.pv_area
    
class bat:
    def __init__(self):
        self.list_var = ['bat_SOC','bat_K_ch','bat_K_dis'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',]
        self.list_param = ['bat_starting_SOC','bat_ch_eff','bat_dis_eff','bat_E_max',
                           'bat_c_rate_ch','bat_c_rate_dis']
        self.list_text_param = ['','','','','','']
        self.list_series = []
        self.list_text_series = []

        #default values in case of no input
        self.bat_starting_SOC = 0.5
        self.bat_ch_eff = 0.95
        self.bat_dis_eff = 0.95
        self.bat_E_max = 100
        self.bat_c_rate_ch = 1
        self.bat_c_rate_dis = 1

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}

    def function_rule(model,t):
            if t ==1:
                return model.bat_SOC[t] == model.bat_starting_SOC
            else:
                return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_to_bat[t-1] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t-1]/model.bat_dis_eff) * model.time_step / model.bat_E_max

    def charge_limit(model,t):
        return model.P_to_bat[t] <= model.bat_E_max * model.bat_K_ch[t] * model.bat_c_rate_ch
    
    def discharge_limit(model,t):
        return model.P_from_bat[t] <= model.bat_E_max * model.bat_K_dis[t] * model.bat_c_rate_dis
    
    def keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1
    
class demand:
    def __init__(self):
        self.list_var = [] #no powers
        self.list_text_var = []
        self.list_param = []
        self.list_text_param = []
        self.list_series = ['P_to_demand','Q_to_demand']
        self.list_text_series =['model.HOURS','model.HOURS']

        #default values in case of no input
        self.P_to_demand = 500
        self.Q_to_demand = 1000

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}

class net:
    def __init__(self):
        self.list_var = ['net_sell_electric','net_buy_electric','net_sell_thermal','net_buy_thermal'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']
        self.list_param = []
        self.list_text_param = []
        self.list_series = ['net_cost_buy_electric','net_cost_sell_electric',
                            'net_cost_buy_thermal','net_cost_sell_thermal']
        self.list_text_series =['model.HOURS','model.HOURS',
                                'model.HOURS','model.HOURS']

        #default values in case of no input
        self.net_cost_buy_electric = 0.4
        self.net_cost_sell_electric = 0.3
        self.net_cost_buy_thermal = 0.2
        self.net_cost_sell_thermal = 0.1

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
    def sell_energy_electric(model,t):
        return model.net_sell_electric[t] == model.P_to_net[t] * model.time_step * model.net_cost_sell_electric[t]
    
    def buy_energy_electric(model,t):
        return model.net_buy_electric[t] == model.P_from_net[t] * model.time_step * model.net_cost_buy_electric[t]
    
    def sell_energy_thermal(model,t):
        return model.net_sell_thermal[t] == model.Q_to_net[t] * model.time_step * model.net_cost_sell_thermal[t]
    
    def buy_energy_thermal(model,t):
        return model.net_buy_thermal[t] == model.Q_from_net[t] * model.time_step * model.net_cost_buy_thermal[t]
    
class CHP:
    def __init__(self):
        self.list_var = ['CHP_fuel_cons','CHP_op_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']
        self.list_param = ['P_CHP_max','P_CHP_min','CHP_P_to_Q_ratio','CHP_fuel_cons_ratio','CHP_fuel_price']
        self.list_text_param = ['','','','','']
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input
        self.P_CHP_max = 20000 #W electric
        self.P_CHP_min = 0.2 * self.P_CHP_max #minimal load of 20% 
        self.CHP_P_to_Q_ratio = 0.5 
        self.CHP_fuel_cons_ratio = 0.5 #dm3 per kWh of P_from_CHP
        self.CHP_fuel_price = 5 # EUROS/dm3 of fuel 

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
    
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
    
class objective:
    def __init__(self):
        self.list_var = []
        self.list_text_var = []
        self.list_param = []
        self.list_text_param = []
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input
