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

class pv:
    def __init__(self):
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
        self.pv_spec_em = 0 #There is no CO2 emission from generation energy with PV
        self.pv_kWp_per_area  = 0.12 # kWP per m2 of Pv
        self.pv_inv_per_kWp = 1000 # EURO per kWp

        #default series
        self.E_pv_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        self.super_class = 'generator'

    def generation_rule(model,t):
        return model.P_from_pv[t] == (model.E_pv_solar[t] / model.time_step) * model.pv_eff * model.pv_area 
    
    def operation_costs(model,t):
        return model.pv_op_cost[t] == model.pv_area * model.pv_spec_op_cost
    
    def emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.pv_spec_em
    
    def investment_costs(model,t):
        if t == 1:
            return model.pv_inv_cost[t] == model.pv_area * model.pv_kWp_per_area * model.pv_inv_per_kWp
        else:
            return model.pv_inv_cost[t] == 0
    
class solar_th:
    def __init__(self):
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
        self.solar_th_spec_em = 0 #There is no CO2 emission from generation energy with PV
        self.solar_th_inv_per_area = 700 #EURO per m2 aperture

        #default series
        self.E_solar_th_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'no',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
        
    def generation_rule(model,t):
        return model.Q_from_solar_th[t] == (model.E_solar_th_solar[t] * model.time_step) * model.solar_th_area * model.solar_th_eff
        
    def operation_cost(model,t):
        return model.solar_th_op_cost[t] == model.solar_th_area * model.solar_th_spec_op_cost
        
    def emission(model,t):
        return model.solar_th_emissions[t] == model.Q_from_solar_th[t] * model.solar_th_spec_em
    
    def investment_costs(model,t):
        if t == 1:
            return model.solar_th_inv_cost[t] == model.solar_th_area * model.solar_th_inv_per_area
        else:
            return model.solar_th_inv_cost[t] == 0


class pvt:
    def __init__(self):
        self.list_var = ['pvt_op_cost','pvt_emissions','pvt_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_param = ['pvt_eff','pvt_area','pvt_spec_op_cost','pvt_spec_em','pvt_Q_to_P_ratio',
                           'pvt_inv_per_area']
        self.list_text_param = ['','','','','','']

        self.list_series = ['E_pvt_solar']
        self.list_text_series =['model.HOURS']

        #default values in case of no input
        self.pvt_eff = 0.20 # aproximate overall efficiency of pv cells 
        self.pvt_area = 50 # m^2
        self.pvt_spec_op_cost = 0.02 # cost per hour per m^2 area of pv installed
        self.pvt_spec_em = 0 #There is no CO2 emission from generation energy with PV
        self.pvt_Q_to_P_ratio = 1.3 #proportion between power and generated heat
        self.pvt_inv_per_area = 850 # EURO per m2 aperture

        #default series
        self.E_pvt_solar = 0.12 # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
        
    def generation_rule(model,t):
        return model.P_from_pvt[t] == (model.E_pvt_solar[t] * model.time_step) * model.pvt_area * model.pvt_eff
    
    def thermal_energy_rule(model,t):
        return model.Q_from_pvt[t] == model.P_from_pvt[t] * model.pvt_Q_to_P_ratio
        
    def operation_cost(model,t):
        return model.pvt_op_cost[t] == model.pvt_area * model.pvt_spec_op_cost
        
    def emission(model,t):
        return model.pvt_emissions[t] == model.Q_from_pvt[t] * model.pvt_spec_em

    def investment_costs(model,t):
        if t == 1:
            return model.pvt_inv_cost[t] == model.pvt_area * model.pvt_inv_per_area
        else:
            return model.pvt_inv_cost[t] == 0

class bat:
    def __init__(self):
        self.list_var = ['bat_SOC','bat_K_ch','bat_K_dis','bat_op_cost','bat_emissions','bat_SOC_max',
                         'bat_integer','bat_cumulated_aging','bat_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',
                              'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'within = pyo.Integers', 'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']
        
        self.list_param = ['bat_starting_SOC','bat_ch_eff','bat_dis_eff','bat_E_max_initial',
                           'bat_c_rate_ch','bat_c_rate_dis','bat_spec_op_cost',
                           'bat_spec_em','bat_DoD','bat_final_SoH','bat_cycles','bat_aging',
                           'bat_inv_per_capacity']
        self.list_text_param = ['','','','','','','','','','','','','']

        self.list_series = []
        self.list_text_series = []

        #default values in case of no input
        # self.bat_E_max = 100 ############
        self.bat_E_max_initial = 100
        self.bat_starting_SOC = 0.5
        self.bat_ch_eff = 0.95
        self.bat_dis_eff = 0.95
        self.bat_c_rate_ch = 1
        self.bat_c_rate_dis = 1
        self.bat_spec_op_cost = 0.01
        self.bat_spec_em = 0 # no emission for operating batteries
        self.bat_DoD = 0.7
        self.bat_final_SoH = 0.7
        self.bat_cycles = 9000 # full cycles before final SoH is reached and battery is replaced
        self.bat_aging = (self.bat_E_max_initial * (1 -self.bat_final_SoH)) / (self.bat_cycles * 2 * self.bat_E_max_initial) / self.bat_E_max_initial
        self.bat_inv_per_capacity = 650 # EURO per kWh capacity

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        
        self.super_class = 'transformer'

    def depth_of_discharge(model,t):
        return model.bat_SOC[t] >= (1 - model.bat_DoD)
    
    def max_state_of_charge(model,t):
        if t == 1:
            return model.bat_SOC[t] <= 1
        else:
            return model.bat_SOC[t] <= model.bat_SOC_max[t]
    
    def cumulated_aging(model,t):
        if t == 1:
            return model.bat_cumulated_aging[t] == (model.P_from_bat[t] + 
                                                    model.P_to_bat[t]) * model.time_step * model.bat_aging
        else:
            return model.bat_cumulated_aging[t] == model.bat_cumulated_aging[t-1] + (model.P_from_bat[t] + 
                                                                                     model.P_to_bat[t]) * model.time_step * model.bat_aging
        
    def upper_integer_rule(model,t):
        if t == 1:
            return model.bat_integer[t] == 0
        else:
            return model.bat_integer[t] <= model.bat_cumulated_aging[t] / (1-model.bat_final_SoH)
    
    def lower_integer_rule(model,t):
        if t == 1:
            return model.bat_integer[t] == 0 
        else:
            return model.bat_integer[t] >= model.bat_cumulated_aging[t] / (1-model.bat_final_SoH) - 1
     
    def aging(model,t):
        if t == 1:
            return model.bat_SOC_max[t] == 1
        else:
            return model.bat_SOC_max[t] == 1 - model.bat_cumulated_aging[t] + model.bat_integer[t] * (1 - model.bat_final_SoH)
    
    def function_rule(model,t):
            if t == 1:
                return model.bat_SOC[t] == model.bat_starting_SOC + (model.P_to_bat[t] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.bat_dis_eff) * model.time_step / model.bat_E_max_initial
            else:
                return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_to_bat[t] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.bat_dis_eff) * model.time_step / model.bat_E_max_initial

    def charge_limit(model,t):
        return model.P_to_bat[t] <= model.bat_E_max_initial * model.bat_K_ch[t] * model.bat_c_rate_ch

    def discharge_limit(model,t):
        return model.P_from_bat[t] <= model.bat_E_max_initial *  model.bat_K_dis[t] * model.bat_c_rate_dis
    
    def keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1

    def operation_costs(model,t):
        return model.bat_op_cost[t] == model.bat_E_max_initial * model.bat_spec_op_cost
    
    def emissions(model,t):
        return model.bat_emissions[t] == (model.P_from_bat[t] + model.P_to_bat[t]) * model.bat_spec_em
    
    def investment_costs(model,t):
        if t == 1:
            return model.bat_inv_cost[t] == model.bat_E_max_initial * model.bat_inv_per_capacity
        else:
            return model.bat_inv_cost[t] == model.bat_E_max_initial * model.bat_inv_per_capacity * (model.bat_integer[t] - model.bat_integer[t-1])

class demand:
    def __init__(self):
        self.list_var = ['demand_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals']

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

    def investment_costs(model,t):
        return model.demand_inv_cost[t] == 0

class net:
    def __init__(self):
        self.list_var = ['net_sell_electric','net_buy_electric','net_sell_thermal','net_buy_thermal'
                         ,'net_emissions','net_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals'
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
    
    def investment_costs(model,t):
        return model.net_inv_cost[t] == 0

class CHP:
    def __init__(self):
        self.list_var = ['CHP_fuel_cons','CHP_op_cost','CHP_emissions','CHP_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals']

        self.list_param = ['P_CHP_max','P_CHP_min','CHP_P_to_Q_ratio','CHP_fuel_cons_ratio','CHP_fuel_price',
                           'CHP_spec_em','CHP_inv_cost_per_power']
        self.list_text_param = ['','','','','','','']
        
        self.list_series = []
        self.list_text_series =[]

        #default values in case of no input
        self.P_CHP_max = 20 #W electric
        self.P_CHP_min = 0
        self.CHP_P_to_Q_ratio = 0.5 
        self.CHP_fuel_cons_ratio = 0.105 #dm3 per kWh of P_from_CHP
        self.CHP_fuel_price = 5 # EUROS/dm3 of fuel 
        self.CHP_spec_em = 2.3 # kg of CO2 emitted per dm3 of fuel (gasoline)
        self.CHP_inv_cost_per_power = 1700 # EURO per kW power

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
    
    def investment_costs(model,t):
        if t == 1:
            return model.CHP_inv_cost[t] == model.P_CHP_max * model.CHP_inv_cost_per_power
        else:
            return model.CHP_inv_cost[t] == 0
    
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

