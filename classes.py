import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import random
import math

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


class charging_station:

    def __init__(self):
        #default values in case of no input
        self.list_var = [] #no powers
        self.list_text_var = []

        self.list_param = []
        self.list_text_param = []
        
        self.list_series = ['P_to_charging_station']
        self.list_text_series =['model.HOURS']

        #defining paramenters:
        self.charging_station_mult = 1.2
        # self.charging_station_demand = self.charging_demand_rule(self.charging_station_mult)
    
        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        
        self.super_class = 'demand'

        #default values in case of no input
        self.reference_date = datetime(2023,10,1)

        self.list_models = ['Tesla MODEL 3','VW E-UP','VW ID.3','Renault ZOE','Smart FORTWO',
                            'Hyundai KONA 39 kWh','Hyundai KONA 64 kWh','Hyundai IONIQ5 58kWh',
                            'Hyundai IONIQ5 77kWh','VW ID.4','Fiat 500 E','BMW I3 22kWh','BMW I3 33kWh',
                            'BMW I3 42kWh','Opel CORSA','MINI Cooper SE','Audi E-TRON','Peugeot 208',
                            'Renault TWINGO','Opel MOKKA','Nissan LEAF 24 kWh','Nissan LEAF 30 kWh',
                            'Nissan LEAF 40 kWh','Nissan LEAF 62 kWh','Audi Q4','Dacia SPRING','TESLA MODEL Y',
                            'VW ID.5','SKODA ENYAQ 77kWh','SKODA ENYAQ 58kWh','CUPRA BORN 77kWh',
                            'CUPRA BORN 58kWh','RENAULT MEGANE E-TECH','Polestar 2','KIA EV6','Porsche Taycan',
                            'Others']

        self.list_mix_models = [0.065,0.033,0.036,0.036,0.022,0.023,0.023,0.015,0.015,0.021,0.067,0.009,0.009,
                                0.009,0.019,0.03,0.033,0.013,0.016,0.018,0.002,0.002,0.002,0.002,0.025,0.021,0.045,
                                0.021,0.018,0.018,0.007,0.007,0.005,0.012,0.01,0.009,0.282]

        self.list_capacity = [50,36.8,58,52,17.6,39,64,58,77,77,24,22,33,42,45,28.9,85,45,21.3,45,24,30,40,62,76.6,
                              26.8,75,77,77,58,77,58,60,75,74,83.7,52.631]
        
        self.df_mix_capacity = pd.DataFrame({'Model': self.list_models,
                                             'Mix': self.list_mix_models,
                                             'Max Capacity': self.list_capacity})


        self.list_SoC = [0,0.05,0.1,0.15,0.2,0.25,0.3,0.35,0.4,0.45,0.5,0.55,0.6,0.65,0.7,0.75,0.8,0.85,0.9,0.95,1]

        self.list_mix_start_SoC = [0,0.07,0.1,0.13,0.15,0.17,0.14,0.12,0.08,0.03,0.01,0,0,0,0,0,0,0,0,0,0]

        self.list_mix_end_SoC = [0,0,0,0,0,0,0,0,0,0,0,0,0.03,0.07,0.1,0.11,0.13,0.17,0.15,0.13,0.11]
        
        self.df_SoC = pd.DataFrame({'SoC':self.list_SoC,
                               'mix initial':self.list_mix_start_SoC,
                               'mix final':self.list_mix_end_SoC})
        
        self.list_means = [8, 12, 16]
        self.list_std = [0.5, 0.5, 0.5]
        self.list_size = [5, 5, 5]
        self.number_hours = 101
        self.number_days = math.ceil(self.number_hours/24)
        
        self.df_data_distribution = pd.DataFrame({'hours of peak': self.list_means,
                                        'std of each peak': self.list_std,
                                        'number of cars in each peak': self.list_size})
    
        self.P_to_charging_station = self.X_charging_demand_calculation()

    #function for transforming timestamp to YYYY-MM_DD hh:mm
    def X_convert_time_stamp(self, day_integer, float_number):
        seconds_integer = int(float_number * 3600)
        delta = timedelta(seconds = seconds_integer,days = day_integer)
        converted_time_stamp = self.reference_date + delta
        converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
        return converted_time_stamp

    #function to return elements of normal distribution
    def X_generate_normal_distribution(self,mean, standard_deviation, size):
        # Generate random numbers from a normal distribution
        samples = np.random.normal(mean, standard_deviation, size)
        return samples

    def X_charging_demand_calculation(self):

        #creating list of car models with mix of EVs in germany
        list_models = []
        for i in self.df_mix_capacity.index:
            times = int(self.df_mix_capacity.loc[i,'Mix']*100)
            for j in range(0,times):
                list_models.append(self.df_mix_capacity.loc[i,'Model'])

        #creating list of start and end SoC with input mix
        list_start_SoC = []
        list_end_SoC = []
        for i in self.df_SoC.index:
            times_start_SoC = int(self.df_SoC.loc[i,'mix initial']*100)
            times_final_SoC = int(self.df_SoC.loc[i,'mix final']*100)
            for j in range(0,times_start_SoC):
                list_start_SoC.append(self.df_SoC.loc[i,'SoC'])
            for j in range(0,times_final_SoC):
                list_end_SoC.append(self.df_SoC.loc[i,'SoC'])

        list_days_models = []
        list_hours = []
        list_initial_SoC = []
        list_final_SoC = []

        for i in range(0,self.number_days):
            for j in self.df_data_distribution.index:

                mean = self.df_data_distribution.loc[j,'hours of peak']
                standard_deviation = self.df_data_distribution.loc[j,'std of each peak']
                size = self.df_data_distribution.loc[j,'number of cars in each peak']

                timestamps = self.X_generate_normal_distribution(mean, standard_deviation, size)
                for k in timestamps:
                    converted_time_stamp = self.X_convert_time_stamp(i, k)
                    list_hours.append(converted_time_stamp)
                    list_days_models.append(random.choice(list_models))

                    initial_SoC = random.choice(list_start_SoC)
                    final_SoC = random.choice(list_end_SoC)
                    while initial_SoC >= final_SoC:
                        initial_SoC = random.choice(list_start_SoC)
                        final_SoC = random.choice(list_end_SoC)

                    list_initial_SoC.append(initial_SoC)
                    list_final_SoC.append(final_SoC)

        df_schedule = pd.DataFrame({'car models': list_days_models,
                                    'time stamp':list_hours,
                                    'initial SoC':list_initial_SoC,
                                    'final SoC':list_final_SoC})

        path_output = 'G:/My Drive/02. Mestrado/TU Darmstadt/Energy Science and Engineering/5. Sommersemester 2023/Masterarbeit/5. Thesis/3. Python Routine/masterthesis/output/'
        df_schedule.to_excel(path_output + 'df_schedule.xlsx',index = False)


        list_time = []

        for i in range(0,self.number_hours):
            delta = timedelta(hours = i + 1)
            converted_time_stamp = self.reference_date + delta
            converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
            list_time.append(converted_time_stamp)

        list_models = []
        list_start_SoC = []
        list_end_SoC = []

        for i in range(1,len(list_time)):
            df_temp = df_schedule[df_schedule['time stamp'].between(list_time[i-1],list_time[i])]
            list_models.append(df_temp['car models'].to_list())
            list_start_SoC.append(df_temp['initial SoC'].to_list())
            list_end_SoC.append(df_temp['final SoC'].to_list())

        list_time = list_time[:-1]
        df_structured = pd.DataFrame({'timestamp':list_time,
                                    'models':list_models,
                                    'start SoC':list_start_SoC,
                                    'end SoC':list_end_SoC})

        df_structured.to_excel(path_output + 'df_structured.xlsx',index = False)
        list_power = []
        list_total_power = []
        for i in df_structured.index:
            models  = df_structured['models'].iloc[i]
            start_SoCs = df_structured['start SoC'].iloc[i]
            end_SoCs = df_structured['end SoC'].iloc[i]
            capacity = 0
            list_power_partial = []
            for j in range(0,len(models)):
                model = models[j]
                start_SoC = start_SoCs[j]
                end_SoC = end_SoCs[j]
                capacity = self.df_mix_capacity.loc[self.df_mix_capacity['Model'] == model,'Max Capacity'].values[0]
                power = (end_SoC-start_SoC) * capacity
                list_power_partial.append(power)
            list_power.append(list_power_partial)
            list_total_power.append(sum(list_power_partial))

        df_structured['power'] = list_power
        df_structured['total power'] = list_total_power
        # df_structured.to_excel(path_output + 'df_structured.xlsx', index = False)

        lista = df_structured['total power'].tolist()
        print(lista)
        print(len(lista))

        return [self.charging_station_mult * i for i in lista]
    
# myClass = charging_station()

