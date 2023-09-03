import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import math
import sys

class charging_station:

    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        #default values in case of no input
        self.list_var = ['charging_station_op_cost','charging_station_inv_cost','charging_station_emissions'] #no powers
        self.list_text_var = ['within = pyo.NegativeReals','within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']

        # self.list_param = ['charging_station_inv_specific_costs','charging_station_selling_price','charging_station_spec_emissions']
        # self.list_text_param = ['','','']
        
        # self.list_series = ['P_to_charging_station']
        # self.list_text_series =['model.HOURS']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        
        self.super_class = 'demand'

        self.path_input = './input/'
        self.path_output = './output/'
        self.name_file = 'df_input.xlsx'

        self.param_charging_station_inv_specific_costs = 200000
        self.param_charging_station_selling_price = 0.60
        self.param_charging_station_spec_emissions = 0.05

        #defining paramenters for functions that are not constraints, includiing default values:
        self.dict_parameters  = {'list_sheets':['other','other','other','other'],
                                 'charging_station_mult': 1.2,
                                 'reference_date':datetime(2023,10,1),
                                 'number_hours': time_span + 1,
                                 'number_days': math.ceil((time_span + 1)/24)}
        
        print(self.dict_parameters['number_days'])
        
        self.dict_series = {'list_sheets':['mix and capacity','mix and capacity','mix and capacity','mix SoC','mix SoC','mix SoC',
                                           'data_charging_station','data_charging_station','data_charging_station'],
                       
                       'list_models':['Tesla MODEL 3','VW E-UP','VW ID.3','Renault ZOE','Smart FORTWO',
                                      'Hyundai KONA 39 kWh','Hyundai KONA 64 kWh','Hyundai IONIQ5 58kWh',
                                      'Hyundai IONIQ5 77kWh','VW ID.4','Fiat 500 E','BMW I3 22kWh','BMW I3 33kWh',
                                      'BMW I3 42kWh','Opel CORSA','MINI Cooper SE','Audi E-TRON','Peugeot 208',
                                      'Renault TWINGO','Opel MOKKA','Nissan LEAF 24 kWh','Nissan LEAF 30 kWh',
                                      'Nissan LEAF 40 kWh','Nissan LEAF 62 kWh','Audi Q4','Dacia SPRING','TESLA MODEL Y',
                                      'VW ID.5','SKODA ENYAQ 77kWh','SKODA ENYAQ 58kWh','CUPRA BORN 77kWh',
                                      'CUPRA BORN 58kWh','RENAULT MEGANE E-TECH','Polestar 2','KIA EV6','Porsche Taycan','Others'],

                        'list_mix_models':[0.065,0.033,0.036,0.036,0.022,0.023,0.023,0.015,0.015,0.021,0.067,
                                           0.009,0.009,0.009,0.019,0.03,0.033,0.013,0.016,0.018,0.002,0.002,
                                           0.002,0.002,0.025,0.021,0.045,0.021,0.018,0.018,0.007,0.007,0.005,
                                           0.012,0.01,0.009,0.282],

                        'list_capacity':[50,36.8,58,52,17.6,39,64,58,77,77,24,22,33,42,45,28.9,85,45,21.3,45,
                                         24,30,40,62,76.6,26.8,75,77,77,58,77,58,60,75,74,83.7,52.631],

                        'list_SoC':[0,0.05,0.1,0.15,0.2,0.25,0.3,0.35,0.4,0.45,0.5,0.55,0.6,0.65,0.7,0.75,0.8,0.85,0.9,0.95,1],

                        'list_mix_start_SoC':[0,0.07,0.1,0.13,0.15,0.17,0.14,0.12,0.08,0.03,0.01,0,0,0,0,0,0,0,0,0,0],

                        'list_mix_end_SoC':[0,0,0,0,0,0,0,0,0,0,0,0,0.03,0.07,0.1,0.11,0.13,0.17,0.15,0.13,0.11],

                        f"{self.name_of_instance}_list_means": [8, 12, 16],
                        f"{self.name_of_instance}_list_std": [0.5, 0.5, 0.5],
                        f"{self.name_of_instance}_list_size": [5, 5, 5] }
        
        # self.list_altered = []
        # self.list_text_altered = []
        print(self.dict_parameters['number_days'])
        # calling functions to try and read parameter values as soon as class is created
        self.read_parameters(self.dict_parameters)
        self.read_series(self.dict_series)

        print(self.dict_parameters['number_days'])
        # creating dataframes that are going to be used in non-constraint function
        self.df_mix_capacity = pd.DataFrame({'Model': self.dict_series['list_models'],
                                             'Mix': self.dict_series['list_mix_models'],
                                             'Max Capacity': self.dict_series['list_capacity']})
        
        self.df_SoC = pd.DataFrame({'SoC':self.dict_series['list_SoC'],
                               'mix initial':self.dict_series['list_mix_start_SoC'],
                               'mix final':self.dict_series['list_mix_end_SoC']})
        
        self.df_data_distribution = pd.DataFrame({'hours of peak': self.dict_series[f"{self.name_of_instance}_list_means"],
                                                  'std of each peak': self.dict_series[f"{self.name_of_instance}_list_std"],
                                                  'number of cars in each peak': self.dict_series[f"{self.name_of_instance}_list_size"]})
        
        self.param_P_to_charging_station = self.charging_demand_calculation()

        print('--------------' + str(len(self.param_P_to_charging_station)))

    def read_parameters(self, parameters):
        count = 0
        list_sheets = parameters['list_sheets']
        del parameters['list_sheets']
        for name, default_value in parameters.items():
            df_others = pd.read_excel(self.path_input + self.name_file, index_col=0, sheet_name = list_sheets[count])
            df_others.index.name = None
            try:
                parameters[name] = df_others.loc[name, 'Value']
            except KeyError:
                pass
            count += 1

    def read_series(self, series):
        count = 0 
        list_sheets = series['list_sheets']
        del series['list_sheets']
        for name, default_value in series.items():
            df_series = pd.read_excel(self.path_input + self.name_file, sheet_name = list_sheets[count])
            try:
                series[name] = df_series[name].tolist()
            except KeyError:
                pass
            count += 1        

    #function for transforming timestamp to YYYY-MM_DD hh:mm
    def convert_time_stamp(self, day_integer, float_number):
        seconds_integer = int(float_number * 3600)
        delta = timedelta(seconds = seconds_integer,days = day_integer)
        converted_time_stamp = self.dict_parameters['reference_date'] + delta
        converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
        return converted_time_stamp

    #function to return elements of normal distribution
    def generate_normal_distribution(self,mean, standard_deviation, size):
        # Generate random numbers from a normal distribution
        samples = np.random.normal(mean, standard_deviation, size)
        return samples

    def charging_demand_calculation(self):
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

        print(self.dict_parameters['number_days'])
        for i in range(0,self.dict_parameters['number_days']):
            for j in self.df_data_distribution.index:
                mean = self.df_data_distribution.loc[j,'hours of peak']
                standard_deviation = self.df_data_distribution.loc[j,'std of each peak']
                size = self.df_data_distribution.loc[j,'number of cars in each peak']

                timestamps = self.generate_normal_distribution(mean, standard_deviation, size)
                for k in timestamps:
                    converted_time_stamp = self.convert_time_stamp(i, k)
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

        # path_output = 'G:/My Drive/02. Mestrado/TU Darmstadt/Energy Science and Engineering/5. Sommersemester 2023/Masterarbeit/5. Thesis/3. Python Routine/masterthesis/output/'
        df_schedule.to_excel(self.path_output + 'df_schedule.xlsx',index = False)

        list_time = []

        for i in range(0,self.dict_parameters['number_hours']):
            delta = timedelta(hours = i + 1)
            converted_time_stamp = self.dict_parameters['reference_date'] + delta
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

        # df_structured.to_excel(path_output + 'df_structured.xlsx',index = False)
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
        df_structured.to_excel(self.path_output + 'df_structured.xlsx', index = False)

        lista = df_structured['total power'].tolist()

        return [self.dict_parameters['charging_station_mult'] * i for i in lista]
        
    def constraint_operation_costs(model,t):
        return model.charging_station_op_cost[t] == - model.param_P_to_charging_station[t] * model.time_step * model.param_charging_station_selling_price
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.charging_station_inv_cost[t] == model.param_charging_station_inv_specific_costs
        else:
            return model.charging_station_inv_cost[t] == 0
        
    def constraint_emissions(model,t):
        return model.charging_station_emissions[t] == model.param_P_to_charging_station[t] * model.param_charging_station_spec_emissions
    

myClass = charging_station('charging_station',100)