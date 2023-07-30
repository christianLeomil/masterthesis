# import numpy as np
# import pandas as pd
# import random
# import matplotlib.pyplot as plt
# from datetime import datetime, timedelta
# import math

# path_input = './input/'
# path_output = './output/'
# name_file = 'inputs5.xlsx'

# df_mix_capacity = pd.read_excel(path_input + name_file, sheet_name = 'Mix and Capacity')
# df_SoC = pd.read_excel(path_input + name_file, sheet_name = 'Mix SoC')

# reference_date = datetime(2023,1,10)

# #function to return elements of normal distribution
# def generate_normal_distribution(mean, standard_deviation, size):
#     # Generate random numbers from a normal distribution
#     samples = np.random.normal(mean, standard_deviation, size)
#     return samples

# #function for transforming timestamp to YYYY-MM_DD hh:mm
# def convert_time_stamp(day_integer, float_number):
#     seconds_integer = int(float_number * 3600)
#     delta = timedelta(seconds = seconds_integer,days = day_integer)
#     converted_time_stamp = reference_date + delta
#     converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')

#     return converted_time_stamp

# #creating list of car models with mix of EVs in germany
# list_models = []
# for i in df_mix_capacity.index:
#     times = int(df_mix_capacity.loc[i,'Mix']*100)
#     for j in range(0,times):
#         list_models.append(df_mix_capacity.loc[i,'Model'])

# #creating list of start and end SoC with input mix
# list_start_SoC = []
# list_end_SoC = []
# for i in df_SoC.index:
#     times_start_SoC = int(df_SoC.loc[i,'mix initial']*100)
#     times_final_SoC = int(df_SoC.loc[i,'mix final']*100)
#     for j in range(0,times_start_SoC):
#         list_start_SoC.append(df_SoC.loc[i,'SoC'])
#     for j in range(0,times_final_SoC):
#         list_end_SoC.append(df_SoC.loc[i,'SoC'])

# list_means = [8, 12, 18]
# list_std = [0.1, 0.1, 0.1]
# list_size = [5, 5, 5]

# df_data_distribution = pd.DataFrame({'hours of peak': list_means,
#                                      'std of each peak':list_std,
#                                      'number of cars in each peak':list_size})

# list_days_models = []
# list_hours = []
# list_initial_SoC = []
# list_final_SoC = []

# number_hours = 100
# number_days = math.ceil(number_hours/24)
# print('-----------number of days------')
# print(number_days)

# for i in range(0,number_days):
#     for j in df_data_distribution.index:

#         mean = df_data_distribution.loc[j,'hours of peak']
#         standard_deviation = df_data_distribution.loc[j,'std of each peak']
#         size = df_data_distribution.loc[j,'number of cars in each peak']

#         timestamps = generate_normal_distribution(mean, standard_deviation, size)
#         for k in timestamps:
#             converted_time_stamp = convert_time_stamp(i, k)
#             list_hours.append(converted_time_stamp)
#             list_days_models.append(random.choice(list_models))

#             initial_SoC = random.choice(list_start_SoC)
#             final_SoC = random.choice(list_end_SoC)
#             while initial_SoC >= final_SoC:
#                 initial_SoC = random.choice(list_start_SoC)
#                 final_SoC = random.choice(list_end_SoC)

#             list_initial_SoC.append(initial_SoC)
#             list_final_SoC.append(final_SoC)

# df_schedule = pd.DataFrame({'car models': list_days_models,
#                             'time stamp':list_hours,
#                             'initial SoC':list_initial_SoC,
#                             'final SoC':list_final_SoC})

# df_schedule.to_excel(path_output + 'df_schedule.xlsx',index = False)

# list_time = []

# for i in range(0,number_hours):
#     delta = timedelta(hours = i + 1)
#     converted_time_stamp = reference_date + delta
#     converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
#     list_time.append(converted_time_stamp)

# list_models = []
# list_start_SoC = []
# list_end_SoC = []

# for i in range(1,len(list_time)):
#     df_temp = df_schedule[df_schedule['time stamp'].between(list_time[i-1],list_time[i])]
#     list_models.append(df_temp['car models'].to_list())
#     list_start_SoC.append(df_temp['initial SoC'].to_list())
#     list_end_SoC.append(df_temp['final SoC'].to_list())

# list_time = list_time[:-1]
# df_structured = pd.DataFrame({'timestamp':list_time,
#                               'models':list_models,
#                               'start SoC':list_start_SoC,
#                               'end SoC':list_end_SoC})

# # df_structured.to_excel(path_output + 'df_structured.xlsx',index = False)
# list_power = []
# list_total_power = []
# for i in df_structured.index:
#     models  = df_structured['models'].iloc[i]
#     start_SoCs = df_structured['start SoC'].iloc[i]
#     end_SoCs = df_structured['end SoC'].iloc[i]
#     capacity = 0
#     list_power_partial = []
#     for j in range(0,len(models)):
#         model = models[j]
#         start_SoC = start_SoCs[j]
#         end_SoC = end_SoCs[j]
#         capacity = df_mix_capacity.loc[df_mix_capacity['Model'] == model,'Max Capacity'].values[0]
#         power = (end_SoC-start_SoC) * capacity
#         list_power_partial.append(power)
#     list_power.append(list_power_partial)
#     list_total_power.append(sum(list_power_partial))

# df_structured['power'] = list_power
# df_structured['total power'] = list_total_power
# df_structured.to_excel(path_output + 'df_structured.xlsx', index = False)




import pandas as pd

list_models = ['Tesla MODEL 3','VW E-UP','VW ID.3','Renault ZOE','Smart FORTWO',
                    'Hyundai KONA 39 kWh','Hyundai KONA 64 kWh','Hyundai IONIQ5 58kWh',
                    'Hyundai IONIQ5 77kWh','VW ID.4','Fiat 500 E','BMW I3 22kWh','BMW I3 33kWh',
                    'BMW I3 42kWh','Opel CORSA','MINI Cooper SE','Audi E-TRON','Peugeot 208',
                    'Renault TWINGO','Opel MOKKA','Nissan LEAF 24 kWh','Nissan LEAF 30 kWh',
                    'Nissan LEAF 40 kWh','Nissan LEAF 62 kWh','Audi Q4','Dacia SPRING','TESLA MODEL Y',
                    'VW ID.5','SKODA ENYAQ 77kWh','SKODA ENYAQ 58kWh','CUPRA BORN 77kWh',
                    'CUPRA BORN 58kWh','RENAULT MEGANE E-TECH','Polestar 2','KIA EV6','Porsche Taycan',
                    'Others']

list_mix_models = ['0.065','0.033','0.036','0.036','0.022','0.023','0.023','0.015','0.015','0.021',
                        '0.067','0.009','0.009','0.009','0.019','0.03','0.033','0.013','0.016','0.018',
                        '0.002','0.002','0.002','0.002','0.025','0.021','0.045','0.021','0.018','0.018',
                        '0.007','0.007','0.005','0.012','0.01','0.009','0.282']

list_capacity = ['50','36.8','58','52','17.6','39','64','58','77','77','24','22','33','42','45','28.9',
                        '85','45','21.3','45','24','30','40','62','76.6','26.8','75','77','77','58','77','58',
                        '60','75','74','83.7','52.6305555555556']

df_mix_capacity = pd.DataFrame({'Model': list_models,
                                        'Mix': list_mix_models,
                                        'Max Capacity': list_capacity})
print(df_mix_capacity)