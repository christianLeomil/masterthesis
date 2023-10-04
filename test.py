import pandas as pd
import classes
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import os

control = {'path_charts' : './output/charts/',
           'path_output' : './output/'}

list_elements = ['pv1', 'bat1', 'demand1', 'net1']
list_type = ['pv', 'bat', 'demand', 'net']

df_aux = pd.DataFrame({'element':list_elements,
                       'type': list_type})


def charts_generator2(control,df_aux):
    path_charts = control['path_charts']
    path_output = control['path_output']
    # path_output = control.path_output
    reference_date = datetime(2023,10,1)
    # reference_date = control.reference_date

    df_final = pd.read_excel(path_output + 'df_final.xlsx')
    df_final['Date'] = reference_date + pd.to_timedelta(df_final['TimeStep'], unit = 'h')
    # df_final = df_final.head(100)
    df_final = df_final[:100]

    x_axis = df_final['TimeStep'].tolist()

    for i in df_aux.index:
        element = df_aux['element'].iloc[i]
        element_type =  df_aux['type'].iloc[i]

        list_connections_electric = [s for s in df_final.columns if '_' + element + '_' in s and 'P_' in s]
        list_connections_thermal = [s for s in df_final.columns if '_' + element + '_' in s and 'Q_' in s]

        folder_path = path_charts + '/' + element + '/'

        try:
            os.mkdir(folder_path)
        except FileExistsError:
            print(f"Folder '{element}' already exists at {folder_path}")
        except Exception as e:
            print(f"An error occurred: {str(e)}")

        plt.figure(figsize=(15, 8))
        bottom = np.zeros(len(x_axis))
        if len(list_connections_electric) > 0:
            for j in list_connections_electric:
                plt.bar(x_axis, df_final[j],bottom = bottom)
                bottom += df_final[j]
            plt.xlabel('Time [hs]')
            plt.ylabel('Power [kW]')
            plt.title(f"Electric power distribution - {element}")
            plt.legend(list_connections_electric)
            plt.savefig(folder_path + 'P_' + element + '.png')
            plt.close()


        plt.figure(figsize=(15, 8))
        bottom = np.zeros(len(x_axis))
        if len(list_connections_thermal) > 0:
            for j in list_connections_thermal:
                plt.bar(x_axis, df_final[j],bottom = bottom)
                bottom += df_final[j]
            plt.xlabel('Time [hs]')
            plt.ylabel('Power [kW]')
            plt.title(f"Thermal power distribution from - {element}")
            plt.legend(list_connections_thermal)
            plt.savefig(folder_path +'Q_' + element + '.png')
            plt.close()

charts_generator2(control,df_aux)