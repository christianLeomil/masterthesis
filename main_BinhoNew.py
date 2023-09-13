import pandas as pd
import pyomo.environ as pyo
import classes_Binho
import inspect

path_output = './output/'
path_input = './input/'
name_file = 'df_input.xlsx'

df_input_series = pd.read_excel(path_input + name_file, sheet_name='series')
df_input_other = pd.read_excel(path_input + name_file,sheet_name = 'other')

df_elements = pd.read_excel(path_input + name_file, sheet_name = 'elements')
df_conect = pd.read_excel(path_input + name_file, sheet_name = 'connect',index_col=0)

df_input_pv = pd.read_excel(path_input + name_file, sheet_name='pv')
df_input_bat = pd.read_excel(path_input + name_file, sheet_name='bat')