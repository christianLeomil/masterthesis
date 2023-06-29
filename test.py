import pandas as pd

name_file = 'test.xlsx'
path_input = './input/'
path_output = './output/'

def my_function(parameter1=None, parameter2=None):
    if parameter1 is not None:
        parameter2 = parameter1
    else:
        parameter2 = 6
    return parameter2

try:
    df = pd.read_excel(path_input + name_file)
    parameter1 = df['parameter1'].iloc[0] 
    parameter2 = my_function(parameter1)
except Exception:
    parameter2 = my_function()

print(parameter2)