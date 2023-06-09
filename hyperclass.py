import pandas as pd

def define_hyper_class(path_input,path_output):
    df_hyper = pd.read_excel(path_input + 'df_input.xlsx',sheet_name='hyperclass')
    return df_hyper

