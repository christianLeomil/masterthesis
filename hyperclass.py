import pandas as pd

def define_hyper_class(path_input,path_output):
    df_hyper = pd.read_excel(path_input + '20230606-df_input-CP_v1.xlsx',sheet_name='hyperclass')
    # df_hyper.set_index('Parameter',  inplace = True)
    return df_hyper


# path_output = './output/'
# path_input = './input/'

# df_hyper = define_hyper_class(path_input,path_output)
# print(df_hyper)
# print('---------------------\n')
# print(df_hyper.loc[0:1,['Parameter','Value']])
# var = df_hyper.loc[1,'Parameter']
# print('---------------------\n')
# print(type(var))
