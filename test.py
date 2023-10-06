import pandas as pd
path_output = './output/'

df_con_electric = pd.read_excel(path_output + 'df_con_electric.xlsx', index_col= 0)
df_con_electric.index.name = None
df_con_thermal = pd.read_excel(path_output + 'df_con_thermal.xlsx', index_col = 0)
df_con_thermal.index.name = None


def compensation_constraints_creator(df_con_electric,df_con_thermal):

    df_con_electric = df_con_electric.filter(like = 'P_to_net', axis = 0)
    df_con_electric.columns = [s.replace('P_from_','') for s in df_con_electric.columns]
    df_con_electric.index = [s.replace('P_to_','') for s in df_con_electric.index]


    df_con_thermal = df_con_thermal.filter(like = 'Q_to_net', axis = 0)
    df_con_thermal.columns = [s.replace('Q_from_','') for s in df_con_thermal.columns]
    df_con_thermal.index = [s.replace('Q_to_','') for s in df_con_thermal.index]

    #looping through the connection matrices and checking which elements are connected to the grid
    list_expressions = []
    list_variables = []

    #looping through electric matrix
    for i in df_con_electric.index:
        for j in df_con_electric.columns:
            if df_con_electric.loc[i,j] != 0:
                energy_flux = df_con_electric.loc[i,j]
                #appending expressions and variables related to compensation values
                list_expressions.append(f"model.P_comp_{j}[t] == model.{energy_flux}[t] * model.param_{j}_compensation",)
                list_variables.append(f"P_comp_{j}")
                #appending expressions and variables related to revenue values
                list_expressions.append(f"model.P_rev_{j}[t] == model.{energy_flux}[t] * model.param_{i}_cost_sell_electric")
                list_variables.append(f"P_rev_{j}")

    #looping through thermal matrix
    for i in df_con_thermal.index:
        for j in df_con_thermal.columns:
            if df_con_thermal.loc[i,j] != 0:
                energy_flux = df_con_thermal.loc[i,j]
                #appending expressions and variables related to compensation values
                list_expressions.append(f"model.Q_comp_{j}[t] == model.{energy_flux}[t] * model.param_{j}_compensation",)
                list_variables.append(f"Q_comp_{j}")
                #appending expressions and variables related to revenue values
                list_expressions.append(f"model.Q_rev_{j}[t] == model.{energy_flux}[t] * model.param_{i}_cost_sell_thermal")
                list_variables.append(f"Q_rev_{j}")

    df_expressions = pd.DataFrame({'expressions':list_expressions,
                                   'variables':list_variables})

    return df_expressions

df_expressions = compensation_constraints_creator(df_con_electric, df_con_thermal)
df_expressions = pd.DataFrame(df_expressions)
df_expressions.to_excel(path_output + 'df_expressions.xlsx',index = False)

