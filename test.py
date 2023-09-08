import pandas as pd

df1 = pd.DataFrame({'Parameter':['String 1','String 2','String 3'],
                    'Value':[1,2,4,]})

df2 = pd.DataFrame({'Parameter':['var__starting_index','be__starting_index','can__starting_index'],
                    'Value':[10, 20, 30]})

def teste(df1,df2):
    for j in df2.index:
        if df2['Parameter'].iloc[j] in df1['Parameter'].tolist():
            df1.loc[df1['Parameter']==df2['Parameter'].iloc[j], 'Value'] = df2['Value'].iloc[j]
        else:
            row_to_append = df2.loc[j]
            df1 = df1.append(row_to_append, ignore_index = True)
    return df1


lista = [1,2,3]

for i in lista:
    print('------------rodada: ' + str(i))
    df2 = pd.DataFrame({'Parameter':['var__starting_index','be__starting_index','can__starting_index'],
                    'Value':[1 * i, 2 * i, 3 * i]})
    df1 = teste(df1,df2)

path_output = './output/teste/'
df1.to_excel(path_output + 'df1_teste.xlsx')


# import pandas as pd

# def breaking_dataframe(df,horizon,saved_position):
#     list_split = []
#     for i in range(0,len(df),saved_position):
#         # print('\n------------'+str(i))
#         if i == 0:
#             list_split.append(df.iloc[i : i + horizon])
#         else:
#             list_split.append(df.iloc[i-1 : i+horizon])
#         # print(list_split[-1])
#     return list_split

# horizon = 5
# saved_position = 4

# path_input = './input/'
# path_output = './output/'

# df = pd.read_excel(path_input + 'df_input.xlsx', sheet_name = 'series', nrows = 30)

# for i,df in enumerate(breaking_dataframe(df,horizon,saved_position)):
#     df.to_excel(path_output + '/teste/df_split' + str(i) + '.xlsx')