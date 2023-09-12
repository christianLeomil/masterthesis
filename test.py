import pandas as pd

df1 = pd.DataFrame({'Parameter':['String 1','String 2','String 3'],
                    'Value':[1,2,4,]})

df2 = pd.DataFrame({'Parameter':['String 1','String 2','String 3'],
                    'Value':[1,2,4,]})

df3 = pd.DataFrame()

# for i in df1.index:
#     df3.append(df1.loc[i],ignore_index = True)
# df3 = df3.append(df1.loc[1],ignore_index = True)
df3 = pd.concat([df3, df2.loc[0:1]], ignore_index = True)

df4 = pd.DataFrame({'Value':['String 2','String 1','String 3'],
                    'Parameter':[10,20,40,]})

df3 = pd.concat([df3, df4.loc[0:3]], ignore_index = True)

print(df3)
