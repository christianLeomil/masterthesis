import pandas as pd

df1 = pd.DataFrame({'Parameter':['String 1','String 2','String 3'],
                    'Value':[1,2,5]})

df2 = pd.DataFrame({'Parameter':['String 1','String 2','String 3'],
                    'Value':[1,2,7]})

df3 = pd.concat([df1,df2],axis = 1)

print(df3)
# df3.to_excel('./output/teste/testeconcat.xlsx')

for i in range(0,4.0):
    print(i)