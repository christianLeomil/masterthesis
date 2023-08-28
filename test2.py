import pandas as pd

df = pd.DataFrame({'teste1':[1,2,3,4,5],
                   'teste2':[10,20,30,40,50]})


list_teste = df['teste2'].tolist()

for i,n in enumerate(list_teste):
    print('\n')
    print(i)
    print(n)