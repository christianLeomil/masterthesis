# class DynamicClass:
#     def __init__(self, n):
#         self.n = n

#         for i in range(n):
#             def method(self, num):  # Define the method inside the loop
#                 print(f"Method {num} called!")
            
#             # Assign the method to the instance
#             setattr(self, f"method_{i}", method.__get__(self, DynamicClass))

# # Create an instance of DynamicClass
# dynamic_obj = DynamicClass(3)

# # Call the dynamically created methods
# dynamic_obj.method_0(0)
# dynamic_obj.method_1(1)
# dynamic_obj.method_2(2)

# import pandas as pd

# path = './input/'
# name_file = 'c_matrix.xlsx'

# df = pd.read_excel(path + name_file)
# df.set_index('from', inplace=True)
# df.index.name = None

# class dynamic_class:
#     def __init__(self,df):
#         self.df = df

#         for i in range(0,len(df)):
#             for j in df.columns:
#                 if df[j].iloc[i] == 1:
#                     name = str(df.index[i]) + '_' + str(j)
#                     print(name)
#                     def dynamic_method(self,num):
#                         return print(num*10)       
#                     setattr(self,name,dynamic_method.__get__(self, dynamic_class))

# dynamic_obj = dynamic_class(df)
# dynamic_obj.bat_demand(4)

# method_list = [method for method in dir(dynamic_obj) if method.startswith('__') is False]
# print(method_list)


# print(eval("2 ** 8"))
# eval("pi = 3.1416")

vector = ['pv','pv_bat','pv_demand']

class teste1:
    def teste(vector):
        globals()[vector[0]] = globals()[vector[1]] +globals()[vector[2]]

teste2 = teste1

method_list = [method for method in dir(teste2) if method.startswith('__') is False]
print(method_list)