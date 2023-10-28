# class YourClass:
#     def __init__(self, string_list):
#         for string in string_list:
#             setattr(self, string, None)
#             # You can initialize the variables to a specific value if needed, e.g., setattr(self, string, 0)

# # Example usage:
# string_list = ["name", "age", "city"]
# your_instance = YourClass(string_list)

# print(your_instance.name)  # This will print None since we didn't initialize it
# print(your_instance.age)
# print(your_instance.city)


import numpy as np

lista = [1,2,3,4,5,6,7,8,9]
lista1 = np.array(lista) * 100
# lista1 = lista1[:8]
lista2 = np.array(lista) * 10
lista3 = ['c','n','p']

print(lista1)

for i,n,m in zip(lista1,lista2,lista3):
    print('\n------------------')
    print(i)
    print(n)
    print(m)