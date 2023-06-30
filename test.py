class MyClass:
    pass

my_instance = MyClass()

instance_class = type(my_instance).__name__
print(instance_class)  # Output: "MyClass"