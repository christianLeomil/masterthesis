class DynamicClass:
    def __init__(self, n):
        self.n = n

        for i in range(n):
            def method(self, num):  # Define the method inside the loop
                print(f"Method {num} called!")
            
            # Assign the method to the instance
            setattr(self, f"method_{i}", method.__get__(self, DynamicClass))

# Create an instance of DynamicClass
dynamic_obj = DynamicClass(3)

# Call the dynamically created methods
dynamic_obj.method_0(0)
dynamic_obj.method_1(1)
dynamic_obj.method_2(2)
