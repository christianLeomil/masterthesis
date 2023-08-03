class MyClass:
    instances = {}  # Dictionary to store instances with custom names

    def __init__(self, name_of_instance):
        self.name_of_instance = name_of_instance
        MyClass.instances[self.name_of_instance] = self  # Store the instance with the custom name
        setattr(self, f"{self.name_of_instance}_parameter1", 10)
        setattr(self, f"{self.name_of_instance}_parameter2", 5)
        setattr(self, f"{self.name_of_instance}_parameter3", self.test_function())

    def test_function(self):
        return getattr(self, f"{self.name_of_instance}_parameter1") * getattr(self, f"{self.name_of_instance}_parameter2")

# Create instances with custom names
myInstance = MyClass('myInstance')
anotherInstance = MyClass('anotherInstance')

# Access instance attributes
print(myInstance.myInstance_parameter1)  # Output: 50
print(anotherInstance.anotherInstance_parameter3)  # Output: 50

class MyClass:
    def __init__(self):
        self.parameter1 = 10
        self.parameter2 = 20
        self.name_of_this_instance = self.return_name_of_instace()

    def return_name_of_instace(self):
        return # function that gets the name of the instance that is going to be created from this class