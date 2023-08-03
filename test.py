class MyClass:
    def __init__(self, name_of_instance):
        self.name_of_instance = name_of_instance
        self.lista_teste = [10,20,30,40]
        setattr(self, f"{self.name_of_instance}_parameter1", 3)
        setattr(self, f"{self.name_of_instance}_parameter2", self.lista_teste)
        setattr(self, f"{self.name_of_instance}_parameter3", self.test_function())

    def test_function(self):
        return getattr(self, f"{self.name_of_instance}_parameter1") * getattr(self, f"{self.name_of_instance}_parameter2")

myClass = MyClass()
print(myClass.myClass_parameter3)  # Output: 50