# Base class (Parent class)
class Vehicle:
    def __init__(self, brand, model, year):
        self.brand = brand
        self.model = model
        self.year = year

    def start_engine(self):
        print(f"{self.year} {self.brand} {self.model}'s engine is starting.")

    def stop_engine(self):
        print(f"{self.year} {self.brand} {self.model}'s engine is stopping.")

# Derived class 1 (Child class)
class Car(Vehicle):
    def __init__(self, brand, model, year, fuel_type):
        super().__init__(brand, model, year)
        self.fuel_type = fuel_type

    def honk(self):
        print(f"{self.year} {self.brand} {self.model} honks!")

# Derived class 2 (Child class)
class Motorcycle(Vehicle):
    def __init__(self, brand, model, year, engine_type):
        super().__init__(brand, model, year)
        self.engine_type = engine_type

    def wheelie(self):
        print(f"{self.year} {self.brand} {self.model} performs a wheelie!")

# Create instances of the derived classes
car = Car("Toyota", "Camry", 2022, "Gasoline")
motorcycle = Motorcycle("Honda", "CBR600RR", 2023, "Liquid-cooled")

# Access methods and attributes of the instances
car.start_engine()
car.honk()
car.stop_engine()

motorcycle.start_engine()
motorcycle.wheelie()
motorcycle.stop_engine()
