# # Parent class (base class)
# class Animal:
#     def __init__(self, name, walking_speed):
#         self.name = name
#         self.walking_speed = walking_speed

#     def speak(self):
#         pass
    
#     def walk(self):
#         return f"{self.name} can walk {self.walking_speed}!"
    
#     def __private_method(self):
#         return f"{self.name} has a private method"

# # Child class (derived class) inheriting from Animal
# class Dog(Animal):
#     def __init__(self, name, walking_speed, eating_amount):
#         super().__init__(name, walking_speed)
#         self.eating_amount = eating_amount

#     def speak(self):
#         return f"{self.name} says Woof!"
    
#     def walk(self):
#         return f"{self.name} likes to run like hell"
    
#     def eating(self):
#         return f"{self.name} eats {self.eating_amount} of food"

# # Child class (derived class) inheriting from Animal
# class Cat(Animal):
#     def speak(self):
#         return f"{self.name} says Meow!"

# # Create instances of the child classes
# dog = Dog("Buddy",'fast','a lot')
# cat = Cat("Whiskers",'slow')

# # Call the speak method on the instances
# print(dog.speak())  # Output: Buddy says Woof!
# print(cat.speak())  # Output: Whiskers says Meow!

# print(dog.walk())  
# print(cat.walk())

# print(dog.eating())


class Animal:
    def __init__(self, name):
        self.name = name

    def speak(self):
        return "Some generic animal sound"

    def _protected_method(self):
        return "This is a protected method in the Animal class"

class Dog(Animal):
    def speak(self):
        return f"{self.name} says Woof!"

    def fetch(self):
        return f"{self.name} is fetching a ball"

class Cat(Animal):
    def speak(self):
        return f"{self.name} says Meow!"

    def chase_mice(self):
        return f"{self.name} is chasing mice"

    def _protected_method(self):
        return f"This is a protected method in the Cat class for {self.name}"

# Create instances of the child classes
dog = Dog("Buddy")
cat = Cat("Whiskers")

# Call public methods
print(dog.speak())       # Output: Buddy says Woof!
print(cat.speak())       # Output: Whiskers says Meow!
print(dog.fetch())       # Output: Buddy is fetching a ball
print(cat.chase_mice())  # Output: Whiskers is chasing mice

# Attempt to call the protected method from both classes
print(cat._protected_method())  # Output: This is a protected method in the Cat class for Whiskers
print(dog._protected_method())
# Calling this from dog will raise an AttributeError
# print(dog._protected_method())  # This will raise an AttributeError
