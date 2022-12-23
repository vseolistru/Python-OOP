# getattr, setattr, delattr

class Person:
    name = 'Ivan'

    def hello():
        print('Hello ')


# print(Person.name)
# Person.age = 23#
# print(getattr(Person, 'age'))
#
# setattr(Person, 'marige', True)#
# print(Person.marige)
#
# delattr(Person, 'name')#
# print(Person.__dict__)
# Person.hello()

class Dog():
    #atribute breed of Dog Class


    def __int__(self, breed, name, spots):
        self.breed = breed
        self.name = name
        self.spots = spots


my_dog = Dog()
my_dog.breed = 'lab'
my_dog.name = 'Tobi'
my_dog.spots = True
print(my_dog)