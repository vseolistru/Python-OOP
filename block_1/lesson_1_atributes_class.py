# getattr, setattr, delattr

class Person:
    name = 'Ivan'
    age = 18

    def __init__(self, name, age):
        self.name = name
        self.age = age
        self.get_name()

    def get_name(self):
        print(self.name)

    def get_age(self):
        print(self.age)


class Dog():
    #atribute breed of Dog Class


    def __int__(self, breed, name, spots):
        self.breed = breed
        self.name = name
        self.spots = spots

class User():

    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        print(self.args)
        print(self.kwargs)


a = User(*[1,2,3], **{'name':'Ikar'})
print(a)


p1 = Person('Ikarito', 24)
p1.get_age()


my_dog = Dog()
my_dog.breed = 'lab'
my_dog.name = 'Tobi'
my_dog.spots = True
print(my_dog)