from lesson_1_atributes_class import Person, Dog


p1 = Person

# p1.name = 'Andrey'


print(id(p1.name))
print(id(Person.name))
print(p1.__dict__)


