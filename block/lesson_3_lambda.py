# lambda позволяет создавать аннонимные функции
# wrong
sq = lambda arg : arg ** 2

my_str = ['python', 'kotlin', 'jscript']
lst = [22,33,44,55,66, 7,6,11,13]

# right
print( [i for i in map(lambda arg : arg ** 2, lst)] )
print ( [i for i in filter(lambda arg: arg % 2 == 0, lst)] )
print( [i for i in map(lambda arg: arg[0:3], my_str)] )