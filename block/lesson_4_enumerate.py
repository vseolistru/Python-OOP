
lst = [10,20,30,40,50]
some_str = 'hello'
dome_brands = ['apple', 'sums', 'huaiway']
some_obj = {'a':3, 'b':4, 'c':7}


# [(0, 10), (1, 20), (2, 30), (3, 40), (4, 50)]
print(list(enumerate(lst)))

for i, v in enumerate(some_obj, 10):
    print(i, v)