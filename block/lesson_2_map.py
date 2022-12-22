# map принимает на вход коллбек и масссив


def square (arg):
    return arg**2

def splicer (my_str):
    if len(my_str) % 2 == 0:
        return 'EVEN'
    else:
        return my_str[0]


my_nums = [1,2,3,4,5]
res = [i for i in map(square, my_nums)]


names = ['Andy', 'Serg', 'DonVito']
result = [i for i in map(splicer, names)]

print(res)
print(result)