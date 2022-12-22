#filter вернет только те элементы для которых коллбек вернет true не мутирует

def check_even (arg) :
    return arg % 2 == 0

my_list  = [2,3,6,7,8,9,4,5]

res = [i for i in filter(check_even, my_list)]

print(res)


