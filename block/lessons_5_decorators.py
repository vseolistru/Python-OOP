def func():
    return 1

def hello(name = 'Some Name'):
    print('hello execute')

    def greet():
        print('\tHello '+name)

    def welcome():
        print(f'\tWelcome {name}')

    # greet()
    # welcome()
    if name == 'Some Name':
        return greet
    else:
        return welcome


def cool():
     def super_cool():
         return 'Super Cool'

     return super_cool

def new_hello():
    return 'Hello Some'

def other (some_func):
    print('----Execute other func----')
    print(some_func())

def new_decorator(original_func):
    def wrap_func():
        print('---before call original---')
        original_func()
        print('---after original call---')
    return wrap_func()


def func_needs_decorate():
    print('it needs a decoration')

res = func

print (hex(id(res())))
print(hex(id(func())))
print('del func')
del func
try:
    print(id(func()))
except:
    print('func was del')

print (res())

print('---new block---')

result = hello('John')

print(result())

new_res = cool()

print(new_res())

other(new_hello)

print('+++wrap+++')

new_decorator(func_needs_decorate)


