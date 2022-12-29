def new_dec():
    def hello_func():
        print('hello_new')
    hello_func()

def hello_decorator(hello):
    print('===before')
    hello()
    print('+++after')
    return hello


def decorator(func):
    def wrapper(*args, **kwargs):
        print(f'++++before {func}')
        res = func(*args, **kwargs)
        print(f'---after {func}')
        return res
    return wrapper


@decorator
def hello(name):
    return (f'Hello {name}')


print(hello('Rogozka'))