class User():
    # private arg name = _name
    def __init__(self, name = 'dafault'):
        self._name = name
    @property
    def name(self):
        print('instance_name', self._name)
        return self._name

    @name.setter
    def name(self, value):
        self._name = value
        print('name = value', self._name)

    @name.deleter
    def name(self):
        print('delete', self._name)
        del self._name

a = User('Hello')
a.name
a.name = 'h1' #setter
del a.name


