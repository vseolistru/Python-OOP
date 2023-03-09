class Point:
  WIDTH = 5

  __slots__ = ['__x', '__y','z'] #разрешенные поля для экземпляров класса

  def __init__(self, x = 0, y = 0):
    self.__x =x; self.__y = y

  def __checkValue(x):
        if (isinstance(x,int) or isinstance(x, float)):

           return True   
        return False
           
 
  def setCoords(self, x ,y):  
      if Point.__checkValue(x) and Point.__checkValue(y):
        self.__x = x 
        self.__y = y
      else:
        print("Координаты должны быть числами")     

  def getCoords (self):
    return self.__x, self.__y

  def __getattribute__(self, item):
    if item =="_Point__x":
      return "Частная переменная"
    else:
      return object.__getattribute__(self, item)    

  def __setattr__(self, key, value):
     if key == "WIDTH":
       raise AttributeError
     else:
       self.__dict__[key] = value  

  def __getattr__(self, item):
    print ("__getattr__:" +item)  

  def __delattr__(self, item):
    print("__delattr__:"+ item)     

pt = Point(1,2)
print(pt.getCoords() )
pt.setCoords('10',5)
print(pt._Point__x)
print(pt.getCoords() )
#pt.WIDTH = 3
pt.zzz
pt.z = 1
del pt.z
