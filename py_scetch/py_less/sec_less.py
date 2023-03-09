import random

class Point:
	
  """Конструктор Класса"""
  def __init__(self, x = 0,y = 0):

    self.x = x
    self.y = y
    x = 1; y = 1
  """
  def __del__(self):
    #Деструктор класса  
    print('Удаление экземпляра класса'+self.__str__())
	"""
  def setCoords(self, x, y):	
    self.a = x
    self.b = y
  
  def __str__(self):
    return f"{self.x}, {self.y}"


class Point3d:

  def __init__(self, lst_a):
    self.lst_a = lst_a
    lst_a = [2,4,6]
   j = str(i)
  a = ('pt'+j)
  x = random.randint(0, 20)
  y = random.randint(0, 15)
  a = Point(x,y)
  print(a)
  lst_obj.append(a)


lst_b =[5,7,9]
pt3d = Point3d.setCoords(lst_b)
print (pt3d)

