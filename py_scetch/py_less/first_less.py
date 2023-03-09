class Point:
	"New Class"
	x = 1 #атрибуты
	y = 1


class Point3D:
	x = 2
	y = 4
	z = 6


pt = Point()
pt.x = 7
pt.y = 9
print(pt.__doc__)
#print(pt.__name__)
#print(dir(pt))

Point.x = 100
print(pt.x, pt.y)
print (id(pt), id(Point), sep = '\n')
print(pt.__dict__)
print(getattr(pt, "x") )
print(getattr(pt, "z", False) )
print(hasattr(pt, 'x') )
setattr(pt, 'z', 12)
print(getattr(pt, 'z') )
delattr(pt,'z')
print(hasattr(pt, 'z') )


print(isinstance(pt, Point))

new3D = Point3D()
obj3D = Point3D()
sam3D = Point3D()


print(new3D.x, new3D.y, new3D.z)
setattr(Point3D, 'a', 13)
print(obj3D.x, obj3D.y, obj3D.z)
print(Point3D.a, Point3D.y, Point3D.z)
print(Point3D.__dict__ )