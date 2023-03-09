
class Point():
  x = 1
  y = 1


a_lst = []
for i in range(5):
  j = str(i)
  a = ('pt'+j)
  a_lst.append(a)

print (a_lst)  
  
for i in a_lst:
   i = Point()  
   print(i)