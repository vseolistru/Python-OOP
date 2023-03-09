

   

    

  

l1 = ['ab','ac','ad','am']
l3 = ['ab','ac','ad','as','af']
"""
a=[]

if l1[0] != l3[0]:
  a.append(l3[0])
if l1[1] != l3[1]:  
  a.append(l3[1])
if l1[2] != l3[2]:  
  a.append(l3[2])
if l1[3] != l3[3]:  
  a.append(l3[3])

print (a)
#res = method_1(l1, l3)
"""
res = [x for x in l3 if x not in l1 or x not in l3]
print (res)

