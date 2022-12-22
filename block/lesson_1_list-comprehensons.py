mystr = 'hello'
res = []

for  i in mystr:
    res.append(i)

print(res)

result = [i for i in mystr]

gen = [i**2 for i in range(3,12)]

new_res = [i for i in range(2,9) if i % 2 == 1]

celcius = [0,5,10,15,20,25,30,35.5]

far = [((9/5)*temp+32) for temp in celcius ]

new_result = [i if i % 2 == 0 else 'ODD' for i in range(2,9)]



print(result)

print(gen)

print(new_res)

print(far)

print(new_result)