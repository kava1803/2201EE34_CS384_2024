def Question1(s :str)->int:
  sum = 0
  size = len(s) - 1
  while len(s)!=1:
    length = len(s)-1
    sum = 0
    while length!=-1:
      sum += int(s[length])
      length -= 1
    s = f'{sum}'
  return sum

s1 = input("Enter String for question 1 : ")
print(f"Unitary Sum of {s1} is : ",Question1(s1))