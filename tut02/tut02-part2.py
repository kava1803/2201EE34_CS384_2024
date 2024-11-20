
def Question2(s : str)->str:
  count = 0
  ans = ""
  for i in range(0,len(s)):
    if i>0 and s[i]!=s[i-1]:
      ans += s[i-1] + f'{count}'
      count = 0
    count += 1
  ans += s[len(s)-1] + f'{count}'
  return ans

s2 = input("Enter String for question 2 : ")
print(f"Compressed string of {s2} is : ",Question2(s2))