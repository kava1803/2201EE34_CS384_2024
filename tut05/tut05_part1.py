# Question 2"""

s = input("Enter string of brackets : ")
list = []
balanced = True;
for i in s:
    if i == "(" or i == "[" or i == "{":
        list.append(i)
    else :
      top = list[-1]
      if (i == ")" and top == "(") or (i == "]" and top == "[") or (i == "}" and top == "{"):
        list.pop();
      else :
        balanced = False;
        break;
if len(list) == 0 and balanced == True:
  print("Balanced")
else :
  print("Not Balanced")