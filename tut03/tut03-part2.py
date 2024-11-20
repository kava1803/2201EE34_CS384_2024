"""# Question 2

"""

def nextGreaterPermutation(s: str) -> str:
    A = list(s)
    n = len(A)
    ind = -1
    for i in range(n-2, -1, -1):
        if A[i] < A[i + 1]:
            ind = i
            break
    if ind == -1:
        return ''.join(reversed(A))
    for i in range(n - 1, ind, -1):
        if A[i] > A[ind]:
            A[i], A[ind] = A[ind], A[i]
            break
    A[ind+1:] = reversed(A[ind+1:])
    return ''.join(A)

def fact(n):
  return n*fact(n-1) if n>1 else 1

def Question_2():
  s = input("Enter your String : ")
  length = len(s)
  factorial = fact(length)
  for i in range(factorial):
    print(s)
    s = nextGreaterPermutation(s)

Question_2()