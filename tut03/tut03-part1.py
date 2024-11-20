# Question 1

def sieve(digits):
  n = pow(10,digits+1);
  n = n-1
  primes = [1]*(n+1)
  primes[0] = 0
  primes[1] = 1
  for i in range(2,int(n**0.5),1):
    if primes[i]:
      for j in range(i*i,n+1,i):
        primes[j] = 0
  return primes

def next_permutation(n,digits):
  last_digit = n%10;
  n = n//10
  n = n + last_digit * int(pow(10,digits-1))
  return n

def Question_1():
  n = int(input("Enter Number to be checked : "))
  num = n
  ans = True
  temp = n
  digits = 0
  while(temp):
    digits +=1
    temp = temp//10
  primes = sieve(digits)
  if primes[n]:
    rotational_prime = True
    for i in range(digits):
      n = next_permutation(n,digits)
      if primes[n]==0:
        rotational_prime = False
        print(f"{num} is not a Rotational Prime!")
        break
    if rotational_prime :
      print(f"{num} is a Rotational Prime!")
  else :
    print(f"{num} is not a Rotational Prime!")


Question_1()