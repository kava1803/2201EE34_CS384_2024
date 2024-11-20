
n = int(input("Enter Number of Strings : "))
Dict = {"z": ["aa", "fdsa"]}

for i in range(n):
    str_input = input(f"String {i + 1} : ")
    s = str_input
    temp = tuple(sorted(s))
    if temp not in Dict:
        list = [str_input]
        Dict[temp] = list
    else:
        Dict[temp].append(str_input)

Dict.pop("z")
for hash in Dict:
  print(Dict[hash])