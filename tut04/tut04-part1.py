# Question 1

# Merge Sort Function
"""

# Merge Sort Function
def merge_sort(arr):
    if len(arr) > 1:
        mid = len(arr) // 2
        L = arr[:mid]
        R = arr[mid:]

        merge_sort(L)
        merge_sort(R)

        i = j = k = 0

        while i < len(L) and j < len(R):
            if float(L[i][1]) < float(R[j][1]):
                arr[k] = L[i]
                i += 1
            else:
                arr[k] = R[j]
                j += 1
            k += 1

        while i < len(L):
            arr[k] = L[i]
            i += 1
            k += 1

        while j < len(R):
            arr[k] = R[j]
            j += 1
            k += 1

"""## Function 1"""

def New_Student(Dict_Student_Grades):
    name = input("Enter name of new Student: ")
    name = name.lower()
    grade_count = int(input(f"Enter number of grades for {name}: "))
    grades = []
    for j in range(grade_count):
        grade = input(f"Enter grade {j+1} for {name}: ")
        grades.append(grade)
    Dict_Student_Grades[name] = grades
    return

"""## Function 2"""

def Update_Grade(Dict_Student_Grades):
    name = input("Enter name of the student whose grades are to be changed: ")
    name = name.lower()
    if name not in Dict_Student_Grades:
        print("Invalid Name\nStudent Not Present in Database")
        return
    grade_count = int(input(f"Enter number of new grades for {name}: "))
    new_grades = []
    for j in range(grade_count):
        new_grade = input(f"Enter new grade {j+1} for {name}: ")
        new_grades.append(new_grade)

    Dict_Student_Grades[name] = new_grades
    return

"""## Function 3"""

#Display the averages
def Display_Average(Dict_Student_Grades):
  for name in Dict_Student_Grades:
    print(name," : ",Dict_Student_Grades[name])

"""## Function 4"""

# Display Sorted Averages Function
def Display_Sorted(Dict_Student_Averages):
    list_of_tuples = list(Dict_Student_Averages.items())
    merge_sort(list_of_tuples)
    for name, average in list_of_tuples:
        print(f"Name: {name} \t Average Grade: {average}")

"""## Main"""

#Main
Dict_Student_Grades = {}
Dict_Student_Averages = {}

def calculate_average(grades):
    total = sum(float(grade) for grade in grades)
    return total / len(grades) if grades else 0

def update_averages():
    for student, grades in Dict_Student_Grades.items():
        Dict_Student_Averages[student] = calculate_average(grades)

n = int(input("Enter Number of students to be Entered: "))
for i in range(n):
    name = input("Enter Student Name: ").lower()
    grade_count = int(input(f"Enter number of grades for {name}: "))
    grades = []
    for j in range(grade_count):
        grade = input(f"Enter grade {j+1} for {name}: ")
        grades.append(grade)
    Dict_Student_Grades[name] = grades

update_averages()

while True:
    func_id = int(input('''Enter Which function to be executed :
    \n 1) Add a new Student
    \n 2) Update grades of a pre-existing student
    \n 3) Display average grade
    \n 4) Display Sorted Students
    \n 5) Exit
    \n'''))
    if func_id == 1:
        New_Student(Dict_Student_Grades)
        update_averages()
    elif func_id == 2:
        Update_Grade(Dict_Student_Grades)
        update_averages()
    elif func_id == 3:
        Display_Average(Dict_Student_Averages)
    elif func_id == 4:
        Display_Sorted(Dict_Student_Averages)
    elif func_id == 5:
        print("Exiting Program.....")
        break
    else:
        print("Invalid Input")
print("Program Exited")