import pandas as pd
from collections import defaultdict

# Load input Excel files (replace with your actual Excel file paths)
ip_1 = pd.read_excel('ip_1.xlsx')
ip_2 = pd.read_excel('ip_2.xlsx')
ip_3 = pd.read_excel('ip_3.xlsx')
# print(ip_1.columns)
# User input for arrangement mode
dense = int(input("Press 1 for sparse arrangement and 2 for dense arrangement: "))
buffer_size = int(input("Enter buffer size per classroom (default is 5): ") or 5)

# Step 1: Determine course size based on the number of students registered in each course
course_student_count = ip_1['course_code'].value_counts().to_dict()

# Sort rooms to prioritize Block 9 before LT
rooms_block_9 = ip_3[ip_3['Block'] == 9].sort_values(by='Exam Capacity', ascending=False)
rooms_LT = ip_3[ip_3['Block'] == 'LT'].sort_values(by='Exam Capacity', ascending=False)

# Step 2: Exam timetable - Extract course lists for each day/session
exam_schedule = defaultdict(lambda: {'Morning': [], 'Evening': []})
for _, row in ip_2.iterrows():
    date = row['Date']
    morning_courses = row['Morning'].split('; ') if row['Morning'] != "NO EXAM" else []
    evening_courses = row['Evening'].split('; ') if row['Evening'] != "NO EXAM" else []
    exam_schedule[date] = {'Morning': morning_courses, 'Evening': evening_courses}

# Step 3: Create seating plan
op_1_data = []
op_2_data = []

for date, sessions in exam_schedule.items():
    for session, courses in sessions.items():
        # Sort courses by size (large courses first)
        courses = sorted(courses, key=lambda x: course_student_count.get(x, 0), reverse=True)
        
        for course in courses:
            student_rolls = ip_1[ip_1['course_code'] == course]['rollno'].tolist()
            student_index = 0
            total_students = len(student_rolls)

            # Fill Block 9 first, then LT, as per requirement
            for rooms in [rooms_block_9, rooms_LT]:
                for _, room in rooms.iterrows():
                    room_capacity = room['Exam Capacity'] - buffer_size
                    if dense == 1:
                        max_course_seats = room_capacity // 2  # Sparse mode
                    else:
                        max_course_seats = room_capacity  # Dense mode

                    allocated_students = min(total_students - student_index, max_course_seats)

                    if allocated_students > 0:
                        roll_list = ";".join(student_rolls[student_index:student_index + allocated_students])
                        op_1_data.append([date, session, course, room['Room No.'], allocated_students, roll_list])
                        student_index += allocated_students

                    # Stop allocating if all students are seated
                    if student_index >= total_students:
                        break
                if student_index >= total_students:
                    break

# Step 4: Generate Room Summary (op_2)
for _, room in ip_3.iterrows():
    room_no = room['Room No.']
    exam_capacity = room['Exam Capacity']
    block = room['Block']
    vacant_seats = exam_capacity - (sum(row[4] for row in op_1_data if row[3] == room_no) + buffer_size)
    op_2_data.append([room_no, exam_capacity, block, max(0, vacant_seats)])

# Step 5: Save output to Excel
op_1_df = pd.DataFrame(op_1_data, columns=['Date', 'Day', 'Course Code', 'Room', 'Allocated Students Count', 'Roll List'])
op_2_df = pd.DataFrame(op_2_data, columns=['Room No.', 'Exam Capacity', 'Block', 'Vacant'])

with pd.ExcelWriter('output.xlsx') as writer:
    op_1_df.to_excel(writer, sheet_name='Seating Plan', index=False)
    op_2_df.to_excel(writer, sheet_name='Room Summary', index=False)

print("Seating plan and room summary generated as 'output.xlsx'.")
