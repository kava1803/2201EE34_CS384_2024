# import pandas as pd

# df = pd.read_csv(r"C:\Users\agrvi\Downloads\allocations (4).xlsx")


# room_no = []
# roll_no = {}

# dates = 
# unique_room = df['Room'].tolist()

# for _ in unique_room:
#     roll_no[_] = df[df[_]]['Room_lsit'].tolist()



# room_no , roll no



import pandas as pd
file_path = r"C:\Users\agrvi\Downloads\allocations (4).xlsx"
data = pd.ExcelFile(file_path)
# seating_plan = data.parse('Seating Plan')
attendance_data = []

for _, row in seating_plan.iterrows():
    date = row['Date']
    room_no = row['Room']
    roll_nos = row['Roll List'].split(';')
    
    for roll_no in roll_nos:
        attendance_data.append({
            'Date': date,
            'Room No.': room_no,
            'Roll No.': roll_no,
            'Name': 'N/A',
            'Signature': '' 
        })

# Convert to a DataFrame
attendance_df = pd.DataFrame(attendance_data)

# Save to a new Excel file
output_path = r"C:\Users\agrvi\Downloads\attendence.xlsx"
attendance_df.to_excel(output_path, index=False)
