import os
import shutil
import pandas as pd
import streamlit as st
# Function to process seating allocation
def process_seating_allocation(input_file, buffer, dense_mode):
    ip_1 = pd.read_excel(input_file, sheet_name="ip_1")
    ip_2 = pd.read_excel(input_file, sheet_name="ip_2")
    ip_3 = pd.read_excel(input_file, sheet_name="ip_3")
    ip_4 = pd.read_excel(input_file, sheet_name="ip_4")

    max_fill = 1.0 if dense_mode else 0.5

    # Prepare data
    course_roll_mapping = ip_1.groupby("course_code")["rollno"].apply(list).to_dict()
    exam_schedule = ip_2.set_index("Date").to_dict(orient="index")
    rooms = ip_3.sort_values(by=["Block", "Exam Capacity"], ascending=[True, False])
    room_capacity = rooms.set_index("Room No.")["Exam Capacity"].to_dict()
    room_block = rooms.set_index("Room No.")["Block"].to_dict()

    initial_room_capacity = room_capacity.copy()
    seating_plan = []
    room_vacancy_details = []

    for date, schedule in exam_schedule.items():
        for session in ["Morning", "Evening"]:
            if pd.isna(schedule[session]) or schedule[session] == "NO EXAM":
                continue

            vacant_rooms_block9 = {room: capacity - buffer for room, capacity in room_capacity.items() if room_block[room] == 9}
            vacant_rooms_lt = {room: capacity - buffer for room, capacity in room_capacity.items() if room_block[room] != 9}
            courses = schedule[session].split("; ")
            courses.sort(key=lambda c: len(course_roll_mapping.get(c, [])), reverse=True)

            for course in courses:
                rolls = course_roll_mapping.get(course, [])
                total_students = len(rolls)
                if total_students == 0:
                    continue

                allocated_rolls = []
                # Allocate to rooms in Block 9 first
                for room, capacity in list(vacant_rooms_block9.items()):
                    max_alloc = int(initial_room_capacity[room] * max_fill)
                    alloc = min(total_students, max_alloc, vacant_rooms_block9[room])  # Ensure we don't exceed available capacity
                    allocated_rolls.extend(rolls[:alloc])
                    rolls = rolls[alloc:]
                    total_students -= alloc
                    seating_plan.append([date, schedule["Day"], session, course, room, alloc, ";".join(allocated_rolls)])
                    vacant_rooms_block9[room] -= alloc
                    if vacant_rooms_block9[room] <= 0:
                        del vacant_rooms_block9[room]
                    if total_students == 0:
                        break

                # If still students remain, allocate to other rooms
                if total_students > 0:
                    allocated_rolls = []
                    for room, capacity in list(vacant_rooms_lt.items()):
                        max_alloc = int(initial_room_capacity[room] * max_fill)
                        alloc = min(total_students, max_alloc, vacant_rooms_lt[room])  # Ensure we don't exceed available capacity
                        allocated_rolls.extend(rolls[:alloc])
                        rolls = rolls[alloc:]
                        total_students -= alloc
                        seating_plan.append([date, schedule["Day"], session, course, room, alloc, ";".join(allocated_rolls)])
                        vacant_rooms_lt[room] -= alloc
                        if vacant_rooms_lt[room] <= 0:
                            del vacant_rooms_lt[room]
                        if total_students == 0:
                            break

                # If there are still students remaining
                if total_students > 0:
                    course_roll_mapping[course] = rolls
                else:
                    del course_roll_mapping[course]

            # Keep track of room vacancy details
            room_vacancy_details.extend([
                [date, schedule["Day"], session, room, initial_room_capacity[room], room_block[room], vacant_rooms_block9.get(room, vacant_rooms_lt.get(room, 0))]
                for room in room_capacity
            ])

    op_1 = pd.DataFrame(seating_plan, columns=["Date", "Day", "Session", "course_code", "Room", "Allocated_students_count", "Roll_list"])
    op_1["Date"] = pd.to_datetime(op_1["Date"]).dt.date

    op_2 = pd.DataFrame(room_vacancy_details, columns=["Date", "Day", "Session", "Room No.", "Exam Capacity", "Block", "Vacant"])
    op_2["Date"] = pd.to_datetime(op_2["Date"]).dt.date

    return op_1, op_2

# Function to generate attendance sheets
def generate_attendance_sheets(op_1, roll_name_mapping):
    folder_name = "Attendance_Sheets"
    os.makedirs(folder_name, exist_ok=True)
    file_paths = []

    for _, row in op_1.iterrows():
        date = row["Date"].strftime("%d_%m_%Y")
        sub_code = row["course_code"]
        room = row["Room"]
        session = row["Session"].lower()
        rolls = row["Roll_list"].split(";")
        
        attendance_data = [{"Roll": roll, "Name": roll_name_mapping.get(roll, ""), "Signature": ""} for roll in rolls]
        for _ in range(5):  # Add empty rows for extra students
            attendance_data.append({"Roll": "", "Name": "", "Signature": ""})

        df = pd.DataFrame(attendance_data)
        file_name = f"{date}_{sub_code}_{room}_{session}.xlsx"
        file_path = os.path.join(folder_name, file_name)
        file_paths.append(file_name)
        df.to_excel(file_path, index=False)
    
    return file_paths

# Streamlit UI setup
def main():
    st.title("Seating Allocation and Attendance System")
    
    uploaded_file = st.file_uploader("Upload the Excel file", type="xlsx")
    
    buffer = st.slider("Set Buffer (Number of students to leave empty per room)", 0, 10, 0)
    dense_mode = st.checkbox("Enable Dense Mode (Allocate rooms fully)")

    if uploaded_file is not None:
        st.write("Processing the uploaded file...")
        
        ip_4 = pd.read_excel(uploaded_file, sheet_name="ip_4")
        roll_name_mapping = ip_4.set_index("Roll")["Name"].to_dict()
        
        op_1, op_2 = process_seating_allocation(uploaded_file, buffer, dense_mode)
        
        attendance_files = generate_attendance_sheets(op_1, roll_name_mapping)
        
        st.subheader("Seating Plan")
        st.dataframe(op_1)
        
        st.subheader("Room Vacancy Status")
        st.dataframe(op_2)
        
        st.subheader("View Attendance Sheets")
        selected_class = st.selectbox("Select a class", attendance_files)
        if selected_class:
            file_path = os.path.join("Attendance_Sheets", selected_class)
            if os.path.exists(file_path):
                st.write(f"Attendance Sheet: {selected_class}")
                df = pd.read_excel(file_path)
                st.dataframe(df)
            else:
                st.error("Selected file does not exist.")

        zip_file_path = "Attendance_Sheets.zip"
        with st.spinner("Zipping attendance sheets..."):
            shutil.make_archive("Attendance_Sheets", 'zip', "Attendance_Sheets")
        
        st.download_button(
            label="Download All Attendance Sheets as ZIP",
            data=open(zip_file_path, "rb").read(),
            file_name="Attendance_Sheets.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()
