import pandas as pd
import numpy as np

input_file = r''  # Update with the correct file path
df = pd.read_excel(input_file)

max_marks = df.iloc[0, 2:].astype(float)  
weightage = df.iloc[1, 2:].astype(float) / 100 

# Remove the first two rows as they are metadata now
df = df.drop([0, 1]).reset_index(drop=True)

# Convert necessary columns to numeric type for calculation
df.iloc[:, 2:] = df.iloc[:, 2:].apply(pd.to_numeric)

# Calculate weighted total for each student
df['Grand Total/100'] = sum(
    (df[col] / max_marks[col] * weightage[col] * 100) for col in max_marks.index
)

# Define grading criteria based on 'Grand Total/100'

df['Grade'] = df['Grand Total/100'].apply(assign_grade)

# Define 'Total Students' and count students in each grade for extra columns
total_students = len(df)
grade_counts = df['Grade'].value_counts().reindex(['AA', 'AB', 'BB', 'BC', 'CC', 'CD', 'DD', 'F'], fill_value=0).astype(int)
df['Total Students'] = total_students

# Additional columns for "grade", "old iapc reco", "Counts", "Round", "Count verified"
df['grade'] = df['Grade']  # Assuming same grade for simplicity
df['old iapc reco'] = grade_counts[df['Grade']].values
df['Counts'] = grade_counts[df['Grade']].values * 1.02  # Adjust this factor if needed
df['Round'] = np.round(df['Counts'])
df['Count verified'] = df['Round'].astype(int)

# Sort the dataframe by roll number and by Grand Total/100
df_roll_sorted = df.sort_values(by='Roll')
df_marks_sorted = df.sort_values(by='Grand Total/100', ascending=False)

# Write both sorted DataFrames to a single Excel file with separate sheets
output_excel_file = 'output_combined.xlsx'
with pd.ExcelWriter(output_excel_file) as writer:
    df_roll_sorted.to_excel(writer, sheet_name='Sorted by Roll Number', index=False)
    df_marks_sorted.to_excel(writer, sheet_name='Sorted by Marks', index=False)

print("Combined output file generated successfully: output_combined.xlsx")