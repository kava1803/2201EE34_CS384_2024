# 11th Part 1


!pip install xlsxwriter
import openpyxl
import pandas as pd
from datetime import datetime

# Define input and output file paths
input_file = '/content/Input_Lab11.xlsx'
output_file = '/content/output.xlsx'

# Create an empty Excel file for the output
wb = openpyxl.Workbook()
wb.save(output_file)

# Define the function to process the input file
def process_file(input_file, output_file):
    # Load the Excel file
    df = pd.read_excel(input_file)

    # Step 1: Ensure 'Total' column is numeric
    df['Total'] = pd.to_numeric(df['Total'], errors='coerce')

    # Drop rows with NaN in 'Total'
    df.dropna(subset=['Total'], inplace=True)

    # Sort the data by 'Total' column
    df = df.sort_values(by='Total', ascending=False).reset_index(drop=True)

    # Step 2: Create dictionaries to store min and max for each grade
    grade_min = {}
    grade_max = {}
    for _, row in df.iterrows():
        grade = row['Grade']
        total = row['Total']
        if grade not in grade_max:
            grade_max[grade] = total
        grade_min[grade] = total

    # Step 3: Define the predefined normalization ranges for each grade
    grade_ranges = {
        'AA': (100, 91), 'AB': (90, 81), 'BB': (80, 71),
        'BC': (70, 61), 'CC': (60, 51), 'CD': (50, 41),
        'DD': (40, 31), 'F': (30, 0), 'PP': (0, 0), 'NP': (0, 0)
    }

    # Normalize the 'Total' column using the grade-specific ranges
    normalized_totals = []
    for _, row in df.iterrows():
        grade = row['Grade']
        total = row['Total']
        if grade in grade_ranges:
            original_max = grade_max[grade]
            original_min = grade_min[grade]
            scale_max, scale_min = grade_ranges[grade]
            normalized = scale_min + ((scale_max - scale_min) * (total - original_min) / (original_max - original_min))
            normalized_totals.append(normalized)
        else:
            normalized_totals.append(total)
    df['Normalized'] = normalized_totals

    # Generate grade statistics
    grade_stats = df.groupby('Grade')['Total'].agg(['min', 'max', 'count']).reset_index()
    grade_stats.columns = ['Grade', 'Min (x)', 'Max (x)', 'Count']

    # Calculate IAPC counts and differences
    total_students = df['Total'].count()
    iapc_percentages = {'AA': 5, 'AB': 15, 'BB': 25, 'BC': 30, 'CC': 15, 'CD': 5, 'DD': 5, 'F': 0, 'PP': 0, 'NP': 0}
    iapc_data = []
    total_diff = 0
    total_count = 0
    for grade, percent in iapc_percentages.items():
        iapc_count = round((percent / 100) * total_students)
        actual_count = grade_stats.loc[grade_stats['Grade'] == grade, 'Count'].values[0] if grade in grade_stats['Grade'].values else 0
        diff = actual_count - iapc_count
        total_diff += diff
        total_count += actual_count
        iapc_data.append((grade, percent, iapc_count, diff))
    iapc_df = pd.DataFrame(iapc_data, columns=['Grade', 'IAPC (%)', 'IAPC-Count', 'Diff'])

    # Save the output to an Excel file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write student data
        df.to_excel(writer, sheet_name='Sheet1', startrow=13, index=False)

        # Write grade statistics and IAPC data side by side
        grade_stats.to_excel(writer, sheet_name='Sheet1', startrow=2, startcol=0, index=False)
        iapc_df.to_excel(writer, sheet_name='Sheet1', startrow=2, startcol=5, index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        workbook.formula_recalc = True

        # Add Subject Code and Date
        worksheet.write(0, 0, 'Subject Code')
        worksheet.write(0, 5, f'Date: {datetime.now().strftime("%B-%Y")}')

        # Add Total Student Count above the right section
        worksheet.write(1, 5, 'Total Students')
        worksheet.write(2, 5, total_students)

        # Add totals for Count and Diff columns
        worksheet.write(2 + len(grade_stats), 3, 'Total Count')
        worksheet.write(3 + len(grade_stats), 3, total_count)

        worksheet.write(2 + len(iapc_df), 8, 'Total Diff')
        worksheet.write(3 + len(iapc_df), 8, total_diff)

        # Adjust column widths and apply formatting
        worksheet.set_column('A:I', 15)

        # Add conditional formatting for the Diff column
        diff_start_row = 4
        diff_col = 9
        diff_end_row = 2 + len(iapc_df)

        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})  # Green background
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})   # Red background

        # Apply conditional formatting
        worksheet.conditional_format(
            f'I{diff_start_row}:I{diff_end_row}',
            {'type': 'cell', 'criteria': '<', 'value': 0, 'format': green_format}
        )
        worksheet.conditional_format(
            f'I{diff_start_row}:I{diff_end_row}',
            {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_format}
        )

# Process the Excel file
process_file(input_file, output_file)

# Step 9: Optionally download the file if needed
from google.colab import files
files.download(output_file)