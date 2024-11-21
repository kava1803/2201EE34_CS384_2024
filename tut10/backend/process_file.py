import sys
import os
import pandas as pd
import numpy as np

def process_file(input_path):
    df = pd.read_excel(input_path)

    max_marks = df.iloc[0, 2:].astype(float)
    weightage = df.iloc[1, 2:].astype(float) / 100

    df = df.drop([0, 1]).reset_index(drop=True)
    df.iloc[:, 2:] = df.iloc[:, 2:].apply(pd.to_numeric)

    df['Grand Total/100'] = sum(
        (df[col] / max_marks[col] * weightage[col] * 100) for col in max_marks.index
    )

    grade_percentages = {
        'AA': 5,
        'AB': 15,
        'BB': 25,
        'BC': 30,
        'CC': 15,
        'CD': 5,
        'DD': 5
    }

    total_students = len(df)
    grade_counts = {grade: int(np.round(percentage / 100 * total_students)) for grade, percentage in grade_percentages.items()}

    df_sorted_by_marks = df.sort_values(by='Grand Total/100', ascending=False).reset_index(drop=True)

    grades = []
    for grade, count in grade_counts.items():
        grades.extend([grade] * count)

    grades += ['F'] * (total_students - len(grades))
    df_sorted_by_marks['Grade'] = grades

    actual_counts = df_sorted_by_marks['Grade'].value_counts().reindex(grade_counts.keys(), fill_value=0)
    grade_df = pd.DataFrame({
        'grade': grade_counts.keys(),
        'iapc': grade_counts.values(),
        'IAPC-Count': actual_counts.values,
        'Diff': actual_counts.values - list(grade_counts.values())
    })

    df_sorted_by_marks['Total Students'] = total_students
    df_sorted_by_marks['Counts'] = actual_counts[df_sorted_by_marks['Grade']].values
    df_sorted_by_marks['Round'] = np.round(df_sorted_by_marks['Counts'])
    df_sorted_by_marks['Count verified'] = df_sorted_by_marks['Round'].astype(int)

    df_sorted_by_roll = df_sorted_by_marks.sort_values(by='Roll').reset_index(drop=True)

    output_directory = 'output'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    output_excel_file = os.path.join(output_directory, 'output-file.xlsx')
    with pd.ExcelWriter(output_excel_file) as writer:
        df_sorted_by_roll.to_excel(writer, sheet_name='Sorted by Roll Number', index=False)
        df_sorted_by_marks.to_excel(writer, sheet_name='Sorted by Marks', index=False)
        grade_df.to_excel(writer, sheet_name='Grade Distribution', index=False)

    print("Combined output file generated successfully:", output_excel_file)

if __name__ == "__main__":
    input_file_path = sys.argv[1]
    process_file(input_file_path)
