import pandas as pd

# Load the Excel file
file_path = 'ExcelData.xlsx'  # Use your local file path here
df = pd.read_excel(file_path, 'pythonF', skiprows=[0])

# Assigning proper column names
df.columns = ['ID', 'Grade', 'Age']

# Convert 'Grade' and 'Age' columns to numeric values
df['Grade'] = pd.to_numeric(df['Grade'], errors='coerce').fillna(0)
df['Age'] = pd.to_numeric(df['Age'], errors='coerce').fillna(0)

# Extract relevant data
IDs = df['ID'].tolist()
grades = df['Grade'].tolist()
ages = df['Age'].tolist()

# Calculate statistics
total_students = len(IDs)
average_grade = round(sum(grades) / total_students, 2)
highest_grade = max(grades)
students_with_highest_grade = [IDs[i] for i, grade in enumerate(grades) if grade == highest_grade]
oldest_student_age = max(ages)
average_age = round(sum(ages) / total_students, 2)
age_mode = df['Age'].mode()[0]

# Count students in different grade ranges
grade_ranges = {
    '0': len([grade for grade in grades if grade == 0]),
    '1-10': len([grade for grade in grades if 1 <= grade <= 10]),
    '11-20': len([grade for grade in grades if 11 <= grade <= 20]),
    '21-30': len([grade for grade in grades if 21 <= grade <= 30]),
    '31-40': len([grade for grade in grades if 31 <= grade <= 40])
}

# Display results
output = f"""
Count of Students in Different Ranges:
0: {grade_ranges['0']}
1-10: {grade_ranges['1-10']}
11-20: {grade_ranges['11-20']}
21-30: {grade_ranges['21-30']}
31-40: {grade_ranges['31-40']}

Total Number of Students: {total_students}
Average Grade: {average_grade}
Highest Grade: {highest_grade}
IDs with Highest Grade: {students_with_highest_grade}

The oldest student has the age of: {oldest_student_age}
The student age average is: {average_age}
The student age mode is: {age_mode}
"""

print(output)
