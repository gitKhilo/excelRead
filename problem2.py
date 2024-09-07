import pandas as pd

file_path = 'ExcelData.xlsx'
df = pd.read_excel(file_path, 'pythonF', skiprows=[0])

df.columns = ['ID', 'Grade', 'Age']

df['Grade'] = pd.to_numeric(df['Grade'], errors='coerce').fillna(0)
df['Age'] = pd.to_numeric(df['Age'], errors='coerce').fillna(0)

IDs = df['ID'].tolist()
grades = df['Grade'].tolist()
ages = df['Age'].tolist()

stdntTotal = len(IDs)
avgGrade = round(sum(grades) / stdntTotal, 2)
topGrade = max(grades)
topStdnt = [IDs[i] for i, grade in enumerate(grades) if grade == topGrade]
topAge = max(ages)
avgAge = round(sum(ages) / stdntTotal, 2)
ageMode = df['Age'].mode()[0]

gradeRanges = {
    '0': len([grade for grade in grades if grade == 0]),
    '1-10': len([grade for grade in grades if 1 <= grade <= 10]),
    '11-20': len([grade for grade in grades if 11 <= grade <= 20]),
    '21-30': len([grade for grade in grades if 21 <= grade <= 30]),
    '31-40': len([grade for grade in grades if 31 <= grade <= 40])
}

output = f"""
Count of Students in Different Ranges:
0: {gradeRanges['0']}
1-10: {gradeRanges['1-10']}
11-20: {gradeRanges['11-20']}
21-30: {gradeRanges['21-30']}
31-40: {gradeRanges['31-40']}

Total Number of Students: {stdntTotal}
Average Grade: {avgGrade}
Highest Grade: {topGrade}
IDs with Highest Grade: {topStdnt}

The oldest student has the age of: {topAge}
The student age average is: {avgAge}
The student age mode is: {ageMode}
"""

print(output)
