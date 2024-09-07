import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'ExcelData.xlsx'
df = pd.read_excel(file_path, 'pythonF', skiprows=[0])

# Clean and prepare the data
df.columns = ['ID', 'Grade', 'Age']
df['Grade'] = pd.to_numeric(df['Grade'], errors='coerce').fillna(0)
df['Age'] = pd.to_numeric(df['Age'], errors='coerce').fillna(0)

# Extract data
IDs = df['ID'].tolist()
grades = df['Grade'].tolist()
ages = df['Age'].tolist()

# Calculate statistics
stdntTotal = len(IDs)
avgGrade = round(sum(grades) / stdntTotal, 2)
topGrade = max(grades)
topStdnt = [IDs[i] for i, grade in enumerate(grades) if grade == topGrade]
topAge = max(ages)
avgAge = round(sum(ages) / stdntTotal, 2)
ageMode = df['Age'].mode()[0]

# Grade ranges for pie chart
gradeRanges = {
    '0': len([grade for grade in grades if grade == 0]),
    '1-10': len([grade for grade in grades if 1 <= grade <= 10]),
    '11-20': len([grade for grade in grades if 11 <= grade <= 20]),
    '21-30': len([grade for grade in grades if 21 <= grade <= 30]),
    '31-40': len([grade for grade in grades if 31 <= grade <= 40])
}

# Create the pie chart
labels = ['0', '1-10', '11-20', '21-30', '31-40']
sizes = [gradeRanges['0'], gradeRanges['1-10'], gradeRanges['11-20'], gradeRanges['21-30'], gradeRanges['31-40']]
colors = ['yellow', 'lightgreen', 'salmon', 'skyblue', 'red']

plt.figure(figsize=(8,6))
plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
plt.title('Grades Pie Chart Distribution')
plt.savefig('pie_chart.png')
plt.show()

# Create the histogram
plt.figure(figsize=(12,5))
plt.hist(grades, bins=range(0, 41), color='skyblue', edgecolor='black', rwidth=0.8)
plt.title('Grade Distribution Histogram')
plt.xlabel('Grade ranges from 0 to 40')
plt.ylabel('Number of Students')
plt.savefig('histogram.png')
plt.show()

# Prepare the output
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

# Print the output to the console
print(output)

# Save the output to a text file
with open('output_results.txt', 'w') as f:
    f.write(output)
