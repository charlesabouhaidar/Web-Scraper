from selenium import webdriver
import requests
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import xlrd
import math
import os
import re
import docx2txt
import numpy as np
import matplotlib.pyplot as plt
import shutil

driver = webdriver.Chrome(ChromeDriverManager().install())

file = pd.ExcelFile(r"PATH/TO/EXCEL/FILE/OF/PREVIOUS/TEST/CONTAINING/STUDENT/IDS")
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_name('COMP 354')
x = []

for rownum in range(sheet.nrows)[3:]:
    x.append(sheet.cell(rownum, 0))
student_ids = []
for i in range(len(x)):
    student_ids.append(int(x[i].value))


def download(url1: str, url2: str, student_id: int):
    r1 = requests.get(url1)
    r2 = requests.get(url2)
    if r1.ok:
        print(student_id, " found!")
        path1 = r"/PATH/TO/WORD/DIRECTORY"
        filename = url1.split('/')[-1].replace(" ", "_")  # be careful with file names
        file_path = os.path.join(path1, filename)
        with open(file_path, 'wb') as f:
            f.write(r1.content)

    elif r2.ok:
        print(student_id, " found!")
        path2 = r"/PATH/TO/WORD/DIRECTORY"
        filename = url2.split('/')[-1].replace(" ", "_")  # be careful with file names
        file_path = os.path.join(path2, filename)
        with open(file_path, 'wb') as f:
            f.write(r2.content)

    else:
        print("Wrong url")


def get_word_grades(file_name):
    string = docx2txt.process(file_name)
    string = string.split()
    for i in range(len(string)):
        if re.search(r"/100", string[i]):
            return string[i]

# delete all contents of directory for privacy reasons
def empty_directory(folder_path):
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


for i in range(len(student_ids)):
    link1 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".pdf"
    link2 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".docx"
    download(link1, link2, student_ids[i])

file_paths1 = []
for root, directories, files in os.walk(r"/PATH/TO/WORD/DIRECTORY"):
    for filename in files:
        filepath1 = os.path.join(root, filename)
        file_paths1.append(filepath1)

file_paths2 = []
for root, directories, files in os.walk(r"/PATH/TO/WORD/DIRECTORY"):
    for filename in files:
        filepath2 = os.path.join(root, filename)
        file_paths2.append(filepath2)

grades = []
for i in file_paths1:
    grades.append(get_word_grades(i))
sum1 = 0
final_grades = []
for grade in grades:
    if isinstance(grade, str):
        final_grades.append(int(grade[:2]))
        sum1 = sum1 + int(grade[:2])
    else:
        final_grades.append(0)

grades_sixties = []
grades_seventies = []
grade_eighties = []
grade_nineties = []
for i in final_grades:
    if 60 <= i < 70:
        grades_sixties.append(i)
    elif 70 <= i < 80:
        grades_seventies.append(i)
    elif 80 <= i < 90:
        grade_eighties.append(i)
    elif i >= 90:
        grade_nineties.append(i)

print("FOR DOCX FILES: ")
print("Number of doc files: ", len(file_paths1))
print("Highest grade: ", max(final_grades))
print("Lowest grade: ", min(final_grades))
print("Average grade: ", np.average(final_grades))
print("Median: ", np.median(final_grades))
print(">59 and <70: ", len(grades_sixties))
print(">69 and <80: ", len(grades_seventies))
print(">79 and <90: ", len(grade_eighties))
print(">89 and <100: ", len(grade_nineties))
print("None, were replaced by 0: ", final_grades.count(0))
bins = np.linspace(math.ceil(min(final_grades)), math.floor(max(final_grades)), 50)
plt.xlim([min(final_grades) - 5, max(final_grades) + 5])
plt.hist(final_grades, bins=bins, alpha=0.5)
plt.title('Test 2 grades')
plt.xlabel('Grades')
plt.ylabel('Count')
plt.xticks(np.arange(min(final_grades), max(final_grades) + 1, 4.0))
plt.show()

empty_directory(r"/PATH/TO/WORD/DIRECTORY")
empty_directory(r"/PATH/TO/PDF/DIRECTORY")
