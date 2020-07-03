from selenium import webdriver
import requests
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import xlrd
import os
import re
import docx2txt
import numpy as np


driver = webdriver.Chrome(ChromeDriverManager().install())

file = pd.ExcelFile(f"/Users/charles/Downloads/t1_marks.xlsx")
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_name('COMP 354')
x = []

for rownum in range(sheet.nrows)[3:]:
    x.append(sheet.cell(rownum, 0))
student_ids = []
for i in range(len(x)):
    student_ids.append(int(x[i].value))

print(student_ids)


def download(url1: str, url2: str, student_id: int):
    r1 = requests.get(url1)
    r2 = requests.get(url2)
    if r1.ok:
        print(student_id, " found!")
        path1 = r"/Users/charles/PycharmProjects/scraping_comp354_grades/grades/pdf"
        filename = url1.split('/')[-1].replace(" ", "_")  # be careful with file names
        file_path = os.path.join(path1, filename)
        with open(file_path, 'wb') as f:
            f.write(r1.content)

    elif r2.ok:
        print(student_id, " found!")
        path2 = r"/Users/charles/PycharmProjects/scraping_comp354_grades/grades/word"
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
        if re.search(r"/96", string[i]):
            return string[i]

file_paths1 = []
for root, directories, files in os.walk("/Users/charles/PycharmProjects/scraping_comp354_grades/grades/word"):
    for filename in files:
        filepath1 = os.path.join(root, filename)
        file_paths1.append(filepath1)


file_paths2 = []
for root, directories, files in os.walk("/Users/charles/PycharmProjects/scraping_comp354_grades/grades/pdf"):
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



print("FOR DOCX FILES: ")
print("Highest grade: ", max(final_grades))
print("Lowest grade: ", min(final_grades))
print("Average grade: ", np.average(final_grades))


for i in range(len(student_ids)):
    link1 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".pdf"
    link2 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".docx"
    download(link1, link2, student_ids[i])

