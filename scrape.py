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


driver = webdriver.Chrome(ChromeDriverManager().install())

# open excel file that has the student IDs from a previous test...
file = pd.ExcelFile(f"/Users/harcles/Downloads/t1_marks.xlsx")
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_name('COMP 354')
x = []

# extract student id cell from the first row starting with the first student ID at index 3 
for rownum in range(sheet.nrows)[3:]:
    x.append(sheet.cell(rownum, 0))

# convert cell to its value (int for student ID)
student_ids = []
for i in range(len(x)):
    student_ids.append(int(x[i].value))

# download file into appropriate folder
def download(url1: str, url2: str, student_id: int):
    
    r1 = requests.get(url1)
    r2 = requests.get(url2)
    
    if r1.ok:
        print(student_id, " found!")
        
        path1 = r"/Users/harcles/PycharmProjects/scraping_comp354_grades/grades/pdf"
        filename = url1.split('/')[-1].replace(" ", "_")  # be careful with file names
        file_path = os.path.join(path1, filename)
        
        with open(file_path, 'wb') as f:
            f.write(r1.content)

    elif r2.ok:
        print(student_id, " found!")
        
        path2 = r"/Users/harcles/PycharmProjects/scraping_comp354_grades/grades/word"
        filename = url2.split('/')[-1].replace(" ", "_")  # be careful with file names
        file_path = os.path.join(path2, filename)
        
        with open(file_path, 'wb') as f:
            f.write(r2.content)

    else:
        print("Wrong url")

# get grade from a file name
def get_word_grades(file_name):
    string = docx2txt.process(file_name)
    string = string.split()
    
    for i in range(len(string)):
        if re.search(r"/96", string[i]):
            return string[i]

def get_file_names_from_directory(file_name):
    list_of_files = []
    for root, directories, files in os.walk(file_name):
        for filename in files:
            filepath = os.path.join(root, filename)
            list_of_files.append(filepath)
            
     return list_of_files


# get the names of all the word files that were downloaded 
word_files = get_file_names_from_directory("/Users/harcles/PycharmProjects/scraping_comp354_grades/grades/word")
                                          
# get the names of all the PDF files that were downloaded   
pdf_files = get_file_names_from_directory("/Users/harcles/PycharmProjects/scraping_comp354_grades/grades/pdf")

# get all grades from DOCX files
grades = []
for i in word_files:
    grades.append(get_word_grades(i))

    
# need to find a way to read PDF files to get grade as well...


# update grades since some of them were None
final_grades = []
for grade in grades:
    if isinstance(grade, str): # checks if grades are not None, some students didn't have a grade so I replaced it with a 0
        final_grades.append(int(grade[:2]))
    else:
        final_grades.append(0)

# grade categories
grades_sixties = [] # grades in the 60s range
grades_seventies = [] # grades in the 70s range
grade_eighties = [] # grades in the 80s range
grade_nineties = [] # grades in the 90s range

# separate the grades into their categories to have more info about how the class did
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
# number of files scrapped
print("Number of doc files: ", len(word_files))

# highest grade in the list
print("Highest grade: ", max(final_grades))

# lowest grade in the list
print("Lowest grade: ", min(final_grades))

# average grade in the list
print("Average grade: ", np.average(final_grades))

# median in the list
print("Median: ", np.median(final_grades))

# categories
print(">59 and <70: ", len(grades_sixties))
print(">69 and <80: ", len(grades_seventies))
print(">79 and <90: ", len(grade_eighties))
print(">89 and <96: ", len(grade_nineties))
print("None, were replaced by 0: ", final_grades.count(0))

# plot
bins = np.linspace(math.ceil(min(final_grades)), math.floor(max(final_grades)), 50)
plt.xlim([min(final_grades) - 5, max(final_grades) + 5])
plt.hist(final_grades, bins=bins, alpha=0.5)
plt.title('Test 2 grades')
plt.xlabel('Grades')
plt.ylabel('Count')
plt.xticks(np.arange(min(final_grades), max(final_grades)+1, 4.0))

plt.show()

# web scraping for student id files, didn't know how to read PDF files... yet
for i in range(len(student_ids)):
    link1 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".pdf"
    link2 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ids[i]) + ".docx"
    download(link1, link2, student_ids[i])
