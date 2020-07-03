from selenium import webdriver
import requests
from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(ChromeDriverManager().install())


def download(url1: str, url2: str, dest_folder: str):
    r1 = requests.get(url1)
    r2 = requests.get(url2)
    if r1.ok:
        with open('/Users/charles/PycharmProjects/scraping_comp354_grades/grades', 'wb') as f:
            f.write(r1.content)

    if r2.ok:
        with open('/Users/charles/PycharmProjects/scraping_comp354_grades/grades', 'wb') as f:
            f.write(r2.content)

    else:
        print("Wrong url")


student_ID = [] //to get from excel file

for i in range(len(student_ID)):
    link1 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ID)+".pdf"
    link2 = "http://users.encs.concordia.ca/~kamthan/courses/comp-354/t2/" + str(student_ID)+".docx"
    download(link1, link2, student_ID[i])
    
