import csv
from trainercourses.course import CourseCollection
from openpyxl import Workbook,load_workbook

path = r'dev\Collection1.xlsx'
cc = CourseCollection.open_excel(path)
def test1():
    courses = list(cc.courses)

    for course in cc:
        print(course.dict())
def test2():
    cc = CourseCollection.open_excel(path)
    cc.build_library(no_save=True)

def test3():
    cc = CourseCollection.open_excel(path)
    cc.save(r'C:\Users\dvanb\source\repos\TrainerCourses\dev\output2')

test3()

