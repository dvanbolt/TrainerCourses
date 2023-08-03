from trainercourses.course import CourseCollection

path = r'dev\Collection1.xlsx'
cc = CourseCollection.open_excel(path)
print(cc.summary(stats=True))