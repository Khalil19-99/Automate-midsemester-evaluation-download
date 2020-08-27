import requests
import xlsxwriter

workbook = xlsxwriter.Workbook('midsemester evaluation.xlsx')

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()
token = input("Enter your token: ")
worksheet.write('A1', 'first name')
worksheet.write('B1', 'last name')
worksheet.write('C1', 'email')

for o in range(3):
    sitee = (
       'http://10.90.105.115/moodle/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_enrol_get_enrolled_users&wstoken=%s&courseid=2')%token
    site=requests.get(sitee)
    worksheet.write('A' + str(o + 2), site.json()[o]["firstname"])
    worksheet.write('B' + str(o + 2), site.json()[o]["lastname"])
    worksheet.write('C' + str(o + 2), site.json()[o]["email"])
for j in range(1,3):
    sitee = (
                'http://10.90.105.115/moodle/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_course_get_courses&wstoken=%s') % token
    site = requests.get(sitee)
    worksheet.write(chr(67+j)+'1', site.json()[j]["fullname"])
    sitee = (
       'http://10.90.105.115/moodle/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grade_items&wstoken=%s&courseid=2')%token
    site=requests.get(sitee)
    # Use the worksheet object to write
    # data via the write() method.
    for i in range(3):
        worksheet.write(chr(67+j) + str(i + 2), site.json()["usergrades"][i]["gradeitems"][3]["gradeformatted"])

    # Finally, close the Excel file
    # via the close() method.
workbook.close()
#admin's token:efcc57f3b9deca9c98686e583d089e09
