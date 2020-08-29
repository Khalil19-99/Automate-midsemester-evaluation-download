import requests
import xlsxwriter
domain= input("Enter the moodle domain: ")
token = input("Enter your token: ")
item = input("Enter your gradeitem: ")
if domain=="moodle.innopolis.university":
    domain="http://10.90.105.115/moodle"
workbook = xlsxwriter.Workbook(str(item)+".xlsx")
item=item.lower()

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'first name')
worksheet.write('B1', 'last name')
worksheet.write('C1', 'email')

for o in range(3):
    sitee = (
       '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_enrol_get_enrolled_users&wstoken=%s&courseid=2')%(domain,token)
    site=requests.get(sitee)
    worksheet.write('A' + str(o + 2), site.json()[o]["firstname"])
    worksheet.write('B' + str(o + 2), site.json()[o]["lastname"])
    worksheet.write('C' + str(o + 2), site.json()[o]["email"])
for j in range(2,4):
    grades = (
            '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grades_table&wstoken=%s&courseid=' + str(
        j))%(domain,token)

    itemid=requests.get(grades)
    for id in range(1,+1000):
        s = itemid.json()["tables"][0]["tabledata"][id]["itemname"]["content"]
        s=s.lower()
        if s.find(item) != -1:
            break
    sitee = (
                '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_course_get_courses&wstoken=%s') %(domain,token)
    site = requests.get(sitee)
    worksheet.write(chr(66+j)+'1', site.json()[j-1]["fullname"])
    sitee = (
       '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grade_items&wstoken=%s&courseid='+str(j))%(domain,token)
    site=requests.get(sitee)
    # Use the worksheet object to write
    # data via the write() method.
    for i in range(3):
        worksheet.write(chr(66+j) + str(i + 2), site.json()["usergrades"][i]["gradeitems"][id-1]["gradeformatted"])

    # Finally, close the Excel file
    # via the close() method.
workbook.close()
#admin's token:efcc57f3b9deca9c98686e583d089e09
#moodle.innopolis.university

