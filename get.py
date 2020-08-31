import requests
import xlsxwriter
import json
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
sitee = (
        '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_enrol_get_enrolled_users&wstoken=%s&courseid=2') % ( domain, token)
o=0
site = (requests.get(sitee)).json()
for student in site:
    worksheet.write('A' + str(o + 2), student["firstname"])
    worksheet.write('B' + str(o + 2), student["lastname"])
    worksheet.write('C' + str(o + 2), student["email"])
    o+=1
url = ('%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_course_get_courses&wstoken=%s')%(domain,token)
js=(requests.get(url)).json()
j=2
courseorder=-1
for ids in js:
    courseorder+=1
    itemisexisting=False
    if(ids["startdate"])==0:
        continue
    grades = (
            '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grades_table&wstoken=%s&courseid=' + str(
        ids["id"]))%(domain,token)

    itemid=(requests.get(grades)).json()
    id=1
    for table in itemid["tables"]:
        s = table["tabledata"][id]["itemname"]["content"]
        s=s.lower()
        if s.find(item) != -1:
            itemisexisting = True
            break
        id+=1
    sitee = (
                '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=core_course_get_courses&wstoken=%s') %(domain,token)
    site = requests.get(sitee)
    if not itemisexisting:
        print("item not found in "+site.json()[courseorder]["fullname"])
        continue
    worksheet.write(chr(66+j)+'1', site.json()[courseorder]["fullname"])
    sitee = (
       '%s/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grade_items&wstoken=%s&courseid='+str(
        ids["id"]))%(domain,token)
    site=requests.get(sitee)
    # Use the worksheet object to write
    # data via the write() method.
    for i in range(o):
         worksheet.write(chr(66+j) + str(i + 2), site.json()["usergrades"][i]["gradeitems"][id-1]["gradeformatted"])
    j+=1
    # Finally, close the Excel file
    # via the close() method.
workbook.close()
#admin's token:efcc57f3b9deca9c98686e583d089e09
#moodle.innopolis.university

