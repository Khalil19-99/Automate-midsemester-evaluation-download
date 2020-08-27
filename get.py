import requests
import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()
for j in range(2,4):
    site = requests.get(
        "http://10.90.105.115/moodle/webservice/rest/server.php?moodlewsrestformat=json&wsfunction=gradereport_user_get_grade_items&wstoken=efcc57f3b9deca9c98686e583d089e09&courseid="+str(j))

    # Use the worksheet object to write
    # data via the write() method.
    worksheet.write('A1', 'name')
    worksheet.write(chr(64+j)+'1', 'courseid='+str(j))
    for i in range(3):
        a = 'A' + str(i + 2)
        b = chr(64+j) + str(i + 2)

        worksheet.write(a, site.json()["usergrades"][i]["userfullname"] + " ")
        worksheet.write(b, site.json()["usergrades"][i]["gradeitems"][3]["gradeformatted"])

    # Finally, close the Excel file
    # via the close() method.
workbook.close()
