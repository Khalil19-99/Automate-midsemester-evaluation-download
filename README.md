# Automate-midsemester-evaluation-download
download for all courses a column 'Mid Semester Evaluation' into a single file.
when you run this code, you will revieve a message "Enter the moodle domain:" here you have to put the moodle domain, which you use it in a browser to access moodle (the domain for innopolis moodle is written as a comment in the end of the code).
then this code asks you to input your token, in the end of the code there is the admin's token which you can use to access any grade iteme
if you want to make your own token, 

    Dashboard > Site administration > Plugins > Web services > Manage tokens
    then click on add token 
finally you have to to enter the grade item you need, for example: midsemester evaluation, after that you will see an Excel.xlsx file contain a table for that grade item (if you enter an item which is not exist in some courses, you will recieve a message that this item not found in those courses)
Note: if there is an error about "xlsxwriter" that means you don't have the package on your own laptop, but you can easly download it, this link can help you https://xlsxwriter.readthedocs.io/getting_started.html 
