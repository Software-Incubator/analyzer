import os
import xlsxwriter
from app.models import connection
from xlrd import open_workbook


def open_excel(years=range(1, 4)):
    """
    this function collect all the students roll_no and section in a list of tuple
        student_data = [(roll_no,section)]
    this particular list is send to the update_section function will update
    the section detail of the student
    :param years: iterable, years for which the sections are to be updated
    :return: True
    """
    for year in years:
        year = str(year)
        col = 0
        wb = open_workbook(os.getcwd() + "/Roll Number lists/" +
                           year + "_year.xlsx")
        student_data = []
        # for s in wb.sheets()[: -1]:
        for s in wb.sheets():
            print 'updating sections for ', s.name
            for row in range(s.nrows):
                data = (s.cell(row, col).value, s.cell(row, col + 1).value)
                student_data.append(data)
        update_section(student_data)
    return True


def update_section(student_data):
    """
    make a query to get the document of the student via roll_no
    if the corresponding document exist then the section field of
    that document is update with the parameter passed while calling the function
    :param student_data:
    :return:
    """
    collection = connection.test.students
    if collection:
        for roll_num, section in student_data:
            # unicode_roll = "u'{}".format(roll_num)
            roll_num = int(roll_num)
            student = collection.find_one(
                {'roll_no': str(roll_num).encode("utf-8").decode("utf-8")})
            if student:
                student['section'] = str(section).encode("utf-8").decode(
                    "utf-8")
                print("Updated")
                print student['section']
                print student['roll_no']
                if student['year'] == '3':
                    student['aggregate_marks'] = u''
                updated_student = collection.Student(student)
                updated_student.save()
            else:
                print(False)
        else:
            print "Updated section of {} students".format(len(student_data))


def generate_list():
    collection  = connection.test.students
    college_students = collection.find({"college_code": "027"})
    workbook = xlsxwriter.Workbook('roll_num_list.xlsx')
    worksheet = workbook.add_worksheet()
    heading_format = workbook.add_format({"bold": True, "align": "center"})
    worksheet.set_column("A:A", 20)
    r, c = 0, 0
    worksheet.write(r, c, "Roll No.", heading_format)
    worksheet.write(r, c+1, "Section", heading_format)
    worksheet.write(r, c+2, "Year", heading_format)
    r += 1
    for student in college_students:
        worksheet.write(r, c, student['roll_no'])
        worksheet.write(r, c+1, student.get('section', 'not provided'))
        worksheet.write(r, c+2, student['year'])
        r += 1
    workbook.close()
    return True

