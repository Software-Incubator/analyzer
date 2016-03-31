import xlsxwriter
import string


from app import connection, app


def make_excel(branch, sem, colg_code='027', output=None):
    print 'excel', colg_code, branch, sem
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    collection = connection.test.students
    merge_format = workbook.add_format({
        'bold': True,
        #'border': 1,
        'align': 'center',
        'valign': 'vcenter',
    })
    format = workbook.add_format()
    format.set_text_wrap()

    i = 0    # i is the index of which list we have to iterate in marks

    if int(sem) % 2 == 0:
        # sem = str(int(sem) - 1)
        i = 1


    year = str((int(sem) - .1) // 2 + 1)[0]
    head = collection.Student.find_one({'branch_code': branch,
                                        'college_code': colg_code,
                                        'year': year})
                                        # 'sem': sem}) semester is removed now
    print 'value of head', head
    if not head:
        worksheet.merge_range('A1:AO1', 'Result not declared')
        workbook.close()
        return workbook
    # print head
    # print head['roll_no']
    # col_mark will find the number of subjects

    worksheet.merge_range('A1:AO1', head['college_name'] +
                          " - "
                          + head['branch_name'] +
                          '- Semester:  ' +
                          sem, merge_format)

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:C', 30)
    worksheet.write(1, 0, "Roll No.", merge_format)
    worksheet.write(1, 1, "Name", merge_format)
    worksheet.write(1, 2, "Father's Name", merge_format)
    col = len(head['marks']) - 1
    print col
    j = 3
    # for sub code and internal, external and total
    for x in head['marks'][:-1]:
        if x['sub_code'][1:4] == 'OE0':
            worksheet.merge_range(1,j,1,j+2, "OE0", merge_format)
        else:
            worksheet.merge_range(1,j,1,j+2, x['sub_code'], merge_format)
        worksheet.write(2, j, 'External')
        worksheet.write(2, j+1, 'Internal')
        worksheet.write(2, j+2, 'Total')
        worksheet.write(2, j+3, 'Carry Papers')
        j += 3
        col -= 1
    # for sum  total of marks in a row
    worksheet.merge_range(1,j,1, j+2, 'Total', merge_format)
    worksheet.set_column(j+3, j+3, 45) # for carry papers
    worksheet.write(1,j+3, 'Carry Papers', merge_format )  # for carry papers
    worksheet.write(2, j, 'External')
    worksheet.write(2, j+1, 'Internal')
    worksheet.write(2, j+2, 'Total')

    # for marks now
    row = 3
    j = 3
    for st in collection.Student.find({'branch_code': branch,
                                       'college_code': colg_code,
                                       'year': year}):
        worksheet.write(row, col, st['roll_no'])
        worksheet.write(row, col + 1, st['name'], format)
        worksheet.write(row, col + 2, st['father_name'], format)
        a = 3
        ext = 0 # for total external marks
        internal = 0 # for total internal marks
        for mark in st['marks'][:-1]:
            worksheet.write(row, a, mark['marks'][0])
            worksheet.write(row, a + 1, mark['marks'][1])
            worksheet.write(row, a + 2, sum(map(int, mark['marks'][:])))
            ext += int(mark['marks'][0])
            internal += int(mark['marks'][1])
            a += 3
        worksheet.write(row, a, ext)
        worksheet.write(row, a + 1, internal)
        worksheet.write(row, a + 2, ext + internal)
        # cp = ''   # cp is carry papers
        # for cps in st['carry_papers'][:]:
        #     cp = cp + str(cps) + ', '
        cp = ', '.join(st['carry_papers'])
        worksheet.write(row, a + 3, cp)

        row += 1

    workbook.close()
    return workbook


def fail_excel(college_code='027', year="1", output=None):
    """
    generates excel for failed students
    :param college_code: code of the college of which the excel is to be made
    :param year: year of students of which excel to be made
    :param output: for download of the excel
    :return: none
    """
    year = str(year)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook("failed_students.xlsx")
    collection = connection.test.students
    branch_codes = collection.distinct("branch_code")
    heading_format = workbook.add_format({"bold": True,
                                          "align": "center",
                                          "valign": "center"})
    print branch_codes
    for branch_code in branch_codes:
        worksheet = workbook.add_worksheet(
            app.config["BRANCH_CODENAMES"][branch_code])
        student = collection.find_one({"college_code": college_code,
                                       "year": str(year),
                                       "branch_code": branch_code})
        if not student:
            continue
        print "student: ", student
        worksheet.merge_range("A1:O1", student['branch_name'], heading_format)
        worksheet.write("A2", "S. No.", heading_format)
        worksheet.write("B2", "Name", heading_format)
        worksheet.write("C2", "Roll. No.", heading_format)
        cell_list = string.ascii_uppercase[3:]
        i = 0
        for sub_dict in student['marks']:
            sub_code = sub_dict['sub_code']
            if sub_code[1:4] == 'OE0':
                sub_code = 'OE0'
            worksheet.write(cell_list[i] + "2", sub_code, heading_format)
            i += 1
        worksheet.write(cell_list[i] + "2", "No. of Backs", heading_format)
        cp_students = collection.find({"college_code": college_code,
                                       "year": year,
                                       "branch_code": branch_code,
                                       "carry_status": {"$ne": "CP(0)"}
                                       })
        j = 2
        for fail_student in cp_students:
            k = 0
            worksheet.write(j, k, str(j-1))
            worksheet.write(j, k+1, fail_student['name'])
            worksheet.write(j, k+2, fail_student['roll_no'])
            k += 3
            carry_papers = fail_student['carry_papers']
            for mark_dict in fail_student['marks']:
                if mark_dict['sub_code'] in carry_papers:
                    worksheet.write(j, k, "F")
                else:
                    worksheet.write(j, k, "-")
                k += 1
            worksheet.write(j, k, str(len(carry_papers)))
            j += 1
    workbook.close()
    return True


def faculty_excel(year):
    collection = connection.test.students
    students = collection.find({'year': year})
    subject_codes = []
    for marks_dict in collection.findOne({"year": year}):
        pass

# fail_excel()