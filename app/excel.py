import xlsxwriter
import math
from xlrd import open_workbook
import string
from app import connection, app

'''
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
    for x in head['marks']:
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
        for mark in st['marks']:
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
    heading_format = workbook.add_format({'bold': True,
                                            #'border': 1,
                                            'align': 'center',
                                            'valign': 'vcenter',})
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
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 15)
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

def college_wise_excel(college_code, year):
    workbook = xlsxwriter.Workbook('college_wise_excel_' + str(year) + '_year.xlsx')
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        'bold': True,
        #'border': 1,
        'align': 'center',
        'valign': 'vcenter',
    })
    format = workbook.add_format()
    format.set_text_wrap()
    collection = connection.test.students
    branch_codes = collection.distinct("branch_code")
    code_names = app.config['BRANCH_CODENAMES']
    r, c = 0,0
    # for headding
    heading = str(app.config['COLLEGE_CODENAMES'][college_code]) + '  YEAR: ' + year + '  2015-16'
    worksheet.write(r, c,heading ,merge_format)
    worksheet.merge_range("A1:H1",heading ,merge_format)
    r += 1
    worksheet.write(r, c, 'S. No', merge_format)
    worksheet.write(r, c+1, 'Branch', merge_format)
    worksheet.write(r, c+2, 'Total', merge_format)
    worksheet.write(r, c+3, 'RND', merge_format)
    worksheet.write(r, c+4, 'RD', merge_format)
    worksheet.write(r, c+5, 'PCP', merge_format)
    worksheet.write(r, c+6, 'Pass', merge_format)
    worksheet.write(r, c+7, 'Pass%', merge_format)
    r += 1
    t_total = 0
    t_rnd = 0
    t_rd = 0
    t_pcp = 0
    t_pass_count = 0

    for branch_code in branch_codes:
        a_stud = collection.find({'college_code':college_code, 'branch_code': branch_code, 'year': year})
        total = app.config["BRANCH_STUDENTS"][branch_code]
        rd = a_stud.count()
        rnd = total - rd
        pcp = collection.find({'college_code':college_code, 'branch_code': branch_code, 'year': year,
                               "carry_status": {"$ne": "CP(0)" }}).count()
        pass_count = rd - pcp
        if rd != 0:
            pass_percent = (float(pass_count) / rd) * 100
        else:
            pass_percent = 0

        worksheet.write(r, c, r-1, format)
        print branch_code
        worksheet.write(r, c+1, app.config['BRANCH_CODENAMES'][branch_code], format)
        worksheet.write(r, c+2, total, format)
        worksheet.write(r, c+3, rnd, format)
        worksheet.write(r, c+4, rd, format)
        worksheet.write(r, c+5, pcp, format)
        worksheet.write(r, c+6, pass_count, format)
        worksheet.write(r, c+7, pass_percent, format)
        r +=1
        # for totals
        t_total = total + t_total
        t_pass_count = pass_count + t_pass_count
        t_rd = rd + t_rd
        t_pcp = pcp + t_pcp
        t_rnd = rnd + t_rnd
    if t_rd != 0:
        t_pass_percent = float(t_pass_count) / t_rd * 100
    else:
        t_pass_percent = 0
    worksheet.write(r, 1, 'Total', merge_format )
    worksheet.write(r, 2 , t_total , format)
    worksheet.write(r, 3, t_rnd, format)
    worksheet.write(r, 4, t_rd, format)
    worksheet.write(r, 5, t_pcp, format)
    worksheet.write(r, 6, t_pass_count, format)
    worksheet.write(r, 7, t_pass_percent, format)
    workbook.close()


def other_college_summary(college_code, year):
    workbook = xlsxwriter.Workbook("result_summary_" + str(year) + "_year.xlsx")
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    collection = connection.test.students
    r, c = 0,0
    # for headding
    heading = str(app.config['COLLEGE_CODENAMES'][college_code]) + '  YEAR: ' + year + '  2015-16'
    worksheet.merge_range("A1:H1",heading ,merge_format)
    r += 1
    worksheet.write(r, c, 'S. No', merge_format)
    worksheet.write(r, c+1, 'Branch', merge_format)
    worksheet.write(r, c+2, 'RD', merge_format)
    worksheet.write(r, c+3, 'PCP', merge_format)
    worksheet.write(r, c+4, 'Pass', merge_format)
    worksheet.write(r, c+5, 'Pass%', merge_format)
    r += 1
    t_rd = 0
    t_pcp = 0
    t_pass_count = 0

    branch_codes = collection.distinct('branch_code')
    for branch_code in branch_codes:
        rd = collection.find({"college_code": college_code, "year": year, "branch_code": branch_code}).count()
        pcp = collection.find({"college_code": college_code,
                               "year": year,
                               "branch_code": branch_code,
                               "carry_status": {"$ne": "CP(0)"}
                               }).count()
        pass_count = rd - pcp
        pass_percent = (float(pass_count) / rd) * 100

        worksheet.write(r, c, r-1, format)
        print branch_code
        worksheet.write(r, c+1, app.config['BRANCH_CODENAMES'][branch_code], format)
        worksheet.write(r, c+2, rd, format)
        worksheet.write(r, c+3, pcp, format)
        worksheet.write(r, c+4, pass_count, format)
        worksheet.write(r, c+5, pass_percent, format)
        r +=1
        # for totals
        t_pass_count = pass_count + t_pass_count
        t_rd = rd + t_rd
        t_pcp = pcp + t_pcp
    t_pass_percent = (float(t_pass_count) / t_rd) * 100
    worksheet.write(r, c+1, 'Total', merge_format )
    worksheet.write(r, c+2, t_rd, format)
    worksheet.write(r, c+3, t_pcp, format)
    worksheet.write(r, c+4, t_pass_count, format)
    worksheet.write(r, c+5, t_pass_percent, format)
    workbook.close()


def ext_avg(year):
    year = str(year)
    collection = connection.test.students
    college_codes = app.config['COLLEGE_CODES']
    avg_list = []
    for colg_code in college_codes:
        students = collection.find({'year': year, 'college_code': colg_code})
        t_ext = 0
        for student in students:
            ext = 0
            for mark in student['marks']:
                ext = ext + mark['marks'][0]
            t_ext = t_ext + ext

        len_students = collection.find({'year': year, 'college_code': colg_code}).count()
        #print 'len_students..' + str(len_students)
        colg_avg = float(t_ext) / len_students
        colg_avg = math.floor(colg_avg)
        #print 'colg_avg' + str(colg_avg) + 'is for..' +str(colg_code)
        max_mark = app.config['MAX_MARKS_YEARWISE'][year]
        percent = colg_avg / max_mark * 100
        percent = math.floor(percent)
        avg_dict = {colg_code: [colg_avg, percent]}
        avg_list.append(avg_dict)

    # now making excel
    workbook = xlsxwriter.Workbook('ext_avg of year=' + year + '.xlsx')
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        'bold': True,
        #'border': 1,
        'align': 'center',
        'valign': 'vcenter',
    })

    cell_format = workbook.add_format({
        'bold': False,
        #'border': 1,
        'align': 'center',
        'valign': 'vcenter',
    })

    worksheet.merge_range('A1:F1', 'External Exam Average Marks' ,merge_format)
    worksheet.merge_range('A2:F2', 'Maximum External Marks: ' + str(max_mark ), merge_format)
    worksheet.write(2, 0, 'College', merge_format)
    worksheet.write(3, 0, 'Average Marks', merge_format)
    worksheet.write(4, 0, 'Percentage %', merge_format)
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 15)

    r = 2
    c = 1
    for colg_dict in avg_list:
        colg_code = colg_dict.keys()[0]
        worksheet.write(r, c, app.config['COLLEGE_CODENAMES'][colg_dict.keys()[0]], merge_format)
        worksheet.write(r + 1, c, colg_dict[colg_code][0], cell_format)
        worksheet.write(r+ 2, c, colg_dict[colg_code][1], cell_format)
        c = c+1

    print avg_list

    workbook.close()

    '''

'''
def teachers_excel(year):
    year = str(year)d({'year': year, 'college_code': college_code}).distinct('branch_code')
    heading_format = workbook.add_f
    workbook = xlsxwriter.Workbook('faculty_performance_year_' + year + '.xlsx')
    worksheet = workbook.add_worksheet('YEAR - ' + year)
    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 17)
    college_code = '027'
    collection = connection.test.students
    branch_codes = collection.finormat({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter'
    })
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })
    r, c = 0, 0
    worksheet.write(r, c, 'Subject Name', heading_format)
    worksheet.write(r, c+1, 'Name Of Faculty', heading_format)
    worksheet.write(r, c+2, 'External Avg', heading_format)
    worksheet.write(r, c+3, 'Avg', heading_format)
    worksheet.write(r, c+4, 'Pass %', heading_format)
    worksheet.write(r, c+5, 'Section', heading_format)
    r += 1
    for branch_code in branch_codes:
        worksheet.merge_range(r, c, r, c+5, app.config["BRANCH_NAMES"][branch_code], heading_format)
        r += 1
        sub_details = dict()
        branch_students = collection.find({'year': year, 'college_code': college_code, 'branch_code':branch_code})
        for student in branch_students:
            carry_papers = student['carry_papers']
            for mark_dict in student['marks']:
                sub_code = mark_dict['sub_code']
                if sub_code[1:3] == 'GP':
                    continue
                if not sub_details.get(sub_code):
                    print 'adding sub code', sub_code
                    sub_details[sub_code] = dict()
                    sub_details[sub_code][student['section']] = [0] * 4
                else:
                    if not sub_details[sub_code].get(student['section']):
                        sub_details[sub_code][student['section']] = [0] * 4
                sub_details[sub_code][student['section']][0] += int(mark_dict['marks'][0])
                sub_details[sub_code][student['section']][1] += int(mark_dict['marks'][1])
                sub_details[sub_code][student['section']][2] += 1
                if sub_code in carry_papers:
                    sub_details[sub_code][student['section']][3] += 1

        for sub_code in sub_details:
            sub_dict = sub_details[sub_code]
            worksheet.merge_range(r, c, r + len(sub_dict) - 1, c, sub_code, cell_format)
            for section in sub_dict:
                external_avg = round(float(sub_dict[section][0]) / sub_dict[section][2], 2)
                total_avg = round(float(
                        sub_dict[section][0] + sub_dict[section][1]) / sub_dict[section][2], 2)
                pass_percent = round(float(
                        sub_dict[section][2] - sub_dict[section][3]) / sub_dict[section][2] * 100, 2)
                worksheet.write(r, c+1, 'Name Of Faculty')
                worksheet.write(r, c+2, external_avg, cell_format)
                worksheet.write(r, c+3, total_avg, cell_format)
                worksheet.write(r, c+4, pass_percent, cell_format)
                worksheet.write(r, c+5, section, cell_format)
                r += 1

    workbook.close()


'''