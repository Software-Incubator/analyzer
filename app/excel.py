import xlsxwriter

from xlrd import open_workbook
import string
from app import connection, app


def make_excel(branch, year='3', colg_code='027', output=None):
    year = str(year)
    print 'excel', colg_code, branch, year
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('basic.xlsx')
    worksheet = workbook.add_worksheet()
    collection = connection.test.students
    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    format = workbook.add_format()
    format.set_text_wrap()

    print 'Branch: ', branch, "College: ", colg_code, "year: ", year
    head = collection.find_one({'branch_code': branch,
                                        'college_code': colg_code,
                                        'year': year})
                                        # 'sem': sem}) semester is removed now
    print 'value of head', head
    if not head:
        worksheet.merge_range('A1:AO1', 'Result not declared')
        workbook.close()
        return False
    # print head
    # print head['roll_no']
    # col_mark will find the number of subjects

    worksheet.merge_range('A1:AO1', head['college_name'] +
                          " - "
                          + head['branch_name'] +
                          '- Year:  ' +
                          year, merge_format)

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:C', 30)
    worksheet.write(1, 0, "Roll No.", merge_format)
    worksheet.write(1, 1, "Name", merge_format)
    worksheet.write(1, 2, "Father's Name", merge_format)
    col = len(head['marks'])
    # print col
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
        j += 3
        col -= 1
    # for sum  total of marks in a row
    worksheet.merge_range(1,j,1, j+2, 'Total', merge_format)
    worksheet.set_column(j+3, j+3, 45) # for carry papers
    worksheet.write(1,j+3, 'Carry Papers', merge_format)  # for carry papers
    worksheet.write(2, j, 'External')
    worksheet.write(2, j+1, 'Internal')
    worksheet.write(2, j+2, 'Total')
    worksheet.merge_range(1, j+4, 2, j+4, "Status", merge_format)

    # for marks now
    row = 3
    j = 3
    for st in collection.find({'branch_code': branch,
                                       'college_code': colg_code,
                                       'year': year}):
        worksheet.write(row, 0, st['roll_no'], format)
        worksheet.write(row, 1, st['name'], format)
        worksheet.write(row, 2, st['father_name'], format)
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
        worksheet.write(row, a + 4, st['carry_status'])  # for status column

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
    workbook = xlsxwriter.Workbook('college_summary_excel_year_' + str(year) + '.xlsx')
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    format = workbook.add_format()
    format.set_text_wrap()
    collection = connection.test.students
    branch_codes = collection.distinct("branch_code")
    r = 0
    # for headding
    heading = str(app.config['COLLEGE_CODENAMES'][college_code]) + '  YEAR: ' + year + '  2015-16'
    worksheet.write(r, 0,heading ,merge_format)
    worksheet.merge_range("A1:I1", heading, merge_format)
    r += 1
    worksheet.write(r, 0, 'S. No', merge_format)
    worksheet.write(r, 1, 'Branch', merge_format)
    worksheet.write(r, 2, 'Total', merge_format)
    worksheet.write(r, 3, 'RND', merge_format)
    worksheet.write(r, 4, 'INCOMP', merge_format)
    worksheet.write(r, 5, 'RD', merge_format)
    worksheet.write(r, 6, 'CP', merge_format)
    worksheet.write(r, 7, 'Pass', merge_format)
    worksheet.write(r, 8, 'Pass%', merge_format)
    r += 1
    t_total = 0
    t_rnd = 0
    t_incomp = 0
    t_rd = 0
    t_cp = 0
    t_pass_count = 0

    for branch_code in branch_codes:
        all_stud = collection.find({'college_code': college_code,
                                    'branch_code': branch_code,
                                    'year': year})
        incomp_stud = collection.find({'college_code': college_code,
                                       'branch_code': branch_code,
                                       'year': year,
                                       'carry_status': 'INCOMP'})
        total = app.config["BRANCH_STUDENTS"].get(branch_code)
        if not total:
            continue
        incomp_count = incomp_stud.count()
        rd = all_stud.count() - incomp_count
        rnd = total - rd
        cp = collection.find({'college_code': college_code,
                              'branch_code': branch_code,
                              'year': year,
                              'carry_status': {'$nin': ['CP(0)', 'INCOMP']}
                              }).count()
        pass_count = rd - cp
        if rd != 0:
            pass_percent = (float(pass_count) / rd) * 100
            pass_percent = round(pass_percent, 2)
        else:
            pass_percent = '-'

        worksheet.write(r, 0, r-1, format)
        print branch_code
        worksheet.write(r, 1, app.config['BRANCH_CODENAMES'][branch_code],
                        format)
        worksheet.write(r, 2, total, format)
        worksheet.write(r, 3, rnd, format)
        worksheet.write(r, 4, incomp_count, format)
        worksheet.write(r, 5, rd, format)
        worksheet.write(r, 6, cp, format)
        worksheet.write(r, 7, pass_count, format)
        worksheet.write(r, 8, pass_percent, format)
        r += 1
        # for totals
        t_total += total
        t_rnd += rnd
        t_incomp += incomp_count
        t_rd = rd + t_rd
        t_cp = cp + t_cp
        t_pass_count += pass_count
    if t_rd != 0:
        t_pass_percent = float(t_pass_count) / t_rd * 100
        t_pass_percent = round(t_pass_percent, 2)
    else:
        t_pass_percent = '-'
    worksheet.write(r, 1, 'Total', merge_format)
    worksheet.write(r, 2, t_total , format)
    worksheet.write(r, 3, t_rnd, format)
    worksheet.write(r, 4, t_incomp, format)
    worksheet.write(r, 5, t_rd, format)
    worksheet.write(r, 6, t_cp, format)
    worksheet.write(r, 7, t_pass_count, format)
    worksheet.write(r, 8, t_pass_percent, format)
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
    for colg_code in college_codes:
        students = collection.find({'year': str(year), 'college_code': colg_code})
        t_ext = 0
        for student in students:
            ext = 0
            for mark in student['marks']:
                ext = ext + mark['marks'][0]
            t_ext = t_ext + ext
            print t_ext


def teachers_excel(year):
    year = str(year)
    workbook = xlsxwriter.Workbook('faculty_performance_year_' + year + '.xlsx')
    worksheet = workbook.add_worksheet('YEAR - ' + year)
    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 17)
    college_code = '027'
    collection = connection.test.students
    branch_codes = collection.find({'year': year, 'college_code': college_code}).distinct('branch_code')
    heading_format = workbook.add_format({
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
                if sub_code[1:3] == 'GP' or sub_code[:2] == 'GP':
                    continue
                if not sub_details.get(sub_code):
                    # print 'adding sub code', sub_code
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
            if branch_code == '00' or branch_code == '10':
                print "subject code: ", sub_code, ' type: ', type(sub_code)
            if len(sub_dict) > 1:
                worksheet.merge_range(r, c, r + len(sub_dict) - 1, c, sub_code, cell_format)
            else:
                worksheet.write(r, c, sub_code, cell_format)
            for section in sub_dict:
                external_avg = round(float(sub_dict[section][0]) / sub_dict[section][2], 2)
                total_avg = round(float(
                        sub_dict[section][0] + sub_dict[section][1]) / sub_dict[section][2], 2)
                pass_percent = round(float(
                        sub_dict[section][2] - sub_dict[section][3]) / sub_dict[section][2] * 100, 2)
                worksheet.write(r, c+1, ' ')
                worksheet.write(r, c+2, external_avg, cell_format)
                worksheet.write(r, c+3, total_avg, cell_format)
                worksheet.write(r, c+4, pass_percent, cell_format)
                worksheet.write(r, c+5, section, cell_format)
                r += 1
    workbook.close()

# for year in [1, 3, 4]:
#     teachers_excel(year=year)

# college_wise_excel('027', '1')



def subject_wise(year='1'):
    """
    subject wise comparison of marks of all 4 colleges
    :param year: year for which the analysis is done
    :return:
    """
    workbook = xlsxwriter.Workbook('subject_wise_year_' + year + '.xlsx')
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "valign"
    })
    cell_format = workbook.add_format({
        "align": "center"
    })
    r, c = 0, 0
    collection = connection.test.students
    college_codes = collection.distinct('college_code')
    college_codenames = app.config["COLLEGE_CODENAMES"]
    worksheet.merge_range(r, c, r+1, c, "Subject Code", merge_format)
    worksheet.merge_range(r, c+1, r+1, c+1, "Subject Name", merge_format)
    for college_code in college_codes:
        worksheet.merge_range(r, c+2, r, c+4,
                              college_codenames[college_code], merge_format)
        worksheet.write(r+1, c+2, "Total", merge_format)
        worksheet.write(r+1, c+3, "Pass %", merge_format)
        worksheet.write(r+1, c+4, "Fail %", merge_format)
        c += 3
    r += 2
    students = collection.find({"year": year,
                                "carry_status": {"$ne": "INCOMP"}})
    year_dict = dict()
    for student in students:
        college_code = student['college_code']
        carry_papers = student['carry_papers']
        for mark_dict in student['marks']:
            sub_code = mark_dict['sub_code']
            is_carry = sub_code in carry_papers
            if year_dict.get(sub_code):
                sub_dict = year_dict.get(sub_code)
                if sub_dict.get(college_code):
                    col_dict = sub_dict.get(college_code)
                    col_dict['total'] += 1
                    if is_carry:
                        col_dict['fail'] += 1
                else:
                    sub_dict[college_code] = {"total": 1}
                    if is_carry:
                        sub_dict[college_code]["fail"] = 1
                    else:
                        sub_dict[college_code]["fail"] = 0
            else:
                sub_dict = dict()
                sub_dict[college_code] = {"total": 1}
                if sub_code in carry_papers:
                    sub_dict[college_code]["fail"] = 1
                else:
                    sub_dict[college_code]["fail"] = 0
                year_dict[sub_code] = sub_dict

    # writing to the excel
    for sub_code in year_dict:
        c = 0
        worksheet.write(r, c, sub_code, cell_format)
        worksheet.write(r, c+1, sub_code, cell_format)
        c += 2
        sub_dict = year_dict[sub_code]
        for college_code in college_codes:
            college_dict = sub_dict.get(college_code)
            if college_dict:
                total = college_dict.get("total")
                fail = college_dict.get("fail")
                passed = total - fail
                pass_percent = round(float(passed)/total * 100, 2)
                fail_percent = round(float(fail)/total * 100, 2)
                worksheet.write(r, c, total)
                worksheet.write(r, c+1, pass_percent)
                worksheet.write(r, c+2, fail_percent)
            else:
                worksheet.write(r, c, '-')
                worksheet.write(r, c+1, '-')
                worksheet.write(r, c+2, '-')
            c += 3
        r += 1
    return True