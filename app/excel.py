import xlsxwriter

from app import connection, app
from xlrd import open_workbook


def make_excel(branch, year='3', colg_code='027', output=None):
    year = str(year)
    print 'excel', colg_code, branch, year
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook(
            'main_year_' + year + ' - ' + colg_code + ' - ' + branch + '.xlsx')
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
        workbook = xlsxwriter.Workbook('failed_year_' + year + '.xlsx')
    collection = connection.test.students
    branch_codes = collection.distinct("branch_code")
    heading_format = workbook.add_format({'bold': True,
                                            'align': 'center',
                                            'valign': 'vcenter',})
    print branch_codes
    for branch_code in branch_codes:
        worksheet = workbook.add_worksheet(
            app.config["BRANCH_CODENAMES"][branch_code])
        student = collection.find_one({"college_code": college_code,
                                       "year": year,
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


def college_wise_excel(college_code='027', year='1'):
    year = str(year)
    workbook = xlsxwriter.Workbook(
        'college_summary_excel_year_' + str(year) + '.xlsx')
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
    year = str(year)
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
        colg_avg = round(colg_avg, 2)
        #print 'colg_avg' + str(colg_avg) + 'is for..' +str(colg_code)
        max_mark = app.config['MAX_MARKS_YEARWISE'][year]
        percent = colg_avg / max_mark * 100
        percent = round(percent, 2)
        avg_dict = {colg_code: [colg_avg, percent]}
        avg_list.append(avg_dict)

    # now making excel
    workbook = xlsxwriter.Workbook('ext_avg_year_' + year + '.xlsx')
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


def sec_wise_ext(year):
    college_code = "027"
    year = str(year)
    collection = connection.test.students
    branch_codes = collection.find({"year": year,
                                    "college_code": college_code}).distinct('branch_code')
    # print "distinct branch codes: ", branch_codes
    workbook = xlsxwriter.Workbook('section_wise_year_' + year + '.xlsx')
    merge_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter"
    })
    cell_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
    }
    )
    branch_codenames = app.config["BRANCH_CODENAMES"]
    branch_names = app.config["BRANCH_NAMES"]
    for branch_code in branch_codes:
        branch_codename = branch_codenames[branch_code]
        branch_name = branch_names[branch_code]
        worksheet = workbook.add_worksheet(branch_codename)
        r, c = 0, 0
        worksheet.merge_range(r, c, r, c+6, branch_name, merge_format)
        worksheet.write(r+1, c, "Subject", merge_format)
        worksheet.set_column("A:A", 50)
        worksheet.write(r+1, c+1, "Code", merge_format)
        worksheet.set_column("B:B", 10)
        worksheet.write(r+1, c+2, "Total Students", merge_format)
        worksheet.set_column("C:C", 15)
        worksheet.write(r+1, c+3, "External Avg", merge_format)
        worksheet.set_column("D:D", 15)
        worksheet.write(r+1, c+4, "Pass %", merge_format)
        worksheet.write(r+1, c+5, "Fail %", merge_format)
        worksheet.write(r+1, c+6, "Incomplete %", merge_format)
        worksheet.set_column("G:G", 15)
        r += 3

        sections = collection.find({'year': year,
                                    'college_code': college_code,
                                    'branch_code': branch_code}).distinct('section')
        for section in sections:
            worksheet.merge_range(r, c, r, c+6, "Section: " + section, merge_format)
            r += 1

            sec_data = {}
            print section
            subjects = []
            students = collection.find({'year': year,
                                        'college_code': college_code,
                                        'branch_code': branch_code,
                                        'section':section})

            print collection.find({'year': year,
                                        'college_code': college_code,
                                        'branch_code': branch_code,
                                        'section':section}).count()

            for student in students:
                for mark_dict in student['marks']:
                    sub = {mark_dict['sub_code']: mark_dict['sub_name']}
                    if sub not in subjects:
                        subjects.append(sub)
            print subjects
            print len(subjects)
            for subject in subjects:
                ext_total = 0
                pass_count = 0
                fail_count = 0
                sub_code = subject.keys()[0]
                sub_name = subject.values()[0]
                students = collection.find({'year': year,
                                            'college_code': college_code,
                                            'branch_code': branch_code,
                                            'section': section,
                                            'carry_status': {'$ne': 'INCOMP'}})
                incomp_count = collection.find({'year': year,
                                                'college_code': college_code,
                                                'branch_code': branch_code,
                                                'section': section,
                                                'carry_status': 'INCOMP'}).count()
                number_of_students = 0
                for student in students:
                    marks = student['marks']
                    for m in marks:
                        if m['sub_code'] == sub_code:
                            ext_total = ext_total + m['marks'][0]
                            number_of_students += 1
                            if sub_code in student['carry_papers']:
                                fail_count += 1
                            else:
                                pass_count += 1
                sec_data['sub_name'] = sub_name
                sec_data['sub_code'] = sub_code
                sec_data['total'] = number_of_students
                if number_of_students == 0:

                    sec_data['ext_avg'] = '-'
                    sec_data['pass'] = '-'
                    sec_data['fail'] = '-'
                    sec_data['incomp'] = '-'
                else:
                    sec_data['ext_avg'] = round(float(ext_total) / number_of_students, 2)
                    sec_data['pass'] = round(float(pass_count) / number_of_students * 100, 2)
                    sec_data['fail'] = round(float(fail_count) / number_of_students * 100, 2)
                    sec_data['incomp'] = round(float(incomp_count) / number_of_students * 100, 2)
                worksheet.write(r, c, sec_data['sub_name'], cell_format)
                worksheet.write(r, c+1, sec_data['sub_code'], cell_format)
                worksheet.write(r, c+2, sec_data['total'], cell_format)
                worksheet.write(r, c+3, sec_data['ext_avg'], cell_format)
                worksheet.write(r, c+4, sec_data['pass'], cell_format)
                worksheet.write(r, c+5, sec_data['fail'], cell_format)
                worksheet.write(r, c+6, sec_data['incomp'], cell_format)
                r += 1


def faculty_performance(year):
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
    worksheet.write(r, c+3, 'Total Avg', heading_format)
    worksheet.write(r, c+4, 'Pass %', heading_format)
    worksheet.write(r, c+5, 'Section', heading_format)
    worksheet.write(r, c+6, "Students in sec", heading_format)
    r += 1
    sub_details = dict()
    students = collection.find({'year': year, 'college_code': college_code,
                                'carry_status': {'$ne': 'INCOMP'}})
    section_faculty_info = get_section_faculty_info()

    for student in students:
        carry_papers = student['carry_papers']
        for mark_dict in student['marks']:
            sub_code = mark_dict['sub_code']
            sub_name = mark_dict['sub_name']
            sub_sec_fac = section_faculty_info.get(sub_code, {})
            if (sub_code[:2] == "GP" or sub_code[1:3] == "GP" or
                    sub_code[-2] == "5"):
                continue
            if sub_code in carry_papers:
                num_carry = 1
            else:
                num_carry = 0
            if not sub_details.get((sub_code, sub_name)):
                if sub_code[1:3] == "OE":
                    for faculty, sections in sub_sec_fac.iteritems():
                        section_str = ','.join(sections)
                        if student['section'] in sections:
                            sub_details[(sub_code, sub_name)] = {
                                section_str: {
                                    'ext_tot': mark_dict['marks'][0],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': faculty
                                }
                            }
                else:
                    for faculty, sections in sub_sec_fac.iteritems():
                        if student['section'] in sections:
                            sub_details[(sub_code, sub_name)] = {
                                student['section']: {
                                    'ext_tot': mark_dict['marks'][0],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': faculty
                                }
                            }
                            break
                    else:
                        sub_details[(sub_code, sub_name)] = {
                            student['section']: {
                                'ext_tot': mark_dict['marks'][0],
                                'num_tot': 1,
                                'marks_tot': sum(mark_dict['marks']),
                                'num_carry': num_carry,
                                'faculty': 'not available'
                            }
                        }
            else:
                sub_dict = sub_details[(sub_code, sub_name)]
                if sub_code[1:3] == "OE":
                    for faculty, sections in sub_sec_fac.iteritems():
                        section_str = ','.join(sections)
                        if student['section'] in sections:
                            if section_str not in sub_dict:
                                sub_details[(sub_code, sub_name)][section_str] = {
                                    'ext_tot': mark_dict['marks'][0],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': faculty
                                }
                            else:
                                sub_details[(sub_code, sub_name)][section_str]['ext_tot'] += mark_dict['marks'][0]
                                sub_details[(sub_code, sub_name)][section_str]['num_tot'] += 1
                                sub_details[(sub_code, sub_name)][section_str]['marks_tot'] += sum(mark_dict['marks'])
                                sub_details[(sub_code, sub_name)][section_str]['num_carry'] += num_carry

                else:
                    for faculty, sections in sub_sec_fac.iteritems():
                        if student['section'] in sections:
                            if not sub_dict.get(student['section']):
                                sub_details[(sub_code, sub_name)][student['section']] = {
                                    'ext_tot': mark_dict['marks'][0],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': faculty
                                }
                            else:
                                sub_details[(sub_code, sub_name)][student['section']]['ext_tot'] += mark_dict['marks'][0]
                                sub_details[(sub_code, sub_name)][student['section']]['num_tot'] += 1
                                sub_details[(sub_code, sub_name)][student['section']]['marks_tot'] += sum(mark_dict['marks'])
                                sub_details[(sub_code, sub_name)][student['section']]['num_carry'] += num_carry
                            break
                    else:
                        if sub_details[(sub_code, sub_name)].get(student['section']):
                            sub_details[(sub_code, sub_name)][student['section']]['ext_tot'] += mark_dict['marks'][0]
                            sub_details[(sub_code, sub_name)][student['section']]['num_tot'] += 1
                            sub_details[(sub_code, sub_name)][student['section']]['marks_tot'] += sum(mark_dict['marks'])
                            sub_details[(sub_code, sub_name)][student['section']]['num_carry'] += num_carry
                        else:
                            sub_details[(sub_code, sub_name)][student['section']] = {
                                'ext_tot': mark_dict['marks'][0],
                                'num_tot': 1,
                                'marks_tot': sum(mark_dict['marks']),
                                'num_carry': num_carry,
                                'faculty': 'not available'
                        }


    for sub_tup in sub_details:
        sub_dict = sub_details[sub_tup]
        num_sections = len(sub_dict)
        if num_sections > 1:
            worksheet.merge_range(r, c, r + num_sections - 1, c,
                                  sub_tup[0] + (('(' + sub_tup[1] + ')')
                                                if sub_tup[1] else ''),
                                  cell_format)
        else:
            worksheet.write(r, c,
                            sub_tup[0] + (('(' + sub_tup[1] + ')')
                                          if sub_tup[1] else ''),
                            cell_format)
        for section in sub_dict:
            section_dict = sub_dict[section]
            faculty = section_dict['faculty']
            num_tot = section_dict['num_tot']
            ext_avg = float(section_dict['ext_tot']) / num_tot
            ext_avg = round(ext_avg, 2)
            tot_avg = float(section_dict['marks_tot']) / num_tot
            tot_avg = round(tot_avg, 2)
            num_carry = section_dict['num_carry']
            num_pass = num_tot - num_carry
            pass_percent = round(float(num_pass) / num_tot * 100, 2)
            worksheet.write(r, c+1, faculty, cell_format)
            worksheet.write(r, c+2, ext_avg, cell_format)
            worksheet.write(r, c+3, tot_avg, cell_format)
            worksheet.write(r, c+4, pass_percent, cell_format)
            worksheet.write(r, c+5, section, cell_format)
            worksheet.write(r, c+6, num_tot, cell_format)
            r += 1

    workbook.close()


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
        "valign": "vcenter"
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
    sub_names = dict()
    for student in students:
        college_code = student['college_code']
        carry_papers = student['carry_papers']
        for mark_dict in student['marks']:
            sub_code = mark_dict['sub_code']
            if not sub_names.get(sub_code):
                sub_names[sub_code] = mark_dict['sub_name']
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
        worksheet.write(r, c+1, sub_names.get(sub_code), cell_format)
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


def get_section_faculty_info():
    """
    reads information from excel and returns dictionary of section faculty info
    :return: dict containing subject, section and faculty information
    """
    wb = open_workbook("/home/animesh/Devel/analyzer/Section-Faculty Information/subject_section_faculty.xlsx")
    sheet = wb.sheet_by_index(0)
    section_faculty_info = dict()
    row, col = 0, 0
    for row in range(sheet.nrows):
        sub_code, sec, faculty_name = (sheet.cell_value(row, col),
                                       sheet.cell_value(row, col+1),
                                       sheet.cell_value(row, col+2))
        if sub_code:
            sub_code, sec, faculty_name = (sub_code.strip(),
                                           sec.strip(),
                                           faculty_name.strip().upper())
            if sub_code[3] == '-' or sub_code[3] == ' ':
                sub_code = sub_code[:3] + sub_code[4:]
            if not sub_code in section_faculty_info:
                section_faculty_info[sub_code] = dict()
                section_faculty_info[sub_code][faculty_name] = [sec, ]
            else:
                if faculty_name in section_faculty_info[sub_code]:
                    if sec not in section_faculty_info[sub_code][faculty_name]:
                        section_faculty_info[sub_code][faculty_name].append(sec)
                else:
                    section_faculty_info[sub_code][faculty_name] = [sec, ]

    return section_faculty_info


faculty_performance(4)