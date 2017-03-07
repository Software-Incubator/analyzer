import os
import string
import xlsxwriter
import collections

from app import connection, app
from xlrd import open_workbook
from werkzeug.utils import secure_filename

arg_to_string = lambda x: str(x)


# TODO 1, 2, 3, 4
def make_excel(college_code='027', years=(1, ), branch_codes=('10',), output=None):
    years = map(arg_to_string, years)
    branch_codes = map(arg_to_string, branch_codes)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook(
            'main_excel_college-' + college_code + '.xlsx')

    for year in years:
        for branch_code in branch_codes:
            worksheet = workbook.add_worksheet('year-' + year +
                                               '_branch-' + branch_code + 'xlsx')
            collection = connection.test.students
            heading_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
            })
            cell_format = workbook.add_format()
            cell_format.set_text_wrap()

            student = collection.find_one({'branch_code': branch_code,
                                           'college_code': college_code,
                                           'year': year})

            if not student:
                worksheet.merge_range('A1:AO1', 'Result not declared')
                workbook.close()
                return False

            r, c = 0, 0
            worksheet.set_row(r, 30)
            r += 1  # skipping one row for heading
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:C', 30)
            worksheet.merge_range(r, c, r + 1, c, "Roll No.", heading_format)
            worksheet.merge_range(r, c + 1, r + 1, c + 1, "Name", heading_format)
            worksheet.merge_range(r, c + 2, r + 1, c + 2, "Father's Name",
                                  heading_format)
            c += 3
            r += 2
            # taking subject codes in a list
            subject_codes = list()
            students = collection.find({'college_code': college_code,
                                        'year': year,
                                        'branch_code': branch_code})
            # print students.count()
            totals = list()  # to keep record of total marks of all the students

            for student in students:
                # print student
                worksheet.write(r, c - 3, student['roll_no'], cell_format)
                worksheet.write(r, c - 2, student['name'], cell_format)
                worksheet.write(r, c - 1, student['father_name'], cell_format)

                std_marks = student['marks'][str(int(year) * 2 - 1)]  # list of all sub marks

                std_ext_total = std_marks[-1]['marks'][0]  # external total of student
                std_int_total = std_marks[-1]['marks'][1]  # internal total of student
                std_marks = std_marks[:-1]
                for sub_dict in std_marks:
                    sub_code = sub_dict['sub_code']
                    sub_marks = map(int, sub_dict['marks'])
                    external_marks = sub_marks[0]
                    internal_marks = sub_marks[1]
                    total_marks = sum(sub_marks)
                    if sub_code not in subject_codes:
                        subject_codes.append(sub_code)
                    i = subject_codes.index(sub_code)
                    col = c + i * 3  # column to which write the marks of the subject
                    worksheet.write(r, col, external_marks, cell_format)
                    worksheet.write(r, col + 1, internal_marks, cell_format)
                    worksheet.write(r, col + 2, total_marks, cell_format)
                r += 1
                std_tot_total = std_ext_total + std_int_total
                # other student details
                carry_papers = ', '.join(student['carry_papers'])
                carry_status = student['carry_status']
                totals.append((std_ext_total, std_int_total,
                               std_tot_total, carry_papers, carry_status))
            r = 1  # setting row back to 2nd row
            for subject_code in subject_codes:
                worksheet.merge_range(r, c, r, c + 2, subject_code, heading_format)
                worksheet.write(r + 1, c, 'External', cell_format)
                worksheet.write(r + 1, c + 1, 'Internal', cell_format)
                worksheet.write(r + 1, c + 2, 'Total', cell_format)
                c += 3
            cur_col_1 = get_alpha_column(c + 3)
            cur_col_2 = get_alpha_column(c + 4)
            worksheet.set_column(cur_col_1 + ':' + cur_col_1, 30)
            worksheet.set_column(cur_col_2 + ':' + cur_col_2, 15)
            worksheet.merge_range(r, c, r, c + 2, 'Total', heading_format)
            worksheet.merge_range(r, c + 3, r + 1, c + 3, 'Carry Papers',
                                  heading_format)
            worksheet.merge_range(r, c + 4, r + 1, c + 4, 'Carry Status',
                                  heading_format)
            worksheet.write(r + 1, c, 'External', cell_format)
            worksheet.write(r + 1, c + 1, 'Internal', cell_format)
            worksheet.write(r + 1, c + 2, 'Total', cell_format)
            r += 2
            for std_details in totals:
                worksheet.write(r, c, std_details[0], cell_format)
                worksheet.write(r, c + 1, std_details[1], cell_format)
                worksheet.write(r, c + 2, std_details[2], cell_format)
                worksheet.write(r, c + 3, std_details[3], cell_format)
                worksheet.write(r, c + 4, std_details[4], cell_format)
                r += 1
            r, c = 0, 0
            worksheet.merge_range(r, c, r, c + 7 + 3 * len(subject_codes),
                                  'Main Excel\n' + student['college_name'] +
                                  ' - ' + student['branch_name'] + ' - Year:  ' + year,
                                  heading_format)

    workbook.close()

    return workbook


def fail_excel(college_code='027', years=('1',), output=None):
    """
    generates excel for failed students
    :param college_code: code of the college of which the excel is to be made
    :param years: years of students of which excel to be made
    :param output: for download of the excel
    :return: none
    """

    years = map(arg_to_string, years)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('failed_year_excel_workbook' + '.xlsx')
    collection = connection.test.students
    branch_codes = collection.distinct("branch_code")
    heading_format = workbook.add_format({'bold': True,
                                          'align': 'center',
                                          'valign': 'vcenter'})
    cell_format = workbook.add_format({'align': 'center',
                                       'valign': 'vcenter'})

    for year in years:
        for branch_code in branch_codes:
            students = collection.find({"college_code": college_code,
                                        "year": year,
                                        "branch_code": branch_code})
            if students.count():
                student = students.next()
                student1 = students.next()
            else:
                continue
            if not student:
                continue
            worksheet = workbook.add_worksheet(
                'year_' + year + app.config['BRANCH_CODENAMES'][branch_code])
            worksheet.set_row(0, 30)
            worksheet.merge_range('A1:O1', 'List of Failed Students\n' +
                                  student['branch_name'], heading_format)
            worksheet.write('A2', 'S. No.', heading_format)
            worksheet.write('B2', 'Name', heading_format)
            worksheet.write('C2', 'Roll. No.', heading_format)
            worksheet.set_column('B:B', 30)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('P:P', 12)
            cell_list = string.ascii_uppercase[3:]
            i = 0
            stud_marks = len(student['marks'][str(int(year) * 2)]) - 1
            stud_marks1 = len(student1['marks'][str(int(year) * 2)]) - 1
            if stud_marks < stud_marks1:
                student = student1
            num_subjects = len(student['marks'][str(int(year) * 2)]) - 1
            for sub_dict in student['marks'][str(int(year) * 2)][:-1]:
                sub_code = sub_dict['sub_code']
                if sub_code[1:4] == 'OE0':
                    sub_code = 'OE0'
                worksheet.write(cell_list[i] + "2", sub_code, heading_format)
                i += 1
            worksheet.write(cell_list[i] + "2", "No. of Backs", heading_format)
            cp_students = collection.find({"college_code": college_code,
                                           "year": year,
                                           "branch_code": branch_code,
                                           "carry_status": {
                                               "$nin": ["PASS", "PWG", "INC", ]
                                           }
                                           })
            j = 2
            for fail_student in cp_students:
                k = 0
                worksheet.write(j, k, str(j - 1))
                worksheet.write(j, k + 1, fail_student['name'])
                worksheet.write(j, k + 2, fail_student['roll_no'])
                k += 3
                carry_papers = fail_student['carry_papers']
                num_carry = len(fail_student['carry_papers'])

                if num_carry == 0:
                    worksheet.merge_range(j, k,
                                          j, k + num_subjects - 1,
                                          'Carry subjects not provided',
                                          cell_format)
                    k += len(fail_student['marks'])

                else:
                    for mark_dict in fail_student['marks'][str(int(year) * 2)][:-1]:
                        if mark_dict['sub_code'] in carry_papers:
                            worksheet.write(j, k, "F")
                        else:
                            worksheet.write(j, k, "-")
                        k += 1
                worksheet.write(j, k, num_carry)
                j += 1

    workbook.close()
    return True


def akgec_summary(years=('3',), output=None):
    college_code = '027'
    years = map(arg_to_string, years)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook(
            'college_summary_excel_year_' + '.xlsx')
    for year in years:
        worksheet = workbook.add_worksheet()
        merge_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
        })
        excel_format = workbook.add_format()
        excel_format.set_text_wrap()
        collection = connection.test.students
        branch_codes = collection.distinct("branch_code")
        r = 0
        # for heading
        heading = str(app.config['COLLEGE_CODENAMES'][
                          college_code]) + '  YEAR: ' + year + '  2015-16'
        worksheet.write(r, 0, heading, merge_format)
        worksheet.merge_range("A1:I1", heading, merge_format)
        r += 1
        worksheet.write(r, 0, 'S. No', merge_format)
        worksheet.write(r, 1, 'Branch', merge_format)
        worksheet.write(r, 2, 'Total', merge_format)
        worksheet.write(r, 3, 'RND', merge_format)
        worksheet.write(r, 4, 'RND %', merge_format)
        worksheet.write(r, 5, 'INCOMP', merge_format)
        worksheet.write(r, 6, 'RD', merge_format)
        worksheet.write(r, 7, 'CP', merge_format)
        worksheet.write(r, 8, 'Pass', merge_format)
        worksheet.write(r, 9, 'Pass%', merge_format)
        r += 1
        t_total = 0
        t_rnd = 0
        t_incomp = 0
        t_rd = 0
        t_cp = 0
        t_pass_count = 0

        # getting dictionary from config file for the maximum marks
        year_max_dict = app.config['MAX_STUDENTS'].get(year)
        for branch_code in branch_codes:
            all_stud = collection.find({'college_code': college_code,
                                        'branch_code': branch_code,
                                        'year': year})
            incomp_stud = collection.find({'college_code': college_code,
                                           'branch_code': branch_code,
                                           'year': year,
                                           'carry_status': 'INC'})
            total = year_max_dict.get(branch_code)
            if not total:
                continue
            incomp_count = incomp_stud.count()
            print 'No of incomplete students: ', incomp_count
            rd = all_stud.count() - incomp_count
            print 'Result declared: ', rd
            rnd = total - rd
            cp = collection.find({'college_code': college_code,
                                  'branch_code': branch_code,
                                  'year': year,
                                  'carry_status': {'$nin': ["PASS", "PWG", "INC", ]}
                                  }).count()
            print "'No of CPs':", cp

            pass_count = rd - cp
            if rd != 0:
                pass_percent = (float(pass_count) / rd) * 100
                pass_percent = round(pass_percent, 2)
            else:
                pass_percent = '-'

            # computing not declared percent
            rnd_percent = float(rnd) / total * 100
            rnd_percent = round(rnd_percent, 2)

            worksheet.write(r, 0, r - 1, excel_format)
            print branch_code
            worksheet.write(r, 1, app.config['BRANCH_CODENAMES'][branch_code],
                            excel_format)
            worksheet.write(r, 2, total, excel_format)
            worksheet.write(r, 3, rnd, excel_format)
            worksheet.write(r, 4, rnd_percent, excel_format)
            worksheet.write(r, 5, incomp_count, excel_format)
            worksheet.write(r, 6, rd, excel_format)
            worksheet.write(r, 7, cp, excel_format)
            worksheet.write(r, 8, pass_count, excel_format)
            worksheet.write(r, 9, pass_percent, excel_format)
            r += 1
            # for totals
            t_total += total
            t_rnd += rnd
            t_incomp += incomp_count
            t_rd = rd + t_rd
            t_cp = cp + t_cp
            t_pass_count += pass_count
        # computing not declared percentage
        t_rnd_percent = float(t_rnd) / t_total * 100
        t_rnd_percent = round(t_rnd_percent, 2)
        if t_rd != 0:
            t_pass_percent = float(t_pass_count) / t_rd * 100
            t_pass_percent = round(t_pass_percent, 2)
        else:
            t_pass_percent = '-'
        worksheet.write(r, 1, 'Total', merge_format)
        worksheet.write(r, 2, t_total, excel_format)
        worksheet.write(r, 3, t_rnd, excel_format)
        worksheet.write(r, 4, t_rnd_percent, excel_format)
        worksheet.write(r, 5, t_incomp, excel_format)
        worksheet.write(r, 6, t_rd, excel_format)
        worksheet.write(r, 7, t_cp, excel_format)
        worksheet.write(r, 8, t_pass_count, excel_format)
        worksheet.write(r, 9, t_pass_percent, excel_format)
    workbook.close()


def other_college_summary(college_code, year):
    year = str(year)
    workbook = xlsxwriter.Workbook("college_summary_year_" + str(year) +
                                   "_college_" + college_code + ".xlsx")
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    excel_format = workbook.add_format()
    worksheet.set_row(0, 30)
    collection = connection.test.students
    r, c = 0, 0
    # for heading
    heading = ('Inter-Branch Comparison\n' +
               str(app.config['COLLEGE_CODENAMES'][college_code]) +
               '  YEAR: ' + year + '  2015-16')
    worksheet.merge_range("A1:F1", heading, merge_format)
    r += 1
    worksheet.write(r, c, 'S. No', merge_format)
    worksheet.write(r, c + 1, 'Branch', merge_format)
    worksheet.write(r, c + 2, 'RD', merge_format)
    worksheet.write(r, c + 3, 'PCP', merge_format)
    worksheet.write(r, c + 4, 'Pass', merge_format)
    worksheet.write(r, c + 5, 'Pass%', merge_format)
    r += 1
    t_rd = 0
    t_pcp = 0
    t_pass_count = 0

    branch_codes = collection.distinct('branch_code')
    for branch_code in branch_codes:
        rd = collection.find({"college_code": college_code, "year": year,
                              "branch_code": branch_code,
                              "carry_status": {"$ne": "INCOMP"}}).count()
        pcp = collection.find({"college_code": college_code,
                               "year": year,
                               "branch_code": branch_code,
                               "carry_status": {"$nin": ["PASS", "PWG", "INC"]}
                               }).count()
        if rd == 0:
            continue
        pass_count = rd - pcp
        pass_percent = (float(pass_count) / rd) * 100

        worksheet.write(r, c, r - 1, excel_format)
        worksheet.write(r, c + 1, app.config['BRANCH_CODENAMES'][branch_code],
                        excel_format)
        worksheet.write(r, c + 2, rd, excel_format)
        worksheet.write(r, c + 3, pcp, excel_format)
        worksheet.write(r, c + 4, pass_count, excel_format)
        worksheet.write(r, c + 5, pass_percent, excel_format)
        r += 1
        # for totals
        t_pass_count += pass_count
        t_rd = rd + t_rd
        t_pcp = pcp + t_pcp
    if not t_rd:
        return False
    t_pass_percent = (float(t_pass_count) / t_rd) * 100
    worksheet.write(r, c + 1, 'Total', merge_format)
    worksheet.write(r, c + 2, t_rd, excel_format)
    worksheet.write(r, c + 3, t_pcp, excel_format)
    worksheet.write(r, c + 4, t_pass_count, excel_format)
    worksheet.write(r, c + 5, t_pass_percent, excel_format)
    workbook.close()
    return True


def ext_avg(years=(4,), output=None):
    years = map(arg_to_string, years)
    collection = connection.test.students
    college_codes = app.config['COLLEGE_CODES']
    # making a workbook
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('ext_avg_year_' + '.xlsx')

    for year in years:
        max_mark = app.config['MAX_MARKS_YEARWISE'][year]  # max ext marks for year
        avg_list = []
        for colg_code in college_codes:
            students = collection.find({'year': year, 'college_code': colg_code,
                                        'carry_status': {'$ne': 'INC'}})
            t_ext = 0

            print students.count()

            for student in students:
                t_ext += student["marks"][str(int(year) * 2)][-1]["marks"][0]

                # ext = 0
                # for mark in student['marks'][str(int(year)*2)]:
                #     ext = ext + mark['marks'][0]
                # t_ext += ext

            len_students = students.count()

            colg_avg = float(t_ext) / len_students
            colg_avg = round(colg_avg, 2)
            percent = colg_avg / max_mark * 100
            percent = round(percent, 2)
            avg_dict = {colg_code: [colg_avg, percent]}
            avg_list.append(avg_dict)

        worksheet = workbook.add_worksheet()
        merge_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
        })

        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
        })
        worksheet.set_row(0, 30)
        worksheet.merge_range('A1:H1',
                              ('Inter College Comparison of External Marks\n'
                               'Year: %s' % year), merge_format)
        worksheet.merge_range('A2:H2',
                              'Maximum External Marks: ' + str(max_mark),
                              merge_format)
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
            worksheet.write(r, c, app.config['COLLEGE_CODENAMES'][colg_code],
                            merge_format)
            worksheet.write(r + 1, c, colg_dict[colg_code][0], cell_format)
            worksheet.write(r + 2, c, colg_dict[colg_code][1], cell_format)
            c += 1

    workbook.close()
    return True


def sec_wise_ext(years=(4,), output=None):
    college_code = "027"
    years = map(arg_to_string, years)
    collection = connection.test.students
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('section_wise_year_' + '.xlsx')

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
    for year in years:

        branch_codes = collection.find({"year": year,
                                        "college_code": college_code}).distinct(
            'branch_code')
        # print "distinct branch codes: ", branch_codes

        for branch_code in branch_codes:
            branch_codename = branch_codenames[branch_code]
            branch_name = branch_names[branch_code]
            worksheet = workbook.add_worksheet(year + branch_codename)
            r, c = 0, 0
            worksheet.merge_range(r, c, r, c + 5,
                                  'AKGEC Section-wise External Marks', merge_format)
            r += 1
            worksheet.merge_range(r, c, r, c + 5, branch_name, merge_format)
            worksheet.write(r + 1, c, "Subject", merge_format)
            worksheet.set_column("A:A", 50)
            worksheet.write(r + 1, c + 1, "Code", merge_format)
            worksheet.set_column("B:B", 10)
            worksheet.write(r + 1, c + 2, "Total Students", merge_format)
            worksheet.set_column("C:C", 15)
            worksheet.write(r + 1, c + 3, "External Avg", merge_format)
            worksheet.set_column("D:D", 15)
            worksheet.write(r + 1, c + 4, "Pass %", merge_format)
            worksheet.write(r + 1, c + 5, "Fail %", merge_format)
            r += 3

            sections = collection.find({'year': year,
                                        'college_code': college_code,
                                        'branch_code': branch_code}).distinct(
                'section')
            for section in sections:
                worksheet.merge_range(r, c, r, c + 6, "Section: " + section,
                                      merge_format)
                r += 1

                sec_data = {}
                subjects = []
                students = collection.find({'year': year,
                                            'college_code': college_code,
                                            'branch_code': branch_code,
                                            'section': section,
                                            'carry_status': {'$ne': 'INC'}})

                for student in students:
                    for mark_dict in student['marks'][str(int(year) * 2)][:-1]:
                        sub_tup = (mark_dict['sub_code'], mark_dict['sub_name'])
                        if sub_tup not in subjects:
                            subjects.append(sub_tup)
                for subject in subjects:
                    ext_total = 0
                    pass_count = 0
                    fail_count = 0
                    sub_code = subject[0]
                    sub_name = subject[1]
                    students = collection.find({'year': year,
                                                'college_code': college_code,
                                                'branch_code': branch_code,
                                                'section': section,
                                                'carry_status': {'$ne': 'INC'}})
                    number_of_students = 0
                    for student in students:
                        marks = student['marks'][str(int(year) * 2)]
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
                    else:
                        sec_data['ext_avg'] = round(
                            float(ext_total) / number_of_students, 2)
                        sec_data['pass'] = round(
                            float(pass_count) / number_of_students * 100, 2)
                        sec_data['fail'] = round(
                            float(fail_count) / number_of_students * 100, 2)
                    worksheet.write(r, c, sec_data['sub_name'], cell_format)
                    worksheet.write(r, c + 1, sec_data['sub_code'], cell_format)
                    worksheet.write(r, c + 2, sec_data['total'], cell_format)
                    worksheet.write(r, c + 3, sec_data['ext_avg'], cell_format)
                    worksheet.write(r, c + 4, sec_data['pass'], cell_format)
                    worksheet.write(r, c + 5, sec_data['fail'], cell_format)
                    r += 1
    workbook.close()
    return True


def faculty_performance(years=(4,), output=None, file=None):
    years = map(arg_to_string, years)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('faculty_performance.xlsx')

    collection = connection.test.students

    for year in years:
        worksheet = workbook.add_worksheet('YEAR - ' + year)
        worksheet.set_column('A:A', 18)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 17)
        worksheet.set_column('G:G', 21)
        worksheet.set_row(0, 30)
        worksheet.set_row(1, 30)
        college_code = '027'

        heading_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter'
        })
        sep = os.linesep
        worksheet.merge_range('A1:G1',
                              'Faculty Performance Excel - Year: {}'.format(year),
                              heading_format)
        r, c = 1, 0
        worksheet.write(r, c, 'Subject Name', heading_format)
        worksheet.write(r, c + 1, 'Name Of Faculty', heading_format)
        worksheet.write(r, c + 2, 'Internal Avg %s Marks' % sep, heading_format)
        worksheet.write(r, c + 3, 'External Avg %s Marks' % sep, heading_format)
        worksheet.write(r, c + 4, 'Total Avg %s Marks' % sep, heading_format)
        worksheet.write(r, c + 5, 'Pass %', heading_format)
        worksheet.write(r, c + 6, 'Section', heading_format)
        worksheet.write(r, c + 7, "Students in sec", heading_format)
        r += 1
        sub_details = dict()
        students = collection.find({'year': year, 'college_code': college_code,
                                    'carry_status': {'$ne': 'INC'},
                                    'branch_code': {
                                        '$in': ['00', '10', '13', '21',
                                                '31', '32', '40']}
                                    })

        section_faculty_info = get_section_faculty_info(file=file)


        for student in students:
            carry_papers = student['carry_papers']
            for mark_dict in student['marks'][str(int(year) * 2)][:-1]:
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
                        for faculty, sections in sub_sec_fac.items():
                            section_str = ', '.join(sections)
                            if student['section'] in sections:
                                sub_details[(sub_code, sub_name)] = {
                                    section_str: {
                                        'ext_tot': mark_dict['marks'][0],
                                        'int_tot': mark_dict['marks'][1],
                                        'num_tot': 1,
                                        'marks_tot': sum(mark_dict['marks']),
                                        'num_carry': num_carry,
                                        'faculty': faculty
                                    }
                                }
                        else:
                            sub_details[(sub_code, sub_name)] = {
                                section_str: {
                                    'ext_tot': mark_dict['marks'][0],
                                    'int_tot': mark_dict['marks'][1],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': 'not available'
                                }
                            }
                    else:
                        for faculty, sections in sub_sec_fac.iteritems():
                            print 'Faculty: ', faculty, '; Sections: ', sections
                            print 'Student section: ', student['section']
                            if student['section'] in sections:
                                sub_details[(sub_code, sub_name)] = {
                                    student['section']: {
                                        'ext_tot': mark_dict['marks'][0],
                                        'int_tot': mark_dict['marks'][1],
                                        'num_tot': 1,
                                        'marks_tot': sum(mark_dict['marks']),
                                        'num_carry': num_carry,
                                        'faculty': faculty
                                    }
                                }
                                break
                        else:
                            print 'Faculty for the subject {} not found'.format(sub_code)
                            sub_details[(sub_code, sub_name)] = {
                                student['section']: {
                                    'ext_tot': mark_dict['marks'][0],
                                    'int_tot': mark_dict['marks'][1],
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
                            section_str = ', '.join(sections)
                            if student['section'] in sections:
                                if section_str not in sub_dict:
                                    sub_dict[section_str] = {
                                        'ext_tot': mark_dict['marks'][0],
                                        'int_tot': mark_dict['marks'][1],
                                        'num_tot': 1,
                                        'marks_tot': sum(mark_dict['marks']),
                                        'num_carry': num_carry,
                                        'faculty': faculty
                                    }
                                else:
                                    sub_dict[section_str][
                                        'ext_tot'] += mark_dict['marks'][0]
                                    sub_dict[section_str][
                                        'int_tot'] += mark_dict['marks'][1]
                                    sub_dict[section_str][
                                        'num_tot'] += 1
                                    sub_dict[section_str][
                                        'marks_tot'] += sum(mark_dict['marks'])
                                    sub_dict[section_str][
                                        'num_carry'] += num_carry
                                break
                        else:
                            if not sub_dict.get(section_str):
                                sub_dict[section_str] = {
                                    'ext_tot': mark_dict['marks'][0],
                                    'int_tot': mark_dict['marks'][1],
                                    'num_tot': 1,
                                    'marks_tot': sum(mark_dict['marks']),
                                    'num_carry': num_carry,
                                    'faculty': 'not available'
                                }
                            else:
                                sub_dict[section_str]['ext_tot'] += mark_dict['marks'][0]
                                sub_dict[section_str]['int_tot'] += mark_dict['marks'][1]
                                sub_dict[section_str]['num_tot'] += 1
                                sub_dict[section_str]['marks_tot'] += sum(mark_dict['marks'])
                                sub_dict[section_str]['num_carry'] += num_carry

                    else:
                        for faculty, sections in sub_sec_fac.iteritems():
                            if student['section'] in sections:
                                if not sub_dict.get(student['section']):
                                    sub_dict[
                                        student['section']] = {
                                        'ext_tot': mark_dict['marks'][0],
                                        'int_tot': mark_dict['marks'][1],
                                        'num_tot': 1,
                                        'marks_tot': sum(mark_dict['marks']),
                                        'num_carry': num_carry,
                                        'faculty': faculty
                                    }
                                else:
                                    sub_dict[
                                        student['section']]['ext_tot'] += \
                                        mark_dict['marks'][0]
                                    sub_dict[
                                        student['section']]['int_tot'] += \
                                        mark_dict['marks'][1]
                                    sub_dict[
                                        student['section']]['num_tot'] += 1
                                    sub_dict[
                                        student['section']]['marks_tot'] += sum(
                                        mark_dict['marks'])
                                    sub_dict[
                                        student['section']][
                                        'num_carry'] += num_carry
                                break
                        else:
                            if sub_dict.get(student['section']):
                                sub_dict[student['section']][
                                    'ext_tot'] += mark_dict['marks'][0]
                                sub_dict[student['section']][
                                    'int_tot'] += mark_dict['marks'][1]
                                sub_dict[
                                    student['section']][
                                    'num_tot'] += 1
                                sub_dict[
                                    student['section']][
                                    'marks_tot'] += sum(mark_dict['marks'])
                                sub_dict[
                                    student['section']][
                                    'num_carry'] += num_carry
                            else:
                                sub_dict[
                                    student['section']] = {
                                    'ext_tot': mark_dict['marks'][0],
                                    'int_tot': mark_dict['marks'][1],
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
            sub_od = collections.OrderedDict(sorted(sub_dict.items(), key=lambda x: x[1]['faculty']))
            for section in sub_od:
                section_dict = sub_od[section]
                faculty = section_dict['faculty']
                num_tot = section_dict['num_tot']
                int_avg = float(section_dict['int_tot']) / num_tot
                ext_avg = float(section_dict['ext_tot']) / num_tot
                ext_avg = round(ext_avg, 2)
                tot_avg = float(section_dict['marks_tot']) / num_tot
                tot_avg = round(tot_avg, 2)
                num_carry = section_dict['num_carry']
                num_pass = num_tot - num_carry
                pass_percent = round(float(num_pass) / num_tot * 100, 2)
                worksheet.write(r, c + 1, faculty, cell_format)
                worksheet.write(r, c + 2, int_avg, cell_format)
                worksheet.write(r, c + 3, ext_avg, cell_format)
                worksheet.write(r, c + 4, tot_avg, cell_format)
                worksheet.write(r, c + 5, pass_percent, cell_format)
                worksheet.write(r, c + 6, section, cell_format)
                worksheet.write(r, c + 7, num_tot, cell_format)
                r += 1

    workbook.close()


def subject_wise(years=('2',), output=None):
    """
    subject wise comparison of marks of all 4 colleges
    :param year: year for which the analysis is done
    :return:
    """
    years = map(arg_to_string, years)
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('subject_wise_year_' + '.xlsx')
    for year in years:

        worksheet = workbook.add_worksheet('year-' + year)
        merge_format = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter"
        })
        cell_format = workbook.add_format({
            "align": "center"
        })
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 40)
        worksheet.set_row(0, 30)
        r, c = 0, 0
        collection = connection.test.students
        college_codes = collection.distinct('college_code')
        worksheet.merge_range(r, c, r, c + 1 + 3 * len(college_codes),
                              ('Inter College Subject-wise Comparison'
                               '\nYear: %s' % year),
                              merge_format)
        r += 1
        college_codenames = app.config["COLLEGE_CODENAMES"]
        worksheet.merge_range(r, c, r + 1, c, "Subject Code", merge_format)
        worksheet.merge_range(r, c + 1, r + 1, c + 1, "Subject Name", merge_format)
        for college_code in college_codes:
            worksheet.merge_range(r, c + 2, r, c + 4,
                                  college_codenames[college_code], merge_format)
            worksheet.write(r + 1, c + 2, "Total", merge_format)
            worksheet.write(r + 1, c + 3, "Pass %", merge_format)
            worksheet.write(r + 1, c + 4, "Fail %", merge_format)
            c += 3
        r += 2
        students = collection.find({"year": year,
                                    "carry_status": {"$ne": "INC"}})

        year_dict = dict()
        sub_names = dict()
        for student in students:
            college_code = student['college_code']
            carry_papers = student['carry_papers']
            for mark_dict in student['marks'][str(int(year) * 2)][:-1]:
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
            worksheet.write(r, c + 1, sub_names.get(sub_code), cell_format)
            c += 2
            sub_dict = year_dict[sub_code]
            for college_code in college_codes:
                college_dict = sub_dict.get(college_code)
                if college_dict:
                    total = college_dict.get("total")
                    fail = college_dict.get("fail")
                    passed = total - fail
                    pass_percent = round(float(passed) / total * 100, 2)
                    fail_percent = round(float(fail) / total * 100, 2)
                    worksheet.write(r, c, total)
                    worksheet.write(r, c + 1, pass_percent)
                    worksheet.write(r, c + 2, fail_percent)
                else:
                    worksheet.write(r, c, '-')
                    worksheet.write(r, c + 1, '-')
                    worksheet.write(r, c + 2, '-')
                c += 3
            r += 1
    workbook.close()
    return True


def pass_percentage(year_range=range(1, 5), output=None):
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('pass_percentage_comparison' + '.xlsx')
    heading_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter"
    })
    cell_format = workbook.add_format({
        "align": "center"
    })
    worksheet = workbook.add_worksheet()
    r, c = 0, 0
    col = connection.test.students
    college_codes = col.distinct('college_code')
    worksheet.merge_range(
        r, c, r, c + len(college_codes),
        "Inter College Pass Percent Comparison",
        heading_format)
    r += 1
    worksheet.write(r, c, "Year", heading_format)
    c += 1
    for college_code in college_codes:
        college_codename = app.config['COLLEGE_CODENAMES'][college_code]
        worksheet.write(r, c, college_codename, heading_format)
        c += 1
    r += 1
    c = 0

    for year in year_range:
        year = str(year)
        worksheet.write(r, c, year, cell_format)
        c += 1
        for college_code in college_codes:
            total_students = col.find({
                'college_code': college_code,
                'year': year,
                'carry_status': {'$ne': 'INC'}
            }).count()
            fail_count = col.find({
                'college_code': college_code,
                'year': year,
                'carry_status': {
                    '$nin': ["PASS", "PWG", "INC"]
                }
            }).count()
            pass_perc = round(float(total_students - fail_count) /
                              total_students * 100, 2)
            worksheet.write(r, c, pass_perc, cell_format)
            c += 1
        r += 1
        c = 0
    workbook.close()
    return True


def branch_wise_pass_percent(years=('2',), output=None):
    """
    creates report with pass percetage of each college for each branch for
    given year
    :param year: int or str, year for which the report is to be generated
    :return: True if successfully created the report
    """
    years = map(arg_to_string, years)
    collection = connection.test.students
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook('inter_college_branch_wise_comparison'
                                       '_year-' + '.xlsx')
    heading_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })
    cell_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
    })
    for year in years:

        worksheet = workbook.add_worksheet()
        r, c = 0, 0
        r += 1  # leave first row for heading
        college_codes = collection.distinct('college_code')
        college_total_info = collections.OrderedDict()
        college_codenames = app.config['COLLEGE_CODENAMES']
        alpha_col = get_alpha_column(c)
        worksheet.set_column(alpha_col + ':' + alpha_col, 30)
        worksheet.set_row(r, 30)
        worksheet.write(r, c, 'Deptt. \\ College', heading_format)
        c += 1
        for college_code in college_codes:
            worksheet.write(r, c, college_codenames[college_code], heading_format)
            c += 1
        r += 1
        # heading
        worksheet.set_row(0, 30)
        worksheet.merge_range(0, 0, 0, c - 1,
                              ('Inter College Branch-wise Pass Percentage\n'
                               'Comparison - Year: {}'.format(year)),
                              heading_format)
        branch_codes = collection.find({'year': year}).distinct('branch_code')
        if '14' in branch_codes:
            branch_codes.remove('14')
        branch_codenames = app.config['BRANCH_CODENAMES']
        for branch_code in branch_codes:
            c = 0
            worksheet.write(r, c, branch_codenames[branch_code], heading_format)
            c += 1
            for college_code in college_codes:
                students_count = collection.find({
                    'year': year,
                    'branch_code': branch_code,
                    'college_code': college_code,
                    'carry_status': {'$ne': 'INC'}
                }).count()
                fail_count = collection.find({
                    'year': year,
                    'branch_code': branch_code,
                    'college_code': college_code,
                    'carry_status': {'$nin': ["PASS", "PWG", "INC"]}
                }).count()
                if college_total_info.get(college_code):
                    college_total_info[college_code][
                        'student_count'] += students_count
                    college_total_info[college_code][
                        'fail_count'] += fail_count
                else:
                    college_total_info[college_code] = dict()
                    college_total_info[college_code][
                        'student_count'] = students_count
                    college_total_info[college_code][
                        'fail_count'] = fail_count
                try:
                    pass_p = float(students_count -
                                   fail_count) / students_count * 100
                    pass_p = round(pass_p, 2)
                except ZeroDivisionError:
                    pass_p = '-'
                worksheet.write(r, c, pass_p, cell_format)
                c += 1
            r += 1
        # for total
        c = 0
        worksheet.write(r, c, 'Total', heading_format)
        c += 1
        for college_code in college_total_info:
            college_info = college_total_info[college_code]
            total_count = college_info['student_count']
            fail_count = college_info['fail_count']
            try:
                pass_p = float(total_count - fail_count) / total_count * 100
                pass_p = round(pass_p, 2)
            except ZeroDivisionError:
                pass_p = '-'
            worksheet.write(r, c, pass_p, heading_format)
            c += 1
    workbook.close()
    return True


def branch_wise_ext_avg(years=('2',), output=None):
    """
    generates report for branchwise comparison of external percentage of
    each college
    :param years: int or str, year for which the analysis has to be done
    :return: True if succesfully generates the report
    """
    years = map(arg_to_string, years)
    collection = connection.test.students
    if output:
        workbook = xlsxwriter.Workbook(output)
    else:
        workbook = xlsxwriter.Workbook(
            'inter_college_branch_wise_external_comparison_year-' + '.xlsx')
    heading_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })
    cell_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
    })
    for year in years:

        worksheet = workbook.add_worksheet()
        r, c = 0, 0
        r += 1  # left first row for heading
        college_codes = collection.distinct('college_code')
        college_total_info = collections.OrderedDict()
        college_codenames = app.config['COLLEGE_CODENAMES']
        alpha_col = get_alpha_column(c)
        worksheet.set_column(alpha_col + ':' + alpha_col, 30)
        worksheet.set_row(r, 30)
        worksheet.write(r, c, 'Deptt. \\ College', heading_format)
        c += 1
        for college_code in college_codes:
            worksheet.write(r, c, college_codenames[college_code], heading_format)
            c += 1
        r += 1
        # heading
        worksheet.set_row(0, 30)
        worksheet.merge_range(0, 0, 0, c - 1,
                              ('Inter College Branch-wise External Percentage\n'
                               'Comparison - Year: {}'.format(year)),
                              heading_format)
        branch_codes = collection.find({'year': year}).distinct('branch_code')

        # removing MCA from the analysis
        if '14' in branch_codes:
            branch_codes.remove('14')
        branch_codenames = app.config['BRANCH_CODENAMES']
        year_max_marks = app.config['MAX_MARKS_YEARWISE'][year]
        for branch_code in branch_codes:
            c = 0
            worksheet.write(r, c, branch_codenames[branch_code], heading_format)
            c += 1
            for college_code in college_codes:
                students = collection.find({
                    'year': year,
                    'branch_code': branch_code,
                    'college_code': college_code,
                    'carry_status': {'$ne': 'INC'}
                })
                student_count = students.count()
                ext_total = 0
                for student in students:
                    student_total = 0
                    for marks_dict in student['marks'][str(int(year) * 2)][:-1]:
                        marks = marks_dict['marks']
                        student_total += marks[0]
                    ext_total += student_total
                if college_total_info.get(college_code):
                    college_info = college_total_info[college_code]
                    college_info['ext_total'] += ext_total
                    college_info['student_count'] += student_count
                else:
                    college_info = dict()
                    college_info['ext_total'] = ext_total
                    college_info['student_count'] = student_count
                    college_total_info[college_code] = college_info
                try:
                    avg_ext = float(ext_total) / student_count
                    avg_ext_percentage = round(avg_ext / year_max_marks * 100, 2)
                except ZeroDivisionError:
                    avg_ext_percentage = '-'
                worksheet.write(r, c, avg_ext_percentage, cell_format)
                c += 1
            r += 1
        # for total
        c = 0
        worksheet.write(r, c, 'Total', heading_format)
        c += 1
        for college_code in college_total_info:
            college_info = college_total_info[college_code]
            total_count = college_info['student_count']
            ext_total = college_info['ext_total']
            try:
                ext_avg_marks = float(ext_total) / total_count
                ext_avg_p = round(ext_avg_marks / year_max_marks * 100, 2)
            except ZeroDivisionError:
                ext_avg_p = '-'
            worksheet.write(r, c, ext_avg_p, heading_format)
            c += 1
    workbook.close()
    return True


# Helper function
def get_section_faculty_info(file=None):
    """
        reads information from excel and returns dictionary of section faculty
         information.
        :return: dict containing subject, section and faculty information
       """

    if file:
        filename = secure_filename(file.filename)
        file.save(app.config['UPLOAD_FOLDER'] + filename)
        wb = open_workbook(filename)
    else:
        wb = open_workbook(
            "/home/nishtha/Desktop/Result-analyser/analyzer/Section-Faculty Information/subject_section_faculty_even_sem_2016.xlsx")
    sheet = wb.sheet_by_index(0)
    section_faculty_info = dict()
    row, col = 0, 0
    for row in range(sheet.nrows):
        sub_code, sec, faculty_name = (sheet.cell_value(row, col),
                                       sheet.cell_value(row, col + 1),
                                       sheet.cell_value(row, col + 2))
        if sub_code:

            sub_code, sec, faculty_name = (sub_code.strip(),
                                           sec.strip(),
                                           faculty_name.strip().upper())

            if faculty_name[-3:] == 'S/I':
                faculty_name = ' '.join(faculty_name.split('  ')[:-1]).strip()

            if sub_code[3] == '-' or sub_code[3] == ' ':
                sub_code = sub_code[:3] + sub_code[4:]

            if len(sec) > 2:
                if sec[2] == ' ':
                    sec = sec[:2] + sec[3:]
                if len(sec) >= 4 and sec[3] == ' ':
                    sec = sec[:3] + sec[4:]

            if sub_code not in section_faculty_info:
                section_faculty_info[sub_code] = dict()
                section_faculty_info[sub_code][faculty_name] = [sec, ]
            else:
                if faculty_name in section_faculty_info[sub_code]:
                    if sec not in section_faculty_info[sub_code][faculty_name]:
                        section_faculty_info[sub_code][faculty_name].append(sec)
                else:
                    section_faculty_info[sub_code][faculty_name] = [sec, ]

    return section_faculty_info


# Helper function
def get_alpha_column(col_num):
    letters = string.ascii_uppercase
    if col_num < 26:
        return letters[col_num]
    else:
        n = col_num // 26
        i = col_num % 26
        return letters[n - 1] + letters[i]

        # faculty_performance()
