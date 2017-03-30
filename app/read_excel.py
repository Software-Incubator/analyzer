import os
import xlrd
import re
from app.models import connection
from pymongo import MongoClient
from config import BRANCH_NAMES, BRANCH_CODENAMES

GP_EXT = 0


def read_excel(year=3, branch_code=31, even_sem=False, filename=None):
    if filename is None:
        fname = os.getcwd() + '/Student_Reports/' + 'B_Tech' + BRANCH_CODENAMES[str(branch_code)] + str(
            year) + 'Yr.xlsx'
    else:
        fname = os.getcwd() + '/Student_Reports/' + str(filename)

    xl_workbook = xlrd.open_workbook(fname)

    client = MongoClient()
    db = client['test']
    collection = db.students

    sheet = xl_workbook.sheet_by_index(0)

    tot_rows = sheet.nrows
    tot_cols = sheet.ncols
    sub_dict_tot = dict()

    # Indicates start of subject columns in excel;hardcoded; will need to be changed as per excel
    sub_start = 5
    col = sub_start
    keys = map(unicode, ['roll_no', 'name', 'father_name'])
    branch_name = unicode(BRANCH_NAMES[str(branch_code)])
    branch_code = unicode(branch_code)
    semester = unicode(int(year) * 2 - (0 if even_sem else 1))

    sub_code_pattern = re.search(r'\w+\d{3}', sheet.cell(0, col).value)

    while sub_code_pattern:
        cell_value = sheet.cell(0, col).value

        if 'AUC' in cell_value:
            pass

        elif 'NGP' in cell_value:

            keys.append(cell_value[sub_code_pattern.start():sub_code_pattern.end()])
            col -= 3

        else:

            keys.append(cell_value[sub_code_pattern.start():sub_code_pattern.end()])
        # Increment in col will need to be adjusted as per the excel
        col += 4
        sub_code_pattern = re.search(r'\w+\d{3}', sheet.cell(0, col).value)

    col -= 1

    for row in range(1, tot_rows):

        student = dict(
            {u'aggregate_marks': 0, u'year': unicode(year), u'college_name': u'Ajay Kumar Garg Engineering College',
             u'marks': {semester: []}, u'branch_name': branch_name, u'branch_code': branch_code,
             u'college_code': u'027'})
        column = 1
        index = 0
        aggregate_marks = 0
        ext_tot = int_tot = 0
        while column < tot_cols:

            cell_value = unicode(sheet.cell(row, column).value)
            if column < sub_start:

                if sheet.cell(0, column).value == 'Status':
                    index -= 1

                else:
                    student[keys[index]] = cell_value
                index += 1

            # 'MarksObt2 needs to be changed as per the excel
            elif sheet.cell(0, column).value == 'MarksObt':

                sub_dict_tot[u'sub_code'] = u'TOT'
                sub_dict_tot[u'sub_name'] = u'total'
                sub_dict_tot[u'marks'] = ([ext_tot, int_tot])
                student[u'marks'][semester].append(sub_dict_tot)

            elif column >= sub_start and column <= col:

                if 'AUC' in sheet.cell(0, column).value:
                    index -= 1

                elif 'OE0' in sheet.cell(0, column).value and float(cell_value) == 0.0 and float(
                        sheet.cell(row, column + 1).value) == 0.0:
                    pass

                elif 'NGP' in sheet.cell(0, column).value:

                    gp_int = float(cell_value)
                    sub_dict = {u'sub_code': keys[index], u'sub_name': u'',
                                u'marks': [GP_EXT, float(gp_int)]}

                    ext_tot += GP_EXT
                    int_tot += gp_int
                    aggregate_marks += gp_int
                    student[u'marks'][semester].append(sub_dict)
                    column -= 3

                else:
                    print cell_value, sheet.cell(0,column).value
                    sub_dict = {u'sub_code': keys[index], u'sub_name': u'',
                                u'marks': [float(cell_value), float(sheet.cell(row, column + 1).value)]}
                    aggregate_marks += float(sheet.cell(row, column + 2).value)
                    ext_tot += float(cell_value)
                    int_tot += float(sheet.cell(row, column + 1).value)
                    student[u'marks'][semester].append(sub_dict)
                column += 3
                index += 1

            elif sheet.cell(0, column).value == 'MaxMarks':
                student[u'max_marks'] = cell_value
                student[u'max_marks'] = cell_value

            elif sheet.cell(0, column).value == 'COP':

                carry_papers = cell_value.strip().split(':')

                if len(carry_papers) > 1:
                    carry_papers = carry_papers[1].split(',')
                    carry_papers = [sub_code.strip() for sub_code in carry_papers]
                elif len(carry_papers) == 1:
                    carry_papers = carry_papers[0].split(',')
                    carry_papers = [sub_code.strip() for sub_code in carry_papers]

                else:
                    carry_papers = list()

                student[u'carry_papers'] = carry_papers
            elif sheet.cell(0, column).value == 'ReStatus':
                carry_status = cell_value[4]

                student[u'carry_status'] = carry_status

            student[u'section'] = u''
            column += 1

        student[u'aggregate_marks'] = unicode(aggregate_marks)

        print student
        collection.insert(student)

    return True
