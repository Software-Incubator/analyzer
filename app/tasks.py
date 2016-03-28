import requests, urllib2 ,urllib, os
import time
import sys


from bs4 import BeautifulSoup
from .models import Student
from app import connection, app
from datetime import datetime
from requests.exceptions import ConnectionError


def get_in_session(session, url):
    """gets in session and exits if connection error
    persists
    """
    i = 0
    while i < 5:
        try:
            response = session.get(url)
            return response
        except ConnectionError:
            print("Connection Error")
            time.sleep(1)
            i += 1
    print("Connection Error")
    sys.exit(0)

def post_in_session(session, url, post_data):
    """posts data to the url with post_data to
    the url and exits if connection error persists
    """
    i = 0
    while i < 5:
        try:
            response = session.post(url, data=post_data)
            return response
        except ConnectionError:
            print("Connection Error")
            time.sleep(1)
            i += 1
    print("Connection Error")
    sys.exit(0)


def roll_num_generator(college_code='027', year=2):
    """generates the first roll numbers of all the branches
    :param: college_code: code of the college of which the first
    roll numbers are generated
    :param: year: year of which the results are to be generated
    :return: list of first roll numbers of each branch specified
    in config file
    """
    branch_codes = app.config['BRANCH_CODES']
    curr_year = datetime.now().year % 100
    year_code = str(curr_year - year - 1)
    roll_nums = [year_code + college_code + branch_code + '001' for
                 branch_code in branch_codes]
    if year >= 2:
        lat_year_code = str(curr_year - year)
        lat_roll_nums = [lat_year_code + college_code + branch_code + '901' for
                         branch_code in branch_codes]
        roll_nums += lat_roll_nums
    return roll_nums


def get_result(session, login_data, year=2):
    """
    gets and saves the result data of given roll number
    :param: roll_num: int, roll number of which the result
    is to be saved
    :param: year: int, year of which the result is to be fetched
    :return: True if result saved successfully, False otherwise
    """
    with session as s:
        url = app.config['URLS'][year]  # getting url from config file
        # login to student's account
        response2 = post_in_session(s, url, post_data=login_data)
        soup = BeautifulSoup(response2.text, 'html.parser')
        try:
            name = soup.find(
                id='ctl00_ContentPlaceHolder1_lblName'
            ).string.strip()
            print("getting result of ",name)
        except AttributeError:
            print('rollno does not exist')
            return False

        colg_code = soup.find(
            id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[1][:-1]
        colg_name =  soup.find(
            id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[0]
        fathers_name = soup.find(
            id='ctl00_ContentPlaceHolder1_lblF_NAME'
        ).string.strip()
        roll_no = soup.find(
            id='ctl00_ContentPlaceHolder1_lblROLLNO'
        ).string
        print(" with roll number: ", roll_no)
        status = soup.find(
            id='ctl00_ContentPlaceHolder1_lblStatus'
        ).string.strip()
        branch_info = soup.find(
            id='ctl00_ContentPlaceHolder1_lblCourse'
        ).string.split('.')[2].split('(')
        branch_name = branch_info[0].lstrip()
        branch_code = branch_info[1][:-1]


        # for marks
        marks_tables = soup.find_all('table')[1].find_all('td', width='50%')
        flag = True
        # marks is a list whose first element is odd sem marks,
        # second element is even sem marks and third element is even sem 
        # marks and third is sum of all marks
        marks = []  # for total result
        # this dict also contains a key of sem whose value is the sem of
        # his result
        # this loop iterates on both the tables

        for marks_table in marks_tables:
            in_tables = marks_table.find_all('table')
            if flag:
                sem = (in_tables[0].find_all('tr')[0].find(
                    'th'
                ).string).split()[0][0]

            mark_list = list()  # for a semester
            in_table = in_tables[1]
            marks_rows = in_table.find_all('tr')[1: ]
            # this loop iterates on all subjects of a semester
            for marks_row in marks_rows:
                sub_dict = dict()  # one for each subject
                row_cells = marks_row.find_all('span')  # details of subject
                sub_code = row_cells[0].string.strip()
                sub_name = row_cells[1].string
                sub_dict['sub_code'] = sub_code
                sub_dict['sub_name'] = sub_name
                sub_marks = []  # different kind of marks of a subject
                try:
                    external = float(row_cells[2].string)
                except ValueError:
                    external = 0
                sub_marks.append(external)
                try:
                    internal = float(row_cells[3].string)
                except (ValueError, IndexError):
                    internal = 0
                sub_marks.append(internal)
                try:
                    carry_over = float(row_cells[4].string)
                except (ValueError, TypeError, IndexError):
                    carry_over = 0
                try:
                    credit = float(row_cells[5].string)
                except (ValueError,TypeError):
                    credit = 0
                except IndexError:
                    credit = 0
                sub_marks.append(credit)
                flag = False

                sub_dict['sub_marks'] = sub_marks
                mark_list.append(sub_dict)  # add each subject marks to list

            marks.append(mark_list)  # append semester marks to main list

        max_marks = float(soup.find(id ='ctl00_ContentPlaceHolder1_lblSTAT_8MRK'
                                  ).string)
        marks.append(max_marks)
        carry_papers = soup.find(
                id='ctl00_ContentPlaceHolder1_lblCarryOver').text.split(',')

        carry_papers = [carry_code.rstrip() for carry_code in carry_papers]
        # print carry_papers
        # print max_marks

        student_data = {
            'roll_no': roll_no,
            'name': name,
            'sem': sem,
            'father_name': fathers_name,
            'branch_code': branch_code,
            'branch_name': branch_name,
            'college_code': colg_code,
            'college_name': colg_name,
            'marks': marks,
            'carry_papers': carry_papers,
        }

        collection  = connection['test'].students
        data = collection.Student(student_data)
        data.save()
        print("result saved for roll number: ", data['roll_no'])
        print "Semester" + str(sem)
    # return True if result saved succussfully
    return True

def get_college_results(college_code='027', year=2):
    """
    gets the result of all branches of the given college code
    of all years
    :param: college_code: str, code of the college of which the result
    to be fetched
    :param: year: int, year of which the result is asked
    :return: True if successfully fetched all the results, False otherwise
    """
    roll_nums = roll_num_generator(college_code='027', year=2)
    with requests.Session() as s:
        url = app.config['URLS'][year]  # getting url from config file
        response = get_in_session(s, url)
        plain_text = response.text
        soup = BeautifulSoup(plain_text, 'html.parser')
        img = soup.find('img')
        imgurl = img.get('src')
        dnld_captcha(imgurl)
        captcha = raw_input('enter captcha: ')
        for roll_num in roll_nums:  # for first roll number of each branch
            r_num=int(roll_num)
            count = 0  # counts number of consecutive failures
            i = 0
            while count <= 5:
                print "Roll number: ", r_num + i
                login_data = get_login_credentials(soup, r_num + i, captcha)
                result = get_result(session=s, login_data=login_data, year=year)
                if result:
                    count = 0
                else:
                    count += 1
                i += 1
            print("Fetched {} results, first roll num: {}".format(
                (i - count), roll_num
            ))
    return True


def dnld_captcha(imageurl):
    name = 'captcha.gif'
    file_path = os.path.join(os.getcwd(), name)
    urllib.urlretrieve(imageurl, file_path)


def get_login_credentials(soup, rollno, captcha):
    data1 = str(soup.find(id='__VIEWSTATE')['value'])
    data2 = str(soup.find(id='__VIEWSTATEGENERATOR')['value'])
    data3 = str(soup.find(id='__EVENTVALIDATION')['value'])

    login_credentials = {
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': data1,
        '__VIEWSTATEGENERATOR': data2,
        '__EVENTVALIDATION': data3,
        'ctl00$ContentPlaceHolder1$txtRoll': str(rollno),
        'ctl00$ContentPlaceHolder1$txtcapture': captcha,
        'ctl00$ContentPlaceHolder1$btnSubmit': 'Submit'
    }
    return login_credentials


get_college_results(college_code='027', year=2)

