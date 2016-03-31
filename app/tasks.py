import requests, urllib2, urllib, os
import time
import sys

from bs4 import BeautifulSoup
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
    year_code = str(curr_year - year)
    roll_nums = [year_code + college_code + branch_code + '001' for
                 branch_code in branch_codes]
    if year >= 2:
        lat_year_code = str(curr_year - year + 1)
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
            name = soup.find(id='ctl00_ContentPlaceHolder1_lblName').string.strip()
            print(name)
        except AttributeError:
            print('rollno does not exist')
            return None
        colg_code = soup.find(
                id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[1][:-1]
        colg_name = soup.find(
                id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[0]
        fathers_name = soup.find(
                id='ctl00_ContentPlaceHolder1_lblFname'
        ).string.strip()
        roll_no = soup.find(
                id='ctl00_ContentPlaceHolder1_lblRollName'
        ).string
        print(" with roll number: ", roll_no)
        status = soup.find(id='ctl00_ContentPlaceHolder1_lblCarryOverStatus').string
        if status:
            status = status.strip()
        branch_info = soup.find(
                id='ctl00_ContentPlaceHolder1_lblCourse'
        ).string.split('.')[2].split('(')
        branch_name = branch_info[0]
        if branch_name:
            branch_name = branch_name.lstrip()
        branch_code = branch_info[1][:-1]

        # for marks, it will contain a list and in that list there
        #  will be dictionaries for each and every subjects
        # key for dictionaries "sub_marks","sub_code", "sub_name",,
        # and there values will be their values
        marks = list()
        marks_table = soup.find(id='ctl00_ContentPlaceHolder1_divRes').find_all('table')[1]
        marks_rows = marks_table.find('tbody').find_all('tr')[1].find('td').find('table').find('tbody').find_all('tr')
        for mark in marks_rows[1:]:  # this will iterate in each row
            details = mark.find_all('span')

            try:
                sub_code = details[0].string
                if sub_code:
                    sub_code = sub_code.strip()
                else:
                    print 'skipping and continuing...first else'
                    continue
            except AttributeError:
                continue
            print sub_code
            if not sub_code:
                print 'skipping and continuing...second else'
                continue
            sub_name = details[1].string
            if sub_name:
                sub_name = sub_name.strip()
            m = list()
            m1 = details[2].string
            if m1:
                m1 = m1.strip()
            try:
                m1 = float(m1)
            except (ValueError, TypeError):
                m1 = 0.0
            m2 = details[3].string
            if m2:
                m2 = m2.strip()
            try:
                m2 = float(m2)
            except (ValueError, TypeError):
                m2 = 0.0
            m.append(m1)
            m.append(m2)
            dict = {'sub_name': sub_name, 'sub_code': sub_code, 'marks': m}
            print "Marks dict: ", dict

            marks.append(dict)
        # for carry papers
        c = soup.find(
            id='ctl00_ContentPlaceHolder1_lblCarryOver').string
        if c:
            c = c.strip()
        if len(c) > 0:
            if c[-1] == ',':
                c = c[:-1]
            carry_papers = c.split(',')
        else:
            carry_papers = list()
        year = str(year)
        year = unicode(year, 'utf-8')
        max_marks = soup.find(id = 'ctl00_ContentPlaceHolder1_lblTotalMarks').string.strip()
        print 'max_marks=', max_marks
        student_data = {
            'roll_no': roll_no,
            'name': name,
            'year': year,
            'father_name': fathers_name,
            'branch_code': branch_code,
            'branch_name': branch_name,
            'college_code': colg_code,
            'college_name': colg_name,
            'marks': marks,
            'max_marks': max_marks,
            'carry_papers': carry_papers,
            'carry_status': status,
        }

        collection  = connection['test'].students
        data = collection.Student(student_data)
        data.save()
        print("result saved for roll number: ", data['roll_no'])
        print "Year: " + str(year)
    # return True if result saved succussfully
    return True


def get_college_results(college_code="027", year=2):
    """
    gets the result of all branches of the given college code
    of all years
    :param college_code: str, code of the college of which the result
    to be fetched
    :param year: int, year of which the result is asked
    :return: True if successfully fetched all the results, False otherwise
    """
    roll_nums = roll_num_generator(college_code=college_code, year=year)
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
    # data2 = str(soup.find(id='__VIEWSTATEGENERATOR')['value'])
    data3 = str(soup.find(id='__EVENTVALIDATION')['value'])

    login_credentials = {
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': data1,
        # '__VIEWSTATEGENERATOR': data2,
        '__EVENTVALIDATION': data3,
        'ctl00$ContentPlaceHolder1$txtRoll': str(rollno),
        'ctl00$ContentPlaceHolder1$txtcapture': captcha,
        'ctl00$ContentPlaceHolder1$btnSubmit': 'Submit'
    }
    return login_credentials


def get_all_result():
    years = [1, 2, 3, 4]
    for college_code in app.config["COLLEGE_CODES"][:]:
        for year in years:
            get_college_results(college_code=college_code, year=year)

get_all_result()
