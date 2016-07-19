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


def roll_num_generator(college_codes=('027',), year=2, mca=False):
    """generates the first roll numbers of all the branches
    :param college_codes: code of the college of which the first
    roll numbers are generated
    :param year: year of which the results are to be generated
    :param mca: boolean, whether the roll numbers to be generated for mca
    :return: list of first roll numbers of each branch specified
    in config file
    """
    if mca:
        branch_codes = app.config['MCA_BRANCH_CODES']
    else:
        branch_codes = app.config['BRANCH_CODES']
    curr_year = datetime.now().year % 100
    year_code = str(curr_year - year)
    roll_nums = list()
    for college_code in college_codes:
        roll_nums += [year_code + college_code + branch_code + '001' for
                      branch_code in branch_codes]
    if year >= 2:
        lat_roll_nums = list()
        lat_year_code = str(curr_year - year + 1)
        for college_code in college_codes:
            lat_roll_nums += [lat_year_code + college_code + branch_code +
                              '901' for branch_code in branch_codes]
        roll_nums += lat_roll_nums
    return roll_nums


def get_result(session, login_data, year=2, mca=False):
    """
    gets and saves the result data of given roll number
    :param session: session object, sesison with breaked captcha
    is to be saved
    :param year: int, year of which the result is to be fetched
    :param mca: bool, True if the result to be fetched for mca
    :return: True if result saved successfully, False otherwise
    """
    with session as s:
        if mca:
            url = app.config['MCA_URLS'][year]  # getting url from config file
        else:
            url = app.config['URLS'][year]
        # login to student's account
        response2 = post_in_session(s, url, post_data=login_data)
        soup = BeautifulSoup(response2.text, 'html.parser')
        try:
            name = soup.find(
                id='ctl00_ContentPlaceHolder1_lblName').string.strip()
            print(name)
        except AttributeError:
            return None
        colg_code = soup.find(
            id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[1][:-1]
        colg_name = soup.find(
            id='ctl00_ContentPlaceHolder1_lblInstName'
        ).string.split('(')[0]
        fathers_name = soup.find(
            id='ctl00_ContentPlaceHolder1_lblF_NAME'
        ).string.strip()
        roll_no = soup.find(
            id='ctl00_ContentPlaceHolder1_lblROLLNO'
        ).string
        status = soup.find(id='ctl00_ContentPlaceHolder1_lblCarryOver').string
        if status:
            status = status.strip()
        if mca:
            branch_info = soup.find(
                id='ctl00_ContentPlaceHolder1_lblCourse'
            ).string.split('(')
        else:
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

        marks_tables = soup.find(id='ctl00_ContentPlaceHolder1_divRes').find_all('table')
        # print 'No of sems: ', marks_tables[1].find('tr').find_all('td', attrs={'width': '50%'})
        marks_tables = marks_tables[1].find('tr').find_all('td', attrs={'width': '50%'})
        # marks_rows = marks_table.find('tbody').find_all('tr')[1].find('td').find('table').find('tbody').find_all('tr')


        odd_marks_rows = marks_tables[0].find_all('tr')[1].find_all('tr')
        even_marks_rows = marks_tables[1].find_all('tr')[1].find_all('tr')
        all_marks_rows = [odd_marks_rows, even_marks_rows]
        print 'subject-wise marks: '
        for odd_sem_marks in odd_marks_rows:
            print odd_sem_marks
        for even_sem_marks in even_marks_rows:
            print even_sem_marks
        all_marks_dict = {}

        count = 1
        for marks_rows in all_marks_rows:
            std_total_ext = 0
            std_total_int = 0
            marks = list()
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
                m1 = details[2].string
                if m1:
                    m1 = m1.strip()
                try:
                    m1 = float(m1)
                except (ValueError, TypeError):
                    m1 = 0.0
                try:
                    m2 = details[3].string
                except IndexError:
                    m2 = '0'
                if m2:
                    m2 = m2.strip()
                try:
                    m2 = float(m2)
                except (ValueError, TypeError):
                    m2 = 0.0
                if sub_code[0:2] == 'GP' or sub_code[1:3] == 'GP':
                    m2 = m1
                    m1 = 0.0
                m = [m1, m2]
                std_total_int += m2
                std_total_ext += m1
                print 'External total: ', std_total_ext
                sub_dict = {'sub_name': sub_name, 'sub_code': sub_code,
                            'marks': m}

                marks.append(sub_dict)
            total_dict = {'sub_name': "total",
                          'marks': [std_total_int, std_total_ext],
                          'sub_code': "TOT"}
            marks.append(total_dict)
            all_marks_dict[str(year * 2 - count)] = marks
            count -= 1

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
        if year == 4:
            max_marks = soup.find(
                id='ctl00_ContentPlaceHolder1_lblSTAT_8MRK'
            ).string.strip()
            aggregate_marks = soup.find(
                id='ctl00_ContentPlaceHolder1_lblAGG_MRK'
            ).string.strip()
        else:
            max_marks = soup.find(
                id='ctl00_ContentPlaceHolder1_lblTotalMarks'
            ).string.strip()
            aggregate_marks = u''
        year = str(year)
        year = unicode(year, 'utf-8')
        student_data = {
            'roll_no': roll_no,
            'name': name,
            'year': year,
            'father_name': fathers_name,
            'branch_code': branch_code,
            'branch_name': branch_name,
            'college_code': colg_code,
            'college_name': colg_name,
            'marks': all_marks_dict,
            'max_marks': max_marks,
            'aggregate_marks': aggregate_marks,
            'carry_papers': carry_papers,
            'carry_status': status,
            'section': u''
        }

        collection = connection['test'].students
        data = collection.Student(student_data)
        data.save()
        print("result saved for roll number: ", data['roll_no'])
        print "Year: " + str(year)
    # return True if result saved succussfully
    return True


def get_college_results(college_codes=("027", ), year=4, mca=False):
    """
    gets the result of all branches of the given college code
    of all years
    :param college_codes: str, code of the college of which the result
    to be fetched
    :param year: int, year of which the result is asked
    :param mca: boolean tells whether to get mca result or B.Tech
    :return: True if successfully fetched all the results, False otherwise
    """
    roll_nums = roll_num_generator(college_codes=college_codes,
                                   year=year, mca=mca)
    if mca:
        urls = app.config['MCA_URLS']
    else:
        urls = app.config['URLS']
    with requests.Session() as s:
        url = urls[year]  # getting url corresponding to year
        print 'url for current year: ', url
        response = get_in_session(s, url)
        plain_text = response.text
        soup = BeautifulSoup(plain_text, 'html.parser')
        img = soup.find('img')
        imgurl = img.get('src')
        dnld_captcha(imgurl)
        captcha = raw_input('enter captcha: ')
        for roll_num in roll_nums:  # for first roll number of each branch
            r_num = int(roll_num)
            count = 0  # counts number of consecutive failures
            i = 0
            while count <= 5:
                print "Roll number: ", r_num + i
                login_data = get_login_credentials(soup, r_num + i, captcha)
                result = get_result(session=s, login_data=login_data,
                                    year=year, mca=mca)
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
    college_codes = app.config["COLLEGE_CODES"]
    years = range(1, 5)
    for year in years:
        get_college_results(college_codes=college_codes, year=year)


get_all_result()

def get_all_mca_result():
    college_codes = app.config["COLLEGE_CODES"]
    years = range(1, 4)
    for year in years:
        get_college_results(college_codes=college_codes, year=year, mca=True)
