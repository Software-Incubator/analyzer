import requests, urllib2 ,urllib, os
from bs4 import BeautifulSoup
#from .models import *


def get_results():
    with requests.Session() as s:
        url = 'http://www.uptu.ac.in/results/gbturesult_11_12/Even2015Result/frmbtech4semester_2015nrqiop.aspx'
        response = s.get(url)
        plain_text = response.text
        soup = BeautifulSoup(plain_text, 'html.parser')
        img = soup.find('img')
        imgurl = img.get('src')
        dnld_captcha(imgurl)
        captcha = raw_input('enter captcha:   ')
        # for time being using only one roll number
        rollno = 1302731024
        login_data = get_login_credentials(soup, rollno, captcha)
        response2 = s.post(url, data=login_data)
        #print(response)
        soup = BeautifulSoup(response2.text, 'html.parser')
        try:
            name = soup.find(id='ctl00_ContentPlaceHolder1_lblName').string.strip()
            print(name)
        except AttributeError:
            print('rollno does not exist')
            exit(0)

        colg_code = soup.find(id='ctl00_ContentPlaceHolder1_lblInstName').string.split('(')[1][:-1]
        fathers_name = soup.find(id='ctl00_ContentPlaceHolder1_lblF_NAME').string.strip()
        roll_no = soup.find(id='ctl00_ContentPlaceHolder1_lblROLLNO').string
        status = soup.find(id='ctl00_ContentPlaceHolder1_lblStatus').string.strip()
        branch_info = soup.find(id='ctl00_ContentPlaceHolder1_lblCourse').string.split('.')[2].split('(')
        branch_name = branch_info[0].lstrip()
        branch_code = branch_info[1][:-1]


        #models.connection.User.insert(student_json)
        # for marks
        marks_tables = soup.find_all('table')[1].find_all('td', width='50%')
        flag = True
        # marks is a list whose first element is odd sem marks,
        # second element is even sem marks and third element is even sem marks and third is sum of all marks
        marks = []
        mark_odd = {}  # dictionary for odd sem marks
        mark_even = {} # dictionary for even sem marks
        # this dict also contains a key of sem whose value is the sem of his result
        # this loop iterates on both the tables
        flag = True
        for marks_table in marks_tables:
            in_tables = marks_table.find_all('table')
            sem = (in_tables[0].find_all('tr')[0].find('th').string).split()[0][0]
            print 'semester is ' + sem
            in_table = in_tables[1]
            marks_rows = in_table.find_all('tr')[1:]
            # this loop iterates on all rows of a particular table
            for marks_row in marks_rows:
                row_cells = marks_row.find_all('span')
                sub_code = row_cells[0].string.split()[0]
                sub_name = row_cells[1].string
                mark_list = []

                try:
                    external = int(row_cells[2].string)
                except ValueError:
                    external = 0
                mark_list.append(external)
                try:
                    internal = int(row_cells[3].string)
                except (ValueError, IndexError):
                    internal = 0
                mark_list.append(internal)
                try:
                    carry_over = int(row_cells[4].string)
                except (ValueError, TypeError, IndexError):
                    carry_over = 0
                try:
                    credit = int(row_cells[5].string)
                except (ValueError,TypeError):
                    credit = 0
                except IndexError:
                    credit = 0
                mark_list.append(credit)

                if flag:
                    mark_odd[sub_code] = mark_list
                else:
                    mark_even[sub_code] = mark_list
            if flag:
                mark_odd['sem'] = sem
            else:
                mark_even['sem'] = sem
            flag = False

        max_marks = int(soup.find(id ='ctl00_ContentPlaceHolder1_lblSTAT_8MRK').string)
        marks = [mark_odd, mark_even,max_marks ]
        carry_papers = soup.find(
                id='ctl00_ContentPlaceHolder1_lblCarryOver').text.split(',')

        carry_papers = [carry_code.rstrip() for carry_code in carry_papers]
        # print carry_papers
        # print max_marks

        student_data = {
            'roll_no': roll_no,
            'name': name,
            'father_name': fathers_name,
            'branch': branch_name,
            'college': colg_code,
            'marks': marks,
            'carry_papers': carry_papers,
            }
        print student_data


def dnld_captcha(imageurl):
   name = 'raghav.gif'
   urllib.urlretrieve(imageurl, os.path.join(os.getcwd(), name)) # download and save image



def get_login_credentials(soup, rollno, captcha):
    data1 = str(soup.find(id='__VIEWSTATE')['value'])
    data2 = str(soup.find(id='__VIEWSTATEGENERATOR')['value'])
    data3 = str(soup.find(id='__EVENTVALIDATION')['value'])
    ''' print(data1)
    print(data2)
    print(data3) '''

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

get_results()

