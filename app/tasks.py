import requests, urllib2 ,urllib, os,re
import random
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
            name = soup.find(id='ctl00_ContentPlaceHolder1_lblName').string
            print(name)
        except AttributeError:
            print('rollno does not exist')
            exit(0)

        colg_code = soup.find(id='ctl00_ContentPlaceHolder1_lblInstName').string.split('(')[1][:-1]
        fathers_name = soup.find(id='ctl00_ContentPlaceHolder1_lblF_NAME').string
        roll_no = soup.find(id='ctl00_ContentPlaceHolder1_lblROLLNO').string
        status = soup.find(id='ctl00_ContentPlaceHolder1_lblStatus').string
        branch_info = soup.find(id='ctl00_ContentPlaceHolder1_lblCourse').string.split('.')[2].split('(')
        branch_name = branch_info[0].lstrip()
        branch_code = branch_info[1][:-1]


        print "College Code: ", colg_code
        print "Father's Name: ", fathers_name
        print "Roll Number: ", roll_no
        print "Status: ", status
        print "Branch Name: ", branch_name
        print "Branch Code: ", branch_code

        student_json = {
            'st_id': roll_no,
            'name': name,
            'father_name': fathers_name,
            'branch': branch_name,
            'college': colg_code,
            'sem': 3,
            'marks': dict(),
           # 'max_marks': max_marks

        }

        #models.connection.User.insert(student_json)
        # for marks
        marks_tables = soup.find_all('table')[1].find_all('td', width='50%')
        flag = True
        for marks_table in marks_tables:
            print("Odd semester marks" if flag else "Even semester marks")
            flag = False
            in_tables = marks_table.find_all('table')
            sem = (in_tables[0].find_all('tr')[0].find('th').string).split()[0][0]
            print 'semester is ' + sem
            in_table = in_tables[1]
            marks_rows = in_table.find_all('tr')[1:]
            for marks_row in marks_rows:
                row_cells = marks_row.find_all('span')
                sub_code = row_cells[0].string
                sub_name = row_cells[1].string
                try:
                    external = int(row_cells[2].string)
                except ValueError:
                    external = 0
                try:
                    internal = int(row_cells[3].string)
                except (ValueError, IndexError):
                    internal = 0
                try:
                    carry_over = int(row_cells[4].string)
                except (ValueError, TypeError, IndexError):
                    carry_over = 0
                try:
                    credit = int(row_cells[5].string)
                except (ValueError, IndexError,TypeError):
                    credit = 0

                '''print sub_code
                print sub_name
                print external
                print internal
                print carry_over
                print credit
                print '****************'  '''

        carry_papers = soup.find(
                id='ctl00_ContentPlaceHolder1_lblCarryOver').text.split(',')

        carry_papers = [carry_code.rstrip() for carry_code in carry_papers]
        print carry_papers
        max_marks = int(soup.find(id ='ctl00_ContentPlaceHolder1_lblSTAT_8MRK').string)
        print max_marks


def dnld_captcha(imageurl):
   #x = random.randrange(1,1000)
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

