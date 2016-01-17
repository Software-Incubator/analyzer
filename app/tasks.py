import requests, urllib2 ,urllib, os
import random
from bs4 import BeautifulSoup

def get_results():
    with requests.Session() as s:
        url = 'http://www.uptu.ac.in/results/gbturesult_11_12/Even2015Result/frmbtech4semester_2015nrqiop.aspx'
        response = s.get(url)
        plain_text = response.text
        soup = BeautifulSoup(plain_text, 'html5lib')
        img = soup.find('img')
        imgurl = img.get('src')
        dnld_captcha(imgurl)
        captcha = raw_input('enter captcha:   ')
        # for time being using only one roll number
        rollno = 1302710116
        login_data = get_login_credentials(soup, rollno, captcha)
        response2 = s.post(url, data=login_data)
        #print(response)
        soup = BeautifulSoup(response2.text, 'html5lib')
        try:
            name = soup.find(id='ctl00_ContentPlaceHolder1_lblName').string
            print(name)
        except AttributeError:
            print('rollno does not exist')

        colg_code = soup.find(id='ctl00_ContentPlaceHolder1_lblInstName').string.split('(')[1][:-1]
        fathers_name = soup.find(id='ctl00_ContentPlaceHolder1_lblF_NAME').string
        roll_no = soup.find(id='ctl00_ContentPlaceHolder1_lblROLLNO').string
        status = soup.find(id='ctl00_ContentPlaceHolder1_lblStatus').string
        branch_name = soup.find(id='ctl00_ContentPlaceHolder1_lblCourse').string.split('(')[0]
        branch_code = soup.find(id='ctl00_ContentPlaceHolder1_lblCourse').string.split('(')[1][:-1]

        '''print colg_code
        print fathers_name
        print roll_no
        print status
        print branch_name
        print branch_code '''
        # for marks
        mark = soup.find_all('table')[1].find_all('table')[0]
        # for sem
        sem = (mark.find('th').string).split()[0][0]
        mark_table = mark.find_all('table')[0]
        # this williterate in all rows of a particular semester
        for tr in mark_table.find_all('tr')[1:]:
            sub_code = tr.find_all('span')[0].string
            sub_name = tr.find_all('span')[1].string
            try:
                external = int(tr.find_all('span')[2].string)
            except ValueError:
                external = 0
            try:
                internal = int(tr.find_all('span')[3].string)
            except (ValueError, IndexError):
                internal = 0

            try:
                carry_over = int(tr.find_all('span')[4].string)
            except (ValueError, TypeError, IndexError):
                carry_over = 0
            try:
                credit = int(tr.find_all('span')[5].string)
            except (ValueError, IndexError):
                credit = 0


            print sub_code
            print sub_name
            print external
            print internal
            print carry_over
            print credit
            print '*******************'




def dnld_captcha(imageurl):
   #x = random.randrange(1,1000)
   name = 'raghav.gif'
   urllib.urlretrieve(imageurl, os.path.join(os.getcwd(), name)) # download and save image



def get_login_credentials(soup, rollno, captcha):
    data1 = str(soup.find(id='__VIEWSTATE')['value'])
    data2 = str(soup.find(id='__VIEWSTATEGENERATOR')['value'])
    data3 = str(soup.find(id='__EVENTVALIDATION')['value'])
    '''print(data1)
    print(data2)
    print(data3)'''

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

