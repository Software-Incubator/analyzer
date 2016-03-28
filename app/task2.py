import requests, os, urllib
from bs4 import BeautifulSoup


def get_college_results(college_code='027', year=1):
    """
    gets the result of all branches of the given college code
    of all years
    :param: college_code: str, code of the college of which the result
    to be fetched
    :param: year: int, year of which the result is asked
    :return: True if successfully fetched all the results, False otherwise
    """
    # roll_nums = ['1502710005', '1502710006','1502710007']
    with requests.Session() as s:
        url = 'http://aktu.ac.in/results/gbturesult_11_12/OddSemester2016/frmBtech1Sem.aspx'  # getting url from config file
        # response = get_in_session(s, url)
        response = s.get(url)
        plain_text = response.text
        soup = BeautifulSoup(plain_text, 'html5lib')
        img = soup.find('img')
        imgurl = img.get('src')
        dnld_captcha(imgurl)
        captcha = raw_input('enter captcha:==')
        for rollno in [1502710005]:
            login_data = get_login_credentials(soup, rollno, captcha)
            response2 = s.post(url, data=login_data)
            # print(response)
            soup = BeautifulSoup(response2.text, 'html5lib')
            try:
                name = soup.find(id='ctl00_ContentPlaceHolder1_lblName').string.strip()
                print(name)
            except AttributeError:
                print('rollno does not exist')

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
            status = soup.find(
                    id='ctl00_ContentPlaceHolder1_lblStatus'
            ).string.strip()
            branch_info = soup.find(
                    id='ctl00_ContentPlaceHolder1_lblCourse'
            ).string.split('.')[2].split('(')
            branch_name = branch_info[0].lstrip()
            branch_code = branch_info[1][:-1]

            # print( 'roll_no==', roll_no,
            # 'name',name,
            # 'father_name', fathers_name,
            # 'branch_code', branch_code,
            # 'branch_name', branch_name,
            # 'college_code', colg_code,
            # 'college_name', colg_name,
            # )

            # for marks,, it will contain a list and in that list there will be dictionaries
            # for each and every subjects
            # key for dictionaries "sub_marks","sub_code", "sub_name",,
            #  and there values will be their values
            marks = []
            marks_table = soup.find(id='ctl00_ContentPlaceHolder1_divRes').find_all('table')[1]
            marks_rows = marks_table.find('tbody').find_all('tr')[1].find('td').find('table').find('tbody').find_all('tr')
            for mark in marks_rows[1:]:  # this will iterate in each row
                details = mark.find_all('span')
                try:
                    sub_code = details[0].string.strip()
                except AttributeError:
                    continue
                sub_name = details[1].string.strip()
                m = []
                m.append(details[2].string.strip())
                m.append(details[3].string.strip())
                dict = {'sub_name': sub_name, 'sub_code': sub_code, 'marks': m}

                marks.append(dict)
            # for carry papers
            c = soup.find(id='ctl00_ContentPlaceHolder1_lblCarryOver').string.strip()
            if c[-1] == ',':
                c = c[ : -1]
            carry_papers = c.split(',')
            print carry_papers





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
        # 'ctl00$ContentPlaceHolder1$ddlsession':'2015-2016',
        'ctl00$ContentPlaceHolder1$btnSubmit': 'Submit'
    }
    return login_credentials


get_college_results()
