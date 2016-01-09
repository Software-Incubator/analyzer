import requests
import urllib.request, random
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
        captcha = input('enter captcha:   ')
        for rollno in range(1302710001, 1302710124):
	        login_data = get_login_credentials(soup, rollno, captcha)
	        response2 = s.post(url, data=login_data)
	        #print(response)
	        soup2 = BeautifulSoup(response2.text, 'html5lib')
	        try:
	        	name = soup2.find(id='ctl00_ContentPlaceHolder1_lblName').string
	        	print(name)
	        except AttributeError:
	        	print('rollno does not exist')
        #n = name.string
        #print(soup2)

def dnld_captcha(imgurl):
	name = str(random.randrange(1,1000)) + '.gif'
	urllib.request.urlretrieve(imgurl, name)


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

