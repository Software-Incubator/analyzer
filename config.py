import os

BRANCH_CODES = ['00', '10', '13', '21', '31', '32', '40', '20']

MCA_BRANCH_CODES = ['14', ]

COLLEGE_CODENAMES = {
    '027': 'AKGEC',
    '029': 'KIET',
    '032': 'ABES',
    '091': 'JSS',
    '143': 'IMS',
    '033': 'RKGIT',
    '161': 'KEC',
    '030': 'IPEC',
    '222': 'ITS',
}

BRANCH_NAMES = {
    "00": "Civil Engineering",
    "10": "Computer Science and Engineering",
    "13": "Information Technology",
    "20": "Electrical Engineering",
    "31": "Electronics and Communication Engineering",
    "32": "Electronics and Instrumentation Engineering",
    "40": "Mechanical Engineering",
    "14": "Master in Computer Applications",
    "21": "Electrical & Electronics Engineering",
}
BRANCH_CODENAMES = {
    "00": "CE",
    "10": "CSE",
    "13": "IT",
    "20": "EE",
    "21": "EN",
    "31": "ECE",
    "32": "EI",
    "40": "ME",
    "14": "MCA",
}

URLS = {
    1: ("http://results.aktu.ac.in/akturesult/Even2016Resul"
        "t/frmbtech2semester_2016safrgh.aspx"),
    2: ("http://results.aktu.ac.in/akturesult/Even2016Resul"
        "t/frmbtech4semtrererb_2016nrqiop.aspx"),
    3: ("http://results.aktu.ac.in/akturesult/Even2016Resul"
        "t/frmbtech6semhjjh_2016fbnhvg.aspx"),
    4: ("http://results.aktu.ac.in/akturesult/Even2016Resul"
        "t/frmBtech78even2016.aspx"),

}

MCA_URLS = {
    1: ('http://aktu.ac.in/results/gbturesult_11_12/Odd'
        'Semester2016/frmMCA1Sem.aspx'),
    2: ('http://aktu.ac.in/results/gbturesult_11_12/Odd'
        'Semester2016/frmMCA3Sem.aspx'),
    3: ('http://aktu.ac.in/results/gbturesult_11_12/Odd'
        'Semester2016/frmMCA5Semome.aspx'),
}

COLLEGE_CODES = ["027", "029", "032", "091", "143", "033", "161", "030",
                 "222", ]

# total students of all branches of Ajay Kumar Garg Engineering College
MAX_STUDENTS = {
    # 1st year
    '1': {
        "00": 119,
        "10": 186,
        "13": 123,
        "21": 112,
        "31": 183,
        "32": 31,
        "40": 180,
        "14": 118,
    },

    # 2nd year
    '2': {
        "00": 128,
        "10": 204,
        "13": 126,
        "21": 116,
        "31": 174,
        "32": 32,
        "40": 210,
        "14": 113,  # assumed data
    },

    # 3rd year
    '3': {
        "00": 133,
        "10": 193,
        "13": 122,
        "21": 118,
        "31": 187,
        "32": 36,
        "40": 185,
        "14": 135,
    },

    # 4th year
    '4': {
        "00": 70,
        "10": 209,
        "13": 113,
        "21": 131,
        "31": 197,
        "32": 35,
        "40": 211
    }

}

MAX_MARKS_YEARWISE = {

    '1': {
        "00": 630,
        "10": 630,
        "13": 630,
        "21": 630,
        "31": 630,
        "32": 630,
        "40": 630
    },
    '2': {
        "00": 670,
        "10": 670,
        "13": 670,
        "21": 670,
        "31": 670,
        "32": 670,
        "40": 670
    },
    '3': {
        "00": 590,
        "10": 590,
        "13": 590,
        "21": 590,
        "31": 608,
        "32": 620,
        "40": 620
    }, #620 for all branches and 590 for ME
    '4': {
        "00": 650,
        "10": 600,
        "13": 600,
        "21": 650,
        "31": 650,
        "32": 650,
        "40": 600
    },
}

WTF_CSRF_ENABLED = True
UPLOAD_FOLDER = os.getcwd() + '/Section-Faculty Information'
