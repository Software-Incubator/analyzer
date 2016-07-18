# branch codes of the branches that exist in college
BRANCH_CODES = ['00', '10', '13', '20', '21', '31', '32', '40',]

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
    "21": "Electrical & Electronics Engineering",
    "20": "Electrical Engineering",
    "31": "Electronics and Communication Engineering",
    "32": "Electronics and Instrumentation Engineering",
    "40": "Mechanical Engineering",
    "14": "Master in Computer Applications",
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
    1: ("http://aktu.ac.in/results/gbturesult_11_12/Odd"
        "Semester2016/frmBtech1Sem.aspx"),
    2: ("http://aktu.ac.in/results/gbturesult_11_12/Odd"
        "Semester2016/frmBtech3Sem.aspx"),
    3: ("http://aktu.ac.in/results/gbturesult_11_12/Odd"
        "Semester2016/frmBtech5Semome.aspx"),
    4: ("http://aktu.ac.in/akturesult/Even2016Result/fr"
        "mBtech78even2016.aspx"),

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


## total students of all branches of Ajay Kumar Garg Engineering College
MAX_STUDENTS = {
    # 1st year
    '1': {
        "00": 119,
        "10": 186,
        "13": 124,
        "21": 113,
        "31": 181,
        "32": 39,
        "40": 182,
        "14": 38,
    },

    # 2nd year
    '2': {
        "00": 134,
        "10": 196,
        "13": 123,
        "21": 121,
        "31": 194,
        "32": 35,
        "40": 190,
        "14": 113,  # assumed data
    },

    # 3rd year
    '3': {
        "00": 70,
        "10": 208,
        "13": 113,
        "21": 132,
        "31": 197,
        "32": 35,
        "40": 212,
        "14": 135,
    },

    # 4th year
    '4':  {
        "00": 69,
        "10": 151,
        "13": 128,
        "21": 136,
        "31": 148,
        "32": 47,
        "40": 209
    }

}

MAX_MARKS_YEARWISE = {

    '1': 620,
    '2': 670,
    '3': 590,
    '4': 550,
}
