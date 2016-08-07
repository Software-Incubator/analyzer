from mongokit import Document
from app import connection


class Student(Document):
    __collection__ = 'student'
    structure = {
        'roll_no': unicode,
        'name': unicode,
        'father_name': unicode,
        'branch_code': unicode,
        'branch_name': unicode,
        'college_code': unicode,
        'college_name': unicode,
        'marks': dict,
        'max_marks': unicode,
        'aggregate': unicode,
        'carry_papers': [basestring],
        'year': unicode,
        'carry_status': unicode,
        'section': unicode,
    }

    indexes = [
        {
            'fields': 'roll_no',
            'unique': True,
        }
    ]
    use_dot_notaton = True
    use_autorefs = True

    def __repr__(self):
        return self['name']



connection.register([Student])
