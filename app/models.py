from mongokit import Connection, Document
from app import connection

class Student(Document):
    __collection__ = 'student'
    structure = {
        'roll_no': unicode,
        'name' : unicode,
        'father_name' : unicode,
        'branch' : unicode,
        'college' : unicode,
        # marks is a list whose first element is odd sem marks,
        # second element is even sem marks and third element is
        # maximum marks
        'marks' : list,
        'carry_papers': [basestring],
    }
    use_dot_notaton = True
    use_autorefs = True

    def __repr__(self):
        return self['name']

connection.register([Student])
