from mongokit import Document, ValidationError
from app import connection
from flask_admin.contrib.pymongo import ModelView as md
from .forms import LoginForm


def max_length(length):
    def validate(value):
        if len(value) >= length:
            return True
        raise ValidationError("%s must be at least {} characters long".format(length))


@connection.register
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
        'aggregate_marks': unicode,
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
    use_dot_notation = True
    use_autorefs = True

    def __repr__(self):
        return self['name']


connection.register([Student])


@connection.register
class User(Document):
    __collection__ = 'user'
    structure = {
        # 'name': unicode,
        'username': unicode,
        # 'email_id': unicode,
        'password': unicode,
        # 'role': unicode,

    }

    validators = {
        'password': max_length(8),

    }

    indexes = [
        # {
        #     'fields': 'email_id',
        #     'unique': True,
        # },
        {
            'fields': 'username',
            'unique': True,
        }
    ]

    use_dot_notation = True
    use_autorefs = True

    def __repr__(self):
        return self['name']

    def is_active(self):
        return True

    def is_authenticated(self):
        return self.authenticated

    def is_anonymous(self):
        return False


connection.register([User])


class UserView(md):
    column_list = ('username', 'password')
    form = LoginForm


