from mongokit import Document
from app import connection

class User(Document):
    structure = {
        'name' : unicode,
        'email' : unicode,
    }
    use_dot_notaton = True

    def __repr__(self):
        return User.name

connection.register([User])