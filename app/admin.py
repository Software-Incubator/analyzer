from pymongo import Connection, MongoClient
from app import admin
from .models import User, UserView

client = MongoClient('localhost', 27017)
db = client.test
col = db.users
#
# connection = Connection()
# col = connection.test.users


admin.add_view(UserView(col))