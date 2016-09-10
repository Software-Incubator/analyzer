from pymongo import MongoClient
from app import admin
from .models import User, UserView

client = MongoClient()
db = client.test
col = db.users



admin.add_view(UserView(col))