from flask import Flask
import os

from mongokit import Connection
from flask.ext.login import LoginManager
from flask_admin import Admin

# configuring mongodb
MONGODB_HOST = 'localhost'
MONGODB_PORT = 27017

app = Flask(__name__)
app.config.from_object(__name__)
app.config.from_object('config')
connection = Connection(app.config['MONGODB_HOST'],
                        app.config['MONGODB_PORT'])

app.secret_key = 'GuessItIfUCan'

lm = LoginManager()
lm.init_app(app)
admin = Admin()
admin.init_app(app)

from app import views, models
