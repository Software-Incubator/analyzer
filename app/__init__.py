from flask import Flask
import os
from mongokit import Connection
from werkzeug.utils import secure_filename

# configuring mongodb
MONGODB_HOST = 'localhost'
MONGODB_PORT = 27017

app = Flask(__name__)
app.config.from_object(__name__)
app.config.from_object('config')
connection = Connection(app.config['MONGODB_HOST'],
                        app.config['MONGODB_PORT'])

UPLOAD_FOLDER = os.getcwd()+'/UPLOAD'
ALLOWED_EXTENSIONS = set(['xlsx',])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'GuessItIfUCan'

from app import views, models
