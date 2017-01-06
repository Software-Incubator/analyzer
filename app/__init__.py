from flask import Flask
import os

from mongokit import Connection
from flask.ext.login import LoginManager
from flask_admin import Admin
from celery import Celery

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


def make_celery(app):
    celery = Celery(app.import_name, backend=app.config['CELERY_BACKEND'],
                    broker=app.config['CELERY_BROKER_URL'])
    celery.conf.update(app.config)
    TaskBase = celery.Task

    class ContextTask(TaskBase):
        abstract = True

        def __call__(self, *args, **kwargs):
            with app.app_context():
                return TaskBase.__call__(self, *args, **kwargs)

    celery.Task = ContextTask
    return celery


app.config.update(
    CELERY_BROKER_URL='amqp://guest@localhost//',
    CELERY_RESULT_BACKEND='amqp://guest@localhost//',
    CELERY_BACKEND = 'amqp://guest@localhost//')

celery = make_celery(app)

from app import views, models
