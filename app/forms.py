from flask.ext.wtf import Form

from wtforms import SelectMultipleField
from wtforms.validators import DataRequired

colg_choices = tuple([('027', 'Ajay Kumar Garg Engineering College'),
                ('029', 'Krishna Institute of Technology'),
                ('091', 'JSS Academy of Technical Education'),
                ( '290', 'ABES Institute of Technology'),
                ])

branch_choices = tuple([('10', 'Computer Science & Engineering'),
                  ('31', 'Electronics & Communication Engineering	'),
                  ('40', 'Mechanical Engineering'),
                  ('21', 'Electrical & Electronics Engineering'),
                  ('13', 'Information Technology'),
                  ('00','Civil Engineering'),
                  ('32', 'Electronics and Instrumentation'),
                  ])

sem_choices = tuple([
    ('1', '1'), ('2', '2'), ('3', '3'), ('4', '4'), ('5', '5'), ('6', '6'), ('7', '7'), ('8', '8')
    ])

class InputForm(Form):
    college = SelectMultipleField('College', choices= colg_choices)
    branch = SelectMultipleField('Branch', choices= branch_choices)
    sem = SelectMultipleField("Semester", choices= sem_choices)