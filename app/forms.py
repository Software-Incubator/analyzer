from flask.ext.wtf import Form

from wtforms import SelectField, SelectMultipleField, StringField, PasswordField
from flask_wtf.file import FileField, FileAllowed
from wtforms.validators import DataRequired

colg_choices = tuple([('027', 'Ajay Kumar Garg Engineering College'),
                      ('029', 'Krishna Institute of Technology'),
                      ('091', 'JSS Academy of Technical Education'),
                      ('290', 'ABES Institute of Technology'),
                      ])

branch_choices = tuple([('10', 'Computer Science & Engineering'),
                        ('31', 'Electronics & Communication Engineering	'),
                        ('40', 'Mechanical Engineering'),
                        ('21', 'Electrical & Electronics Engineering'),
                        ('13', 'Information Technology'),
                        ('00', 'Civil Engineering'),
                        ('32', 'Electronics and Instrumentation'),
                        ])

# sem_choices = tuple([
#     ('1', '1'), ('2', '2'), ('3', '3'), ('4', '4'), ('5', '5'), ('6', '6'), ('7', '7'), ('8', '8')
# ])


#
year_choices = tuple([
    ('1', '1'), ('2', '2'), ('3', '3'), ('4', '4'),
])

excel_choices = tuple([
    (1, 'Main Excel'), (2, 'Fail Excel'), (3, 'Akgec Summary'), (4, 'External average'), (5, 'Section Wise Summary'),
    (6, 'Subject Wise'), (7, 'Pass Percent'), (8, 'Branch Wise Pass Percent'), (9, 'Branch Wise ext avg'),
    (10, 'Faculty Performance'),
])


class InputForm(Form):
    excel = SelectField('Excel', choices=excel_choices,
                        validators=[])


class MainExcelForm(Form):
    college = SelectMultipleField('College', choices=colg_choices,
                                  validators=[DataRequired()])
    branch = SelectMultipleField('Branch', choices=branch_choices,
                                 validators=[DataRequired()])
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class FailExcelForm(Form):
    college = SelectMultipleField('College', choices=colg_choices,
                                  validators=[DataRequired()])
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class AkgecForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class OtherColgForm(Form):
    college = SelectMultipleField('College', choices=colg_choices,
                                  validators=[DataRequired()])


class ExtAvgForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class SecWiseForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class FacultyForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])
    file = FileField("Upload excel of faculty information",
                     validators=[DataRequired(),
                                 FileAllowed('xlsx', )])


class SubWiseForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class PassPercentForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class BranchWisePassForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])

class BranchWiseExtForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class LoginForm(Form):
    username = StringField('username')
    password = PasswordField('password')
