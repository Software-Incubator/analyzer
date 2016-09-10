from flask.ext.wtf import Form
from app import connection
from wtforms import SelectField, SelectMultipleField, StringField, PasswordField
from flask_wtf.file import FileField, FileRequired, FileAllowed
from wtforms.validators import DataRequired, Email, Length

col = connection.test.students
col.update({"college_code": 'KRISHNA INSTITUTE OF ENGINEERING AND TECHNOLOGY'}, {"$set": {"college_code": "029"}},
           upsert=False, multi=True)

colg_choices = col.distinct('college_code')
branch_choices = col.distinct('branch_choices')

# sem_choices = tuple([
#     ('1', '1'), ('2', '2'), ('3', '3'), ('4', '4'), ('5', '5'), ('6', '6'), ('7', '7'), ('8', '8')
# ])

year_choices = col.distinct('year')

excel_choices = tuple([
    ('1', 'Main Excel'), ('2', 'Fail Excel'), ('3', 'Akgec Summary'), ('4', 'External average'),
    ('5', 'Section Wise Summary'),
    ('6', 'Subject Wise'), ('7', 'Pass Percent'), ('8', 'Branch Wise Pass Percent'), ('9', 'Branch Wise ext avg'),
    ('10', 'Faculty Performance'),
])


class InputForm(Form):
    excel = SelectField('Excel', choices=excel_choices,
                        validators=[])


class MainExcelForm(Form):
    college = SelectField('College', choices=colg_choices,
                          validators=[DataRequired()])
    branch = SelectMultipleField('Branch', choices=branch_choices,
                                 validators=[DataRequired()])
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class FailExcelForm(Form):
    college = SelectField('College', choices=colg_choices,
                          validators=[DataRequired()])
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class AkgecForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class OtherColgForm(Form):
    college = SelectField('College', choices=colg_choices,
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
                     validators=[FileRequired(),
                                 FileAllowed(['xlsx', 'xls'], 'File Type Incorrect')
                                 ])


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
    username = StringField('username',
                           validators=[DataRequired()])
    password = PasswordField('password',
                             validators=[DataRequired()])

# class SignUpForm(Form):
#     username = StringField('username', validators=[DataRequired()])
#     name = StringField('name', validators=DataRequired())
#     email_id = StringField('email_id', validators=[DataRequired(),
#                            Email()])
#     password = PasswordField('password',
#                              validators=[DataRequired(),Length(min=8)])
