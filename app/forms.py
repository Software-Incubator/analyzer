from flask.ext.wtf import Form
from app import connection
from wtforms import SelectField, SelectMultipleField, StringField, PasswordField, FileField
from flask_wtf.file import FileField, FileRequired, FileAllowed
from wtforms.validators import DataRequired

col = connection.test.students
col.update({"college_code": 'KRISHNA INSTITUTE OF ENGINEERING AND TECHNOLOGY'}, {"$set": {"college_code": "029"}},
           upsert=False, multi=True)

colg_codes = col.distinct('college_code')
colg_names = col.distinct('college_name')
branch_codes = col.distinct('branch_code')
branch_names = col.distinct('branch_name')
colg_choices = tuple(zip(colg_codes, colg_names))
branch_choices = tuple(zip(branch_codes, branch_names))

year_choices = tuple((str(x), str(x)) for x in col.distinct('year'))

excel_choices = tuple([
    ('1', 'Main Excel'), ('2', 'Fail Excel'), ('3', 'Faculty Performance'), ('4', 'External average'),
    ('5', 'Section Wise Summary'),
    ('6', 'Subject Wise'), ('7', 'Pass Percent'), ('8', 'Branch Wise Pass Percent'), ('9', 'Branch Wise ext avg'),
    ('10', 'Akgec Summary'),
])


class InputForm(Form):
    excel = SelectField('Excel', choices=excel_choices,
                        validators=[DataRequired()])


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


class FacultyForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])
    file = FileField("Upload excel of faculty information",
                     validators=[FileRequired(),
                                 FileAllowed(['xlsx', 'xls'], 'File Type Incorrect')
                                 ])


class YearForm(Form):
    year = SelectMultipleField("Year", choices=year_choices,
                               validators=[DataRequired()])


class CrawlForm(Form):
    year = SelectField("Select Year", choices=year_choices,
                       validators=[DataRequired()])


class LoginForm(Form):
    username = StringField('username',
                           validators=[DataRequired()])
    password = PasswordField('password',
                             validators=[DataRequired()])
