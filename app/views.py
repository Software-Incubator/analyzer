from io import BytesIO
from app import app

from app.models import User
from mongokit import Connection
from .data_updater import open_excel

from .forms import *
from .excel import *
from flask import render_template, Response, request, redirect, url_for, session, flash


# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


@app.route('/', methods=['GET', 'POST'])
def index():
    form = InputForm(request.form)
    print form

    if request.method == 'POST' and form.validate_on_submit():
        fnum = request.form.get('excel')
        return redirect(url_for('excel_generator', fnum=fnum))

    return render_template("index.html", form=form)


@app.route('/excel_generator', methods=['GET', 'POST'])
def excel_generator():
    fnum = int(request.args['fnum'])  # fnum value in prev as well as that passed here is of type unicode

    if fnum:
        if fnum == 1:
            form = MainExcelForm(request.form)
            title = "Main Excel Form"

        elif fnum == 2:
            form = FailExcelForm(request.form)
            title = "Fail Excel Form"

        elif fnum == 3:
            form = AkgecForm(request.form)
            title = "Akgec Summary Form"

        elif fnum == 4:
            form = ExtAvgForm(request.form)
            title = "External Average form"

        elif fnum == 5:
            form = SecWiseForm(request.form)
            title = "Section Wise Form"

        elif fnum == 6:
            form = SubWiseForm(request.form)
            title = "Subject Wise Form"

        elif fnum == 7:
            form = PassPercentForm(request.form)
            title = "Pass Percentage Form"

        elif fnum == 8:
            form = BranchWisePassForm(request.form)
            title = "Branch Wise Pass Percentage Form"

        elif fnum == 9:
            form = BranchWiseExtForm(request.form)
            title = "Branch Wise External Avg Form"

        else:
            form = FacultyForm(request.form)
            print dir(request.form.data)
            title = "Faculty Performance Form"

        if request.method == 'POST' and form.validate_on_submit():

            output = BytesIO()
            section_file = None
            if 'college' in form:
                college = form.college.data
                print college

            if 'branch' in form:
                branch = form.branch.data

            if 'year' in form:
                years = form.year.data

            if 'file' in form:
                section_file = form.file.data

            if fnum == 1:
                make_excel(branch_codes=branch, years=years, college_code=college, output=output)
                output.seek(0)
                filename = 'main_year_' + '.xlsx'

            elif fnum == 2:
                fail_excel(college_code=college, years=years, output=output)
                output.seek(0)
                filename = 'fail_excel' + '.xlsx'

            elif fnum == 3:
                akgec_summary(years, output=output)
                output.seek(0)
                filename = 'Akgec_Summary_' + '.xlsx'

            elif fnum == 4:
                ext_avg(years, output=output)
                output.seek(0)
                filename = 'External_Average' + '.xlsx'

            elif fnum == 5:
                sec_wise_ext(years, output=output)
                output.seek(0)
                filename = 'Section_wise' + '.xlsx'

            elif fnum == 6:
                subject_wise(years=years, output=output)
                output.seek(0)
                filename = 'Subject_Wise' + '.xlsx'

            elif fnum == 7:
                pass_percentage(year_range=years, output=output)
                output.seek(0)
                filename = 'Pass_Percentage' + '.xlsx'

            elif fnum == 8:
                branch_wise_pass_percent(years=years, output=output)
                output.seek(0)
                filename = 'Branch_Wise_Pass_Percentage' + '.xlsx'

            elif fnum == 9:
                branch_wise_ext_avg(years=years, output=output)
                output.seek(0)
                filename = 'Branch_Wise_External_Average' + '.xlsx'

            else:
                faculty_performance(years=years, output=output, file=section_file)
                output.seek(0)
                filename = 'faculty_performance.xlsx'
            response = Response(output.read(),
                                content_type="application/vnd.openxmlformats-"
                                             "officedocument.spreadsheetml.sheet")

            response.headers["Content-Disposition"] = "attachment; filename=" + filename

            return response

        return render_template("excel_generate.html", form=form, title=title)


connection = Connection()
collection = connection.test.users
user = {}
user['username'] = unicode('name')
user['password'] = unicode('password')
collection.insert(user)


@app.route('/login', methods=['GET', 'POST'])
def login():
    title = 'Sign In'
    error = None
    form = LoginForm(request.form)
    if form.validate_on_submit():
        if form.username.data == user['username'] and form.password.data == user['password']:
            session['username'] = form.username.data
            flash('Successfully Logged In')

            return redirect(url_for('index'))
        else:
            error = 'Invalid Credentials'
    return render_template('login.html', form=form, title=title)
