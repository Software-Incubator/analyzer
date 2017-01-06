from io import BytesIO
from app import app
from app import celery

from mongokit import Connection

from .forms import *
from .excel import *
from flask import render_template, Response, request, redirect, url_for, session, flash, jsonify, g
from .tasks import get_college_results

same_session = None
response = None


@celery.task
def crawl(captcha, same_session, response):
    get_college_results(session=same_session, response=response, captcha=captcha,
                        college_codes=app.config["COLLEGE_CODES"],
                        year=session['year'])


@app.route('/', methods=['GET', 'POST'])
def index():
    form = InputForm(request.form)

    if request.method == 'POST' and form.validate_on_submit():
        fnum = request.form.get('excel')
        return redirect(url_for('excel_generator', fnum=fnum))

    return render_template("index.html", form=form)


# @app.route('/get_captcha', methods=['GET', 'POST'])
# def get_captcha():
#     if request.is_xhr:
#         year = int(request.form['year'])
#         session['year'] = year
#         data = get_crawl_data(year=year)  will need to define get_crawl_data in tasks.py to get info
#         same_session = data[0]
#         response = data[1]
#         return jsonify(imgurl=data[2])


@app.route('/get_data', methods=['GET', 'POST'])
def get_crawl_info():
    """
    """
    form = CrawlForm(request.form)

    # if request.is_xhr:
    #     year = int(request.form['year'])
    #     session['year'] = year
    #     data = get_college_results(college_codes=app.config["COLLEGE_CODES"], year=year, in_view=True)
    #
    #     same_session = data[1]
    #     return jsonify(imgurl=data[1])

    if request.method == 'POST' and request.form['captcha']:
        captcha = str(request.form['captcha'])

        task = crawl.apply_async(args=[captcha, same_session, response])

    return render_template('get_data.html', form=form)


@app.route('/excel_generator', methods=['GET', 'POST'])
def excel_generator():
    fnum = int(request.args['fnum'])  # fnum value in prev as well as that passed here is of type unicode
    if fnum == 1:
        form = MainExcelForm(request.form)
        title = "Main Excel Form"

    elif fnum == 2:
        form = FailExcelForm(request.form)
        title = "Fail Excel Form"

    elif fnum == 3:
        form = FacultyForm(request.form)
        title = "Faculty Performance Form"

    else:
        form = YearForm(request.form)
        if fnum == 4:
            title = "External Average form"

        elif fnum == 5:
            title = "Section Wise Form"

        elif fnum == 6:
            title = "Subject Wise Form"

        elif fnum == 7:
            title = "Pass Percentage Form"

        elif fnum == 8:
            title = "Branch Wise Pass Percentage Form"

        elif fnum == 9:
            title = "Branch Wise External Avg Form"

        else:
            title = "Akgec Summary Form"

    if request.method == 'POST' and form.validate_on_submit():
        output = BytesIO()
        section_file = None
        if 'college' in form:
            college = form.college.data

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
            faculty_performance(years=years, output=output, file=section_file)
            output.seek(0)
            filename = 'faculty_performance.xlsx'

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
            akgec_summary(years, output=output)
            output.seek(0)
            filename = 'Akgec_Summary_' + '.xlsx'

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
