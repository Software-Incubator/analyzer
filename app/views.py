from io import BytesIO
from app import app
# from app import ALLOWED_EXTENSIONS, UPLOAD_FOLDER
from .data_updater import open_excel
from .forms import *
from .excel import *
from flask import render_template, Response, flash, request, redirect, url_for, session


# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    form = InputForm(request.form)
    # print request.form['excel']
    # print "entered" right
    print 'Method: ', request.method
    print form.validate_on_submit()  # gives false why?
    print form.csrf_token
    if request.method == 'POST' and not form.validate_on_submit():
        fnum = request.form.get('excel')
        print fnum
        return redirect(url_for('excel_generator'), request.form.get('excel'))

    return render_template("index.html", form=form)


@app.route('/excel_generator')
def excel_generator(fnum):
    if fnum:
        if fnum == 1:
            form = MainExcelForm(request.form)

        elif fnum == 2:
            form = FailExcelForm(request.form)

        elif fnum == 3:
            form = AkgecForm(request.form)

        elif fnum == 4:
            form = ExtAvgForm(request.form)

        elif fnum == 5:
            form = SecWiseForm(request.form)

        elif fnum == 6:
            form = SubWiseForm(request.form)

        elif fnum == 7:
            form = PassPercentForm(request.form)

        elif fnum == 8:
            form = BranchWisePassForm(request.form)

        elif fnum == 9:
            form = BranchWiseExtForm(request.form)

        else:
            form = FacultyForm(request.form)

        if request.method == 'POST' and form.validate():
            output = BytesIO()
            colg = form.college.data
            branch = form.branch.data
            year = form.year.data
            section_file = form.file

            if fnum == 1 or fnum == 2:
                for college in colg:
                    for year in year:
                        if fnum == 1:
                            for branch in branch:
                                make_excel(branch_code=branch, year=year, college_code=colg, output=output)
                                output.seek(0)
                                filename = 'main_year_' + str(year) + '_branch_' + branch + '.xlsx'
                                response = Response(output.read(),
                                                    content_type="application/vnd.openxmlformats-"
                                                                 "officedocument.spreadsheetml.sheet")
                                response.headers["Content-Disposition"] = "attachment; filename=" + filename

                        else:
                            fail_excel(college_code=colg, year=year, output=output)
                            output.seek(0)
                            filename = 'main_year_' + str(year) + '_branch_' + branch + '.xlsx'
                            response = Response(output.read(),
                                                content_type="application/vnd.openxmlformats-"
                                                             "officedocument.spreadsheetml.sheet")
                            response.headers["Content-Disposition"] = "attachment; filename=" + filename

            else:
                for year in year:
                    if fnum == 3:
                        akgec_summary(year)

                    elif fnum == 4:
                        ext_avg(year)

                    elif fnum == 5:
                        sec_wise_ext(year)

                    elif fnum == 6:
                        subject_wise(year)

                    elif fnum == 7:
                        pass_percentage(year_range=range(1, year))

                    elif fnum == 8:
                        branch_wise_pass_percent(year=year)

                    elif fnum == 9:
                        branch_wise_ext_avg(year=year)

                    else:
                        open_excel()
                        faculty_performance(year)


            return response

        return render_template("excel_generate.html", form=form)

        # elif fnum == 2:
        #     fail_excel
        #




        #     colg = form.college.data
        #     branch = form.branch.data
        #     sem = form.year.data[0]
        #     sem = int(sem)
        #     year = int((sem - .1) // 2 + 1)
        #     print 'In the views:', colg, branch, sem
        #     output = BytesIO()
        #     for college in colg:
        #         for year in year:
        #             for branch in branch:
        #                 make_excel(branch_code=branch, year=year, college_code=colg, output=output)
        #                 output.seek(0)
        #                 filename = 'main_year_' + str(year) + '_branch_' + branch + '.xlsx'
        #                 response = Response(output.read(),
        #                                     content_type="application/vnd.openxmlformats-"
        #                                                  "officedocument.spreadsheetml.sheet")
        #                 response.headers["Content-Disposition"] = "attachment; filename=" + filename
        #                 return response
        #
        # return render_template('excel_generate.html', form=form)
        # " Create a custom validation function to verify data as per selected excel. "
