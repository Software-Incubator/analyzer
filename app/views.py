from io import BytesIO
from app import app
from app import ALLOWED_EXTENSIONS, UPLOAD_FOLDER
from .forms import InputForm
from .excel import make_excel
from flask import render_template, Response

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    form = InputForm()

    if form.validate_on_submit():
        # form.college.choices = colg_choices
        # form.branch.choices = branch_choices
        # form.sem.choices = sem_choices
        colg = form.college.data[0]
        branch = form.branch.data[0]
        sem = form.sem.data[0]
        sem = int(sem)
        year = int((sem - .1) // 2 + 1)
        print 'In the views:', colg, branch, sem
        output = BytesIO()
        make_excel(branch_code=branch, year=year, college_code=colg, output=output)
        output.seek(0)
        filename = 'main_year_' + str(year) + '_branch_' + branch + '.xlsx'
        response = Response(output.read(),
                            content_type="application/vnd.openxmlformats-"
                                         "officedocument.spreadsheetml.sheet")
        response.headers["Content-Disposition"] = "attachment; filename=" + filename
        return response
    return render_template('index.html', form=form)


" Create a custom validation function to verify data as per selected excel. "

