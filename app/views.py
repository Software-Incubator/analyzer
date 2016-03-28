from io import BytesIO
from app import app
from .forms import InputForm
from .excel import make_excel
from flask import render_template, Response


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
        print 'In the views:', colg, branch, sem
        output = BytesIO()
        make_excel(branch=branch, sem=sem, colg_code=colg, output=output)
        output.seek(0)
        response = Response(output.read(),
                            content_type="application/vnd.openxmlformats-"
                                         "officedocument.spreadsheetml.sheet")
        response.headers["Content-Disposition"] = "attachment; filename=result.xlsx"
        return response
    return render_template('index.html', form=form)
