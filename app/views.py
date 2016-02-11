from app import app
from .forms import InputForm
from .excel import make_excel
from flask import render_template


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
        print 'views' , colg, branch, sem
        make_excel(colg_code=colg, branch=branch, sem=sem )
    return render_template('index.html', form = form)

