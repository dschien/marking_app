import logging
from logging.handlers import SysLogHandler

from flask import Flask, render_template, flash, redirect, url_for, session
from flask.ext.sqlalchemy import SQLAlchemy
from flask.ext.wtf import Form
from wtforms import StringField, SelectField, IntegerField
from wtforms.validators import DataRequired, NumberRange

from marking_data import MarkingData


file_handler = SysLogHandler()
file_handler.setLevel(logging.WARNING)

app = Flask(__name__)
app.config.from_object('config')
db = SQLAlchemy(app)
app.logger.addHandler(file_handler)

marking_data = MarkingData(app)

choices = [('Uncertainty', 'Uncertainty'), ('Improvements', 'Improvements'), ('Interpretation', 'Interpretation'),
           ("Tool", "Tool"), ("Writing style", "Writing style")]


class SearchForm(Form):
    token = StringField('token', validators=[DataRequired()])
    # sections = ["Uncertainty", "Improvements", "Interpretation", "Tool", "Writing style"]
    section = SelectField(u'Section', choices=choices, validators=[DataRequired()])
    limit = IntegerField('limit', default=20, validators=[DataRequired()])


class FilterForm(Form):
    mark = IntegerField('mark', validators=[DataRequired(), NumberRange(min=0, max=12)])
    section = SelectField(u'Section', choices=choices, validators=[DataRequired()])


@app.route('/filter', methods=['GET', 'POST'])
def filter():
    form = FilterForm()
    if form.validate_on_submit():
        flash('Filtering by mark >= %i, section=%s' %
              (form.mark.data, form.section.data))

        session['mark'] = form.mark.data
        session['section'] = form.section.data

        return redirect(url_for('filter_results'))
    return render_template('filter.html',
                           title='Search',
                           form=form)


@app.route('/search', methods=['GET', 'POST'])
def search():
    form = SearchForm()
    if form.validate_on_submit():
        flash('Searching for token "%s", section=%s' %
              (form.token.data, form.section.data))

        session['token'] = form.token.data
        session['section'] = form.section.data
        session['limit'] = form.limit.data

        return redirect(url_for('results'))
    return render_template('search.html',
                           title='Search',
                           form=form)


@app.route('/results', methods=['GET', 'POST'])
def results():
    results = marking_data.search_token(session['token'], session['section'], session['limit'])
    return render_template('results.html',
                           title='Search Results',
                           results=results)


@app.route('/filter_results', methods=['GET', 'POST'])
def filter_results():
    results, marks = marking_data.filter(int(session['mark']), session['section'])
    return render_template('filter_results.html',
                           title='Filter Results',
                           results=results, marks=marks)


@app.route('/')
def index():
    return render_template('index.html',
                           title='Home',
    )


if __name__ == '__main__':
    app.run(debug=True)