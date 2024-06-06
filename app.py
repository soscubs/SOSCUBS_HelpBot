from flask import Flask, render_template, request, redirect, url_for
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField
from wtforms.validators import DataRequired, Email
import openpyxl

app = Flask(__name__)
app.config['SECRET_KEY'] = 'yoursecretkey'

class InfoForm(FlaskForm):
    name = StringField('May We Know Your Name?', validators=[DataRequired()])
    contact = StringField('Please Enter Your Contact Number:', validators=[DataRequired()])
    email = StringField('Please Enter Your Email Address:', validators=[DataRequired(), Email()])
    address = StringField('Where Do You Reside?', validators=[DataRequired()])
    age = StringField('What Is Your Child\'s Age?', validators=[DataRequired()])
    submit = SubmitField('Submit')

@app.route('/', methods=['GET', 'POST'])
def index():
    form = InfoForm()
    if form.validate_on_submit():
        data = [form.name.data, form.contact.data, form.email.data, form.address.data, form.age.data]
        save_to_excel(data)
        return redirect(url_for('options'))
    return render_template('index.html', form=form)

def save_to_excel(data):
    workbook = openpyxl.load_workbook('sos_cubs_data.xlsx')
    sheet = workbook.active
    sheet.append(data)
    workbook.save('sos_cubs_data.xlsx')

@app.route('/options')
def options():
    return render_template('options.html')

if __name__ == '__main__':
    app.run(debug=True)
