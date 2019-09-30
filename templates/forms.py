from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import StringField, PasswordField, BooleanField, SubmitField, IntegerField
from wtforms.validators import Length, EqualTo, DataRequired, Email
from wtforms.fields.html5 import DateField


#Flask forms for key pages.


class Form_Generate(FlaskForm):
    name = StringField('Auditor Name', validators = [DataRequired()])
    email = StringField('Auditor Email', validators=[DataRequired(), Email()])
    yearend = DateField('Year End Date', format='%Y-%m-%d')
    client_name = StringField('Client Name', validators = [DataRequired()])
    client_signatory = StringField('Client Signatory', validators = [DataRequired()])
    samplesize = IntegerField('Sample Size')

    datafile = FileField('Select Data file...')
    letter_head = FileField('Select letter head (pdf) file...')
    signature = FileField('Select Signature (image/pdf) file...')


    submit = SubmitField('Generate DC Docs & Emails')
