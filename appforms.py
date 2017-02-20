from flask_wtf import Form
from wtforms import StringField, SubmitField

class FeecalcForm(Form):
    freefield = StringField("Put what you want here")
    submit = SubmitField("Send to archive")