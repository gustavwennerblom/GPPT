from flask_wtf import Form
from wtforms import StringField

class FeecalcForm(Form):
    freefield = StringField("Put what youw want here")