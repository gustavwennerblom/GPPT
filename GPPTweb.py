from flask import Flask, render_template
from appforms import FeecalcForm
import getcalcs
from DBhelper import DBHelper

app = Flask(__name__)
app.secret_key = "devkey"
db = DBHelper()


@app.route("/")
def home():
    # return ("We're getting started")
    form = FeecalcForm()
    return render_template("home.html", form=form)


@app.route("/<region>")
def region(region):
    list = db.get_column_values("Date")
    return render_template("region.html", region=region)

if __name__ == "__main__":
    print("main process")
    app.run(debug=True)
