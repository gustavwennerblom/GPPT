from flask import Flask, render_template
from appforms import FeecalcForm



app = Flask(__name__)
app.secret_key="devkey"

@app.route("/")
def home():
    # return ("We're getting started")
    form = FeecalcForm()
    return render_template("home.html", form=form)

if __name__ == "__main__":
    print("main process")
    app.run(debug=True)
