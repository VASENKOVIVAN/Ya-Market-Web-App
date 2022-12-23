from flask import Flask

app = Flask(__name__)

@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"

@app.route("/about")
def about():
    return "<p>О сайте</p>"

if __name__ == "__main__":
    app.run(debug=True)