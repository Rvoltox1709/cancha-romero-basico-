from flask import Flask
from models import db

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///cancha.db'
app.config['SECRET_KEY'] = 'cancha-secreta-2024'

db.init_app(app)

with app.app_context():
    db.create_all()

from routes import *

if __name__ == '__main__':
    app.run(debug=True)