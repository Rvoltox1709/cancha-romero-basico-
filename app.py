import os
from flask import Flask
from models import db

app = Flask(__name__)

database_url = os.environ.get('DATABASE_URL')
if database_url and database_url.startswith('postgres://'):
    database_url = database_url.replace('postgres://', 'postgresql://', 1)

app.config['SQLALCHEMY_DATABASE_URI'] = database_url or 'sqlite:///cancha.db'
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'cancha-secreta-2024')

db.init_app(app)

with app.app_context():
    db.create_all()

from routes import *

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port)