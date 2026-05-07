from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class Reserva(db.Model):
    id           = db.Column(db.Integer, primary_key=True)
    fecha        = db.Column(db.Date, nullable=False)
    hora_inicio  = db.Column(db.String(5), nullable=False)
    hora_fin     = db.Column(db.String(5), nullable=False)
    cliente      = db.Column(db.String(100), nullable=False)
    telefono     = db.Column(db.String(20))
    monto_total  = db.Column(db.Float, nullable=False)
    monto_pagado = db.Column(db.Float, default=0)
    pagado       = db.Column(db.Boolean, default=False)
    creado_en    = db.Column(db.DateTime, default=datetime.utcnow)
    metodo_pago  = db.Column(db.String(50))
    
