from flask import request, jsonify
from app import app, db
from models import Reserva
from datetime import datetime, date

# --- VER RESERVAS DE UN DÍA ---
@app.route('/reservas', methods=['GET'])
def ver_reservas():
    fecha_str = request.args.get('fecha')
    if fecha_str:
        fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date()
    else:
        fecha = date.today()
    
    reservas = Reserva.query.filter_by(fecha=fecha).order_by(Reserva.hora_inicio).all()
    
    resultado = []
    for r in reservas:
        resultado.append({
            'id': r.id,
            'fecha': str(r.fecha),
            'hora_inicio': r.hora_inicio,
            'hora_fin': r.hora_fin,
            'cliente': r.cliente,
            'telefono': r.telefono,
            'monto_total': r.monto_total,
            'monto_pagado': r.monto_pagado,
            'pagado': r.pagado
        })
    
    return jsonify(resultado)


# --- AGREGAR RESERVA ---
@app.route('/reservas', methods=['POST'])
def agregar_reserva():
    datos = request.get_json()
    
    nueva = Reserva(
        fecha        = datetime.strptime(datos['fecha'], '%Y-%m-%d').date(),
        hora_inicio  = datos['hora_inicio'],
        hora_fin     = datos['hora_fin'],
        cliente      = datos['cliente'],
        telefono     = datos.get('telefono', ''),
        monto_total  = datos['monto_total'],
        monto_pagado = datos.get('monto_pagado', 0),
        pagado       = datos.get('pagado', False)
    )
    
    db.session.add(nueva)
    db.session.commit()
    
    return jsonify({'mensaje': 'Reserva creada', 'id': nueva.id}), 201


# --- EDITAR RESERVA ---
@app.route('/reservas/<int:id>', methods=['PUT'])
def editar_reserva(id):
    reserva = Reserva.query.get_or_404(id)
    datos = request.get_json()
    
    reserva.cliente      = datos.get('cliente', reserva.cliente)
    reserva.telefono     = datos.get('telefono', reserva.telefono)
    reserva.hora_inicio  = datos.get('hora_inicio', reserva.hora_inicio)
    reserva.hora_fin     = datos.get('hora_fin', reserva.hora_fin)
    reserva.monto_total  = datos.get('monto_total', reserva.monto_total)
    reserva.monto_pagado = datos.get('monto_pagado', reserva.monto_pagado)
    reserva.pagado       = datos.get('pagado', reserva.pagado)
    
    db.session.commit()
    
    return jsonify({'mensaje': 'Reserva actualizada'})


# --- ELIMINAR RESERVA ---
@app.route('/reservas/<int:id>', methods=['DELETE'])
def eliminar_reserva(id):
    reserva = Reserva.query.get_or_404(id)
    db.session.delete(reserva)
    db.session.commit()
    return jsonify({'mensaje': 'Reserva eliminada'})


# --- RESUMEN DEL DÍA ---
@app.route('/resumen/dia', methods=['GET'])
def resumen_dia():
    fecha_str = request.args.get('fecha')
    if fecha_str:
        fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date()
    else:
        fecha = date.today()
    
    reservas = Reserva.query.filter_by(fecha=fecha).all()
    
    total_reservas  = len(reservas)
    total_esperado  = sum(r.monto_total for r in reservas)
    total_cobrado   = sum(r.monto_pagado for r in reservas)
    total_pendiente = total_esperado - total_cobrado
    
    return jsonify({
        'fecha': str(fecha),
        'total_reservas': total_reservas,
        'total_esperado': total_esperado,
        'total_cobrado': total_cobrado,
        'total_pendiente': total_pendiente
    })


# --- RESUMEN DEL MES ---
@app.route('/resumen/mes', methods=['GET'])
def resumen_mes():
    anio = int(request.args.get('anio', date.today().year))
    mes  = int(request.args.get('mes', date.today().month))
    
    reservas = Reserva.query.filter(
        db.extract('year', Reserva.fecha) == anio,
        db.extract('month', Reserva.fecha) == mes
    ).all()
    
    total_reservas  = len(reservas)
    total_esperado  = sum(r.monto_total for r in reservas)
    total_cobrado   = sum(r.monto_pagado for r in reservas)
    total_pendiente = total_esperado - total_cobrado
    
    return jsonify({
        'anio': anio,
        'mes': mes,
        'total_reservas': total_reservas,
        'total_esperado': total_esperado,
        'total_cobrado': total_cobrado,
        'total_pendiente': total_pendiente
    })

from flask import render_template

@app.route('/')
def inicio():
    return render_template('index.html')

from flask import send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import io

@app.route('/exportar/excel', methods=['GET'])
def exportar_excel():
    mes  = int(request.args.get('mes', date.today().month))
    anio = int(request.args.get('anio', date.today().year))

    reservas = Reserva.query.filter(
        db.extract('year', Reserva.fecha) == anio,
        db.extract('month', Reserva.fecha) == mes
    ).order_by(Reserva.fecha, Reserva.hora_inicio).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Reservas {mes}-{anio}"

    # Encabezado principal
    ws.merge_cells('A1:I1')
    ws['A1'] = f'Reservas Cancha - {mes}/{anio}'
    ws['A1'].font = Font(bold=True, size=14, color='FFFFFF')
    ws['A1'].fill = PatternFill('solid', fgColor='1a3a1a')
    ws['A1'].alignment = Alignment(horizontal='center')

    # Cabeceras de columnas
    cabeceras = ['Fecha', 'Hora inicio', 'Hora fin', 'Cliente', 'Teléfono', 'Monto total', 'Monto pagado', 'Pendiente', 'Estado']
    for col, cab in enumerate(cabeceras, 1):
        celda = ws.cell(row=2, column=col, value=cab)
        celda.font = Font(bold=True, color='FFFFFF')
        celda.fill = PatternFill('solid', fgColor='2d6a2d')
        celda.alignment = Alignment(horizontal='center')

    # Datos
    for fila, r in enumerate(reservas, 3):
        pendiente = r.monto_total - r.monto_pagado
        ws.cell(row=fila, column=1, value=str(r.fecha))
        ws.cell(row=fila, column=2, value=r.hora_inicio)
        ws.cell(row=fila, column=3, value=r.hora_fin)
        ws.cell(row=fila, column=4, value=r.cliente)
        ws.cell(row=fila, column=5, value=r.telefono)
        ws.cell(row=fila, column=6, value=r.monto_total)
        ws.cell(row=fila, column=7, value=r.monto_pagado)
        ws.cell(row=fila, column=8, value=pendiente)
        ws.cell(row=fila, column=9, value='Pagado' if r.pagado else 'Pendiente')

        # Color por estado
        color = 'e8f5e1' if r.pagado else 'fff3cd'
        for col in range(1, 10):
            ws.cell(row=fila, column=col).fill = PatternFill('solid', fgColor=color)

    # Fila de totales
    fila_total = len(reservas) + 3
    ws.cell(row=fila_total, column=3, value='TOTAL').font = Font(bold=True)
    ws.cell(row=fila_total, column=6, value=sum(r.monto_total for r in reservas)).font = Font(bold=True)
    ws.cell(row=fila_total, column=7, value=sum(r.monto_pagado for r in reservas)).font = Font(bold=True)
    ws.cell(row=fila_total, column=8, value=sum(r.monto_total - r.monto_pagado for r in reservas)).font = Font(bold=True)

    # Ancho de columnas
    anchos = [12, 12, 10, 20, 15, 13, 14, 12, 12]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho

    # Guardar en memoria y enviar
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'cancha_reservas_{mes}_{anio}.xlsx'
    )

from models import Reserva, Transaccion

@app.route('/transacciones', methods=['GET'])
def ver_transacciones():
    fecha_str = request.args.get('fecha')
    fecha = datetime.strptime(fecha_str, '%Y-%m-%d').date() if fecha_str else date.today()
    transacciones = Transaccion.query.filter_by(fecha=fecha).order_by(Transaccion.creado_en.desc()).all()
    return jsonify([{
        'id': t.id,
        'metodo': t.metodo,
        'monto': t.monto,
        'descripcion': t.descripcion,
        'fecha': str(t.fecha)
    } for t in transacciones])

@app.route('/transacciones', methods=['POST'])
def agregar_transaccion():
    datos = request.get_json()
    nueva = Transaccion(
        fecha       = datetime.strptime(datos['fecha'], '%Y-%m-%d').date(),
        metodo      = datos['metodo'],
        monto       = datos['monto'],
        descripcion = datos.get('descripcion', '')
    )
    db.session.add(nueva)
    db.session.commit()
    return jsonify({'mensaje': 'Transaccion registrada'}), 201