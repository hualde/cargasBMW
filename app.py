from flask import Flask, render_template, request, jsonify, send_file
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import json
import os
from io import BytesIO
from datetime import datetime
import base64

app = Flask(__name__)

# Almacenamiento de datos en memoria
data_store = {
    '0': {'izquierda': 0, 'derecha': 0},
    '25': {'izquierda': 0, 'derecha': 0},
    '50': {'izquierda': 0, 'derecha': 0},
    '75': {'izquierda': 0, 'derecha': 0},
    '100': {'izquierda': 0, 'derecha': 0}
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['GET'])
def get_data():
    """Obtener todos los datos"""
    return jsonify(data_store)

@app.route('/api/data', methods=['POST'])
def update_data():
    """Actualizar datos"""
    global data_store
    try:
        new_data = request.json
        for percentage, values in new_data.items():
            if percentage in data_store:
                data_store[percentage] = {
                    'izquierda': float(values.get('izquierda', 0)),
                    'derecha': float(values.get('derecha', 0))
                }
        return jsonify({'success': True, 'data': data_store})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400

@app.route('/api/patrones', methods=['GET'])
def get_patrones():
    """Obtener todos los patrones de carga"""
    try:
        json_path = os.path.join(os.path.dirname(__file__), 'patrones_carga.json')
        with open(json_path, 'r', encoding='utf-8') as f:
            patrones = json.load(f)
        return jsonify(patrones)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/patrones/<patron_id>', methods=['GET'])
def get_patron(patron_id):
    """Obtener un patrón específico"""
    try:
        json_path = os.path.join(os.path.dirname(__file__), 'patrones_carga.json')
        with open(json_path, 'r', encoding='utf-8') as f:
            patrones = json.load(f)
        if patron_id in patrones:
            return jsonify(patrones[patron_id])
        else:
            return jsonify({'error': 'Patrón no encontrado'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generar-informe', methods=['POST'])
def generar_informe():
    """Generar un informe PDF"""
    try:
        data = request.json
        referencia = data.get('referencia', 'Sin referencia')
        cargas = data.get('cargas', {})
        referencia_bmw = data.get('referencia_bmw', None)
        dispersion = data.get('dispersion', 0)
        ruido = data.get('ruido', 0)
        imagen_grafico = data.get('imagen_grafico', None)
        
        # Crear el PDF en memoria con márgenes reducidos
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=A4,
            leftMargin=15*mm,
            rightMargin=15*mm,
            topMargin=10*mm,
            bottomMargin=15*mm
        )
        elements = []
        
        # Estilos
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#667eea'),
            spaceAfter=10,
            alignment=TA_CENTER
        )
        
        # Título
        elements.append(Paragraph(f"Informe de Cargas - {referencia}", title_style))
        elements.append(Spacer(1, 6))
        
        # Información general
        info_data = [
            ['Referencia:', referencia],
            ['Fecha:', datetime.now().strftime('%d/%m/%Y %H:%M:%S')]
        ]
        
        if referencia_bmw:
            info_data.append(['Referencia BMW:', referencia_bmw])
        if dispersion > 0:
            info_data.append(['Dispersión:', f'{dispersion}%'])
        if ruido > 0:
            info_data.append(['Ruido:', f'{ruido} mm/s²'])
        
        info_table = Table(info_data, colWidths=[60*mm, 130*mm])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.grey),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BACKGROUND', (1, 0), (1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(info_table)
        elements.append(Spacer(1, 10))
        
        # Tabla de cargas
        carga_headers = [['Porcentaje', 'Izquierda (daN)', 'Derecha (daN)']]
        carga_data = carga_headers.copy()
        
        for percent in ['0', '25', '50', '75', '100']:
            if percent in cargas:
                carga_data.append([
                    f'{percent}%',
                    f"{cargas[percent].get('izquierda', 0):.1f}",
                    f"{cargas[percent].get('derecha', 0):.1f}"
                ])
        
        carga_table = Table(carga_data, colWidths=[50*mm, 70*mm, 70*mm])
        carga_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        elements.append(Paragraph('Datos de Cargas', styles['Heading2']))
        elements.append(Spacer(1, 6))
        elements.append(carga_table)
        elements.append(Spacer(1, 10))
        
        # Añadir gráfico si está disponible
        if imagen_grafico:
            try:
                # Decodificar la imagen base64
                imagen_data = imagen_grafico.split(',')[1] if ',' in imagen_grafico else imagen_grafico
                imagen_bytes = base64.b64decode(imagen_data)
                
                # Crear objeto Image desde bytes
                img_buffer = BytesIO(imagen_bytes)
                img = Image(img_buffer, width=170*mm, height=120*mm)
                img.hAlign = 'CENTER'
                
                elements.append(Paragraph('Gráfico de Cargas', styles['Heading2']))
                elements.append(Spacer(1, 12))
                elements.append(img)
            except Exception as e:
                print(f"Error al añadir gráfico al PDF: {e}")
        
        # Construir PDF
        doc.build(elements)
        buffer.seek(0)
        
        # Nombre del archivo
        nombre_archivo = f"Informe_{referencia.replace(' ', '_')}.pdf"
        
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=nombre_archivo
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)

