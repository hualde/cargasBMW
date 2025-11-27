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
import pandas as pd
import re

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

@app.route('/carga-masiva')
def carga_masiva():
    """Vista para carga masiva de datos desde Excel"""
    return render_template('carga_masiva.html')

@app.route('/api/procesar-excel', methods=['POST'])
def procesar_excel():
    """Procesar archivo Excel y extraer datos de cargas"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No se proporcionó ningún archivo'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No se seleccionó ningún archivo'}), 400
        
        # Leer el Excel
        df = pd.read_excel(file, header=None)
        
        # Extraer referencia del título (fila 0, columna 0)
        referencia_excel = None
        if len(df) > 0 and pd.notna(df.iloc[0, 0]):
            titulo = str(df.iloc[0, 0])
            # Buscar patrón de referencia en el título
            match = re.search(r'Referencia\s+([a-zA-Z0-9]+)', titulo)
            if match:
                referencia_excel = match.group(1)
        
        # Los encabezados están en la fila 2 (índice 2)
        # Datos empiezan en la fila 3 (índice 3)
        if len(df) < 3:
            return jsonify({'error': 'El archivo Excel no tiene el formato esperado'}), 400
        
        # Mapeo de NumPaso a porcentaje: 0=0%, 1=25%, 2=50%, 3=75%, 4=100%
        # Excluir NumPaso 5 según requerimiento
        paso_to_percent = {0: '0', 1: '25', 2: '50', 3: '75', 4: '100'}
        
        # Procesar datos agrupados por Pieza y NumPaso
        # Usar un diccionario para acumular valores y contar ocurrencias
        datos_acumulados = {}
        
        # Iterar desde la fila 3 (índice 3) en adelante
        for idx in range(3, len(df)):
            row = df.iloc[idx]
            
            # Verificar que la fila tenga datos válidos
            if pd.isna(row.iloc[1]):  # Columna Pieza
                continue
            
            try:
                pieza = int(row.iloc[1])  # Columna 1: Pieza
                of = int(row.iloc[3]) if pd.notna(row.iloc[3]) else None  # Columna 3: OF (Orden de Fabricación)
                num_paso = int(row.iloc[4]) if pd.notna(row.iloc[4]) else None  # Columna 4: NumPaso
                
                # Excluir filas con NumPaso = 5
                if num_paso == 5:
                    continue
                
                # Validar que OF y num_paso sean válidos
                if of is None or num_paso is None:
                    continue
                
                carga_izda = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0  # Columna 5: CargaIZDA
                carga_drch = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0  # Columna 6: CargaDRCH
                
                if num_paso not in paso_to_percent:
                    continue
                
                percent = paso_to_percent[num_paso]
                # Clave incluye OF y Pieza para diferenciar piezas del mismo número pero de distintas OF
                key = (of, pieza, percent)
                
                # Acumular valores para calcular promedio
                if key not in datos_acumulados:
                    datos_acumulados[key] = {
                        'izquierda': [],
                        'derecha': []
                    }
                
                datos_acumulados[key]['izquierda'].append(carga_izda)
                datos_acumulados[key]['derecha'].append(carga_drch)
                
            except (ValueError, IndexError) as e:
                continue
        
        # Calcular promedios y organizar por OF y Pieza
        datos_por_pieza = {}
        for (of, pieza, percent), valores in datos_acumulados.items():
            # Clave única que combina OF y Pieza
            clave_pieza = f'OF{of}_Pieza{pieza}'
            
            if clave_pieza not in datos_por_pieza:
                datos_por_pieza[clave_pieza] = {
                    'referencia': f'Pieza {pieza} - OF {of}',
                    'of': of,
                    'pieza': pieza,
                    'cargas': {
                        '0': {'izquierda': 0, 'derecha': 0},
                        '25': {'izquierda': 0, 'derecha': 0},
                        '50': {'izquierda': 0, 'derecha': 0},
                        '75': {'izquierda': 0, 'derecha': 0},
                        '100': {'izquierda': 0, 'derecha': 0}
                    }
                }
            
            # Calcular promedio
            avg_izda = sum(valores['izquierda']) / len(valores['izquierda']) if valores['izquierda'] else 0
            avg_drch = sum(valores['derecha']) / len(valores['derecha']) if valores['derecha'] else 0
            
            datos_por_pieza[clave_pieza]['cargas'][percent]['izquierda'] = avg_izda
            datos_por_pieza[clave_pieza]['cargas'][percent]['derecha'] = avg_drch
        
        # Convertir a formato JSON
        resultado = {
            'referencia_excel': referencia_excel,
            'piezas': {}
        }
        
        for pieza, datos in sorted(datos_por_pieza.items()):
            resultado['piezas'][str(pieza)] = datos
        
        return jsonify(resultado)
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error al procesar el archivo: {str(e)}'}), 500

@app.route('/api/generar-informe-masivo', methods=['POST'])
def generar_informe_masivo():
    """Generar un informe PDF para carga masiva con múltiples piezas"""
    try:
        data = request.json
        referencia_excel = data.get('referencia_excel', None)
        piezas = data.get('piezas', {})
        referencia_bmw = data.get('referencia_bmw', None)
        dispersion = data.get('dispersion', 0)
        imagen_grafico = data.get('imagen_grafico', None)
        
        if not piezas:
            return jsonify({'error': 'No hay piezas para generar el informe'}), 400
        
        # Crear el PDF en memoria
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
            fontSize=18,
            textColor=colors.HexColor('#667eea'),
            spaceAfter=10,
            alignment=TA_CENTER
        )
        
        # Título
        titulo = "Informe de Cargas Masivas"
        if referencia_excel:
            titulo += f" - {referencia_excel}"
        elements.append(Paragraph(titulo, title_style))
        elements.append(Spacer(1, 6))
        
        # Información general
        info_data = [
            ['Fecha:', datetime.now().strftime('%d/%m/%Y %H:%M:%S')],
            ['Número de piezas:', str(len(piezas))]
        ]
        
        if referencia_excel:
            info_data.append(['Referencia Excel:', referencia_excel])
        if referencia_bmw:
            info_data.append(['Referencia BMW:', referencia_bmw])
        if dispersion > 0:
            info_data.append(['Dispersión:', f'{dispersion}%'])
        
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
        
        # Tabla resumen de todas las piezas
        resumen_headers = [['Pieza', '0%', '25%', '50%', '75%', '100%']]
        resumen_data = resumen_headers.copy()
        
        # Ordenar piezas por OF y número de pieza
        piezas_ordenadas = sorted(piezas.items(), key=lambda x: (
            x[1].get('of', 0),
            x[1].get('pieza', 0)
        ))
        
        for pieza_id, pieza_info in piezas_ordenadas:
            referencia = pieza_info.get('referencia', pieza_id)
            cargas = pieza_info.get('cargas', {})
            
            # Calcular valores promedio para cada porcentaje
            valores = []
            for percent in ['0', '25', '50', '75', '100']:
                if percent in cargas:
                    izda = cargas[percent].get('izquierda', 0)
                    drch = cargas[percent].get('derecha', 0)
                    promedio = (abs(izda) + abs(drch)) / 2
                    valores.append(f'{promedio:.1f}')
                else:
                    valores.append('-')
            
            resumen_data.append([referencia] + valores)
        
        resumen_table = Table(resumen_data, colWidths=[50*mm, 25*mm, 25*mm, 25*mm, 25*mm, 25*mm])
        resumen_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        elements.append(Paragraph('Resumen de Cargas por Pieza', styles['Heading2']))
        elements.append(Spacer(1, 6))
        elements.append(resumen_table)
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
        nombre_archivo = f"Informe_Masivo_{referencia_excel.replace(' ', '_') if referencia_excel else datetime.now().strftime('%Y%m%d')}.pdf"
        
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

