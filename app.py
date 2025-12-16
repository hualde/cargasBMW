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
    """Procesar archivo Excel y extraer datos de par (Nm)"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No se proporcionó ningún archivo'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No se seleccionó ningún archivo'}), 400
        
        # Leer solo la pestaña "Análisis de test"
        try:
            df = pd.read_excel(file, sheet_name='Análisis de test', header=None)
        except ValueError:
            # Si no existe la pestaña, intentar leer la primera
            file.seek(0)  # Resetear el archivo
            df = pd.read_excel(file, header=None)
            return jsonify({'error': 'No se encontró la pestaña "Análisis de test"'}), 400
        
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
        
        # Mapeo de NumPaso a porcentaje: solo pasos 0 y 4 (0% y 100%)
        paso_to_percent = {0: '0', 4: '100'}
        
        # Procesar datos agrupados por Pieza y NumPaso
        # Usar un diccionario para acumular valores y contar ocurrencias
        datos_acumulados = {}
        # Diccionario separado para consumos del paso 5 (solo para el 100%)
        consumos_paso5 = {}
        
        # Iterar desde la fila 3 (índice 3) en adelante
        # Detener cuando encontremos líneas en blanco (todas las columnas relevantes vacías)
        for idx in range(3, len(df)):
            row = df.iloc[idx]
            
            # Verificar si la fila está completamente vacía o tiene líneas en blanco
            # Verificar columnas clave: Pieza (columna B, índice 1), OF (columna D, índice 3), NumPaso (columna E, índice 4)
            # valores de par (columnas H e I, índices 7 y 8) y consumo (columnas J e K, índices 9 y 10)
            # Si todas las columnas relevantes están vacías, detener el procesamiento
            try:
                if (pd.isna(row.iloc[1]) and pd.isna(row.iloc[3]) and pd.isna(row.iloc[4]) and 
                    pd.isna(row.iloc[7]) and pd.isna(row.iloc[8]) and 
                    pd.isna(row.iloc[9]) and pd.isna(row.iloc[10])):
                    # Si todas las columnas relevantes están vacías, detener el procesamiento
                    break
            except IndexError:
                # Si no hay suficientes columnas, detener el procesamiento
                break
            
            # Verificar que la fila tenga datos válidos en Pieza
            try:
                if pd.isna(row.iloc[1]):  # Columna Pieza
                    continue
            except IndexError:
                continue
            
            try:
                pieza = int(row.iloc[1])  # Columna B (índice 1): Pieza
                of = int(row.iloc[3]) if pd.notna(row.iloc[3]) else None  # Columna D (índice 3): OF (Orden de Fabricación)
                num_paso = int(row.iloc[4]) if pd.notna(row.iloc[4]) else None  # Columna E (índice 4): NumPaso
                
                # Validar que OF y num_paso sean válidos
                if of is None or num_paso is None:
                    continue
                
                # Columnas J e K (índices 9 y 10) contienen los valores de consumo en amperios
                consumo_izda = round(float(row.iloc[9]), 2) if pd.notna(row.iloc[9]) else 0  # Columna J (índice 9): Consumo Izquierda (A)
                consumo_drch = round(float(row.iloc[10]), 2) if pd.notna(row.iloc[10]) else 0  # Columna K (índice 10): Consumo Derecha (A)
                
                # Si es NumPaso 5, guardar solo los consumos para el 100%
                if num_paso == 5:
                    key_paso5 = (of, pieza)
                    consumos_paso5[key_paso5] = {
                        'consumo_izquierda': consumo_izda,
                        'consumo_derecha': consumo_drch
                    }
                    continue
                
                # Para pasos 0 y 4, procesar normalmente
                if num_paso not in paso_to_percent:
                    continue
                
                # Columnas H e I (índices 7 y 8) contienen los valores de par en Nm
                # Redondear a 1 decimal para mantener la precisión del Excel
                par_izda = round(float(row.iloc[7]), 1) if pd.notna(row.iloc[7]) else 0  # Columna H (índice 7): Par Izquierda
                par_drch = round(float(row.iloc[8]), 1) if pd.notna(row.iloc[8]) else 0  # Columna I (índice 8): Par Derecha
                
                percent = paso_to_percent[num_paso]
                # Clave incluye OF y Pieza para diferenciar piezas del mismo número pero de distintas OF
                key = (of, pieza, percent)
                
                # Si hay múltiples filas con el mismo OF, Pieza y NumPaso, se sobrescribe con la última
                # (no se promedia, se toma el último valor encontrado)
                datos_acumulados[key] = {
                    'izquierda': [par_izda],  # Solo guardar el último valor
                    'derecha': [par_drch],
                    'consumo_izquierda': consumo_izda,  # Consumo en amperios
                    'consumo_derecha': consumo_drch
                }
                
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
                        '0': {'izquierda': 0, 'derecha': 0, 'consumo_izquierda': 0, 'consumo_derecha': 0},
                        '100': {'izquierda': 0, 'derecha': 0, 'consumo_izquierda': 0, 'consumo_derecha': 0}
                    }
                }
            
            # Calcular promedio y redondear a 1 decimal (para par)
            avg_izda = round(sum(valores['izquierda']) / len(valores['izquierda']), 1) if valores['izquierda'] else 0
            avg_drch = round(sum(valores['derecha']) / len(valores['derecha']), 1) if valores['derecha'] else 0
            
            datos_por_pieza[clave_pieza]['cargas'][percent]['izquierda'] = avg_izda
            datos_por_pieza[clave_pieza]['cargas'][percent]['derecha'] = avg_drch
            
            # Para el 100%, usar consumos del paso 5 si están disponibles, sino usar los del paso 4
            if percent == '100':
                key_paso5 = (of, pieza)
                if key_paso5 in consumos_paso5:
                    # Usar consumos del paso 5 para el 100%
                    datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_izquierda'] = consumos_paso5[key_paso5].get('consumo_izquierda', 0)
                    datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_derecha'] = consumos_paso5[key_paso5].get('consumo_derecha', 0)
                else:
                    # Si no hay paso 5, usar los del paso 4
                    datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_izquierda'] = valores.get('consumo_izquierda', 0)
                    datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_derecha'] = valores.get('consumo_derecha', 0)
            else:
                # Para el 0%, usar consumos normalmente
                datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_izquierda'] = valores.get('consumo_izquierda', 0)
                datos_por_pieza[clave_pieza]['cargas'][percent]['consumo_derecha'] = valores.get('consumo_derecha', 0)
        
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
        titulo = "Informe de Par Masivo"
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
        
        # Tabla resumen de todas las piezas (solo 0% y 100%)
        resumen_headers = [['Pieza', '0% Par (Nm)', '0% Consumo (A)', '100% Par (Nm)', '100% Consumo (A)']]
        resumen_data = resumen_headers.copy()
        
        # Ordenar piezas por OF y número de pieza
        piezas_ordenadas = sorted(piezas.items(), key=lambda x: (
            x[1].get('of', 0),
            x[1].get('pieza', 0)
        ))
        
        for pieza_id, pieza_info in piezas_ordenadas:
            referencia = pieza_info.get('referencia', pieza_id)
            cargas = pieza_info.get('cargas', {})
            
            # Calcular valores para 0% y 100%
            valores = []
            for percent in ['0', '100']:
                if percent in cargas:
                    izda = cargas[percent].get('izquierda', 0)
                    drch = cargas[percent].get('derecha', 0)
                    promedio_par = (abs(izda) + abs(drch)) / 2
                    consumo_izda = cargas[percent].get('consumo_izquierda', 0)
                    consumo_drch = cargas[percent].get('consumo_derecha', 0)
                    promedio_consumo = (abs(consumo_izda) + abs(consumo_drch)) / 2
                    valores.append(f'{promedio_par:.1f}')
                    valores.append(f'{promedio_consumo:.2f}')
                else:
                    valores.append('-')
                    valores.append('-')
            
            resumen_data.append([referencia] + valores)
        
        resumen_table = Table(resumen_data, colWidths=[50*mm, 30*mm, 30*mm, 30*mm, 30*mm])
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
        elements.append(Paragraph('Resumen de Par por Pieza', styles['Heading2']))
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
                
                elements.append(Paragraph('Gráfico de Par', styles['Heading2']))
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
        elements.append(Paragraph(f"Informe de Par - {referencia}", title_style))
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
        carga_headers = [['Porcentaje', 'Izquierda (Nm)', 'Derecha (Nm)']]
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
        elements.append(Paragraph('Datos de Par', styles['Heading2']))
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
                
                elements.append(Paragraph('Gráfico de Par', styles['Heading2']))
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

