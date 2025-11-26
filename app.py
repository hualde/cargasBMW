from flask import Flask, render_template, request, jsonify
import json
import os

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

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)

