from flask import Flask, render_template, request, jsonify
import pandas as pd
import numpy as np
from datetime import datetime
import os

app = Flask(__name__)

# Configuração para upload de arquivos
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Variáveis globais para armazenar os dados
df_base = None
df_atributos = None
resultados = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Verificar se os arquivos foram enviados
        if 'base_file' not in request.files or 'atributos_file' not in request.files:
            return jsonify({'error': 'Por favor, selecione ambos os arquivos'}), 400

        base_file = request.files['base_file']
        atributos_file = request.files['atributos_file']

        if base_file.filename == '' or atributos_file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400

        # Processar arquivo base
        global df_base
        df_base = pd.read_csv(base_file, sep=";", encoding='latin-1')
        df_base['DATA_HORA'] = pd.to_datetime(df_base['DATA_HORA'])
        df_base['MES'] = df_base['DATA_HORA'].dt.month
        df_base['ANO'] = df_base['DATA_HORA'].dt.year

        # Processar arquivo de atributos
        global df_atributos
        df_atributos = pd.read_excel(atributos_file)
        df_atributos.dropna(subset=['Código'], inplace=True)

        return jsonify({'message': 'Arquivos carregados com sucesso'})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/calcular', methods=['POST'])
def calcular():
    try:
        global resultados
        resultados = []
        
        # Aqui você implementará a lógica de cálculo específica do seu novo projeto
        # Por enquanto, retornamos uma mensagem de exemplo
        resultados.append({
            'mensagem': 'Cálculo realizado com sucesso',
            'data': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })

        return jsonify({'message': 'Cálculo concluído com sucesso', 'resultados': resultados})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True) 