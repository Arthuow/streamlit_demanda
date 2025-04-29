from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
import io
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
df_atributos_Dados = None
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
        global df_atributos, df_atributos_Dados
        df_atributos_Dados = pd.read_excel(atributos_file, sheet_name="Dados")
        df_atributos_Dados.dropna(subset=['Codigo'], inplace=True)
        
        df_atributos = pd.read_excel(atributos_file, sheet_name="Dados Técnicos")
        df_atributos.dropna(subset=['Cód. do Trafo/Alimentador'], inplace=True)

        return jsonify({'message': 'Arquivos carregados com sucesso'})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/calcular', methods=['POST'])
def calcular_demanda():
    try:
        global resultados
        resultados = []
        
        # Configurar anos e meses
        anos = range(2024, 2025)
        meses = range(1, 13)
        
        for ano in anos:
            df_filtrado_ano = df_base[df_base["ANO"] == ano]
            for mes in meses:
                if mes not in df_filtrado_ano['MES'].values:
                    continue
                    
                df_filtrado_mes = df_filtrado_ano[df_filtrado_ano["MES"] == mes]
                
                for selecao in df_atributos['Cód. do Trafo/Alimentador']:
                    if selecao == 0:
                        continue
                        
                    try:
                        potencia_instalada = df_atributos.loc[
                            df_atributos['Cód. do Trafo/Alimentador'] == selecao,
                            'Potencia Instalada'
                        ].values[0]
                    except IndexError:
                        potencia_instalada = 0
                        
                    # Processamento dos dados para cada alimentador
                    EAE = str(selecao) + '-EAE'
                    EAR = str(selecao) + '-EAR'
                    ERE = str(selecao) + '-ERE'
                    ERR = str(selecao) + '-ERR'
                    
                    filtro = df_atributos_Dados["Codigo"] == EAE
                    indices_encontrados = df_atributos_Dados.loc[filtro, "Codigo"].index
                    
                    if len(indices_encontrados) > 0:
                        indice_entrada = indices_encontrados[0]
                        descricao = df_atributos_Dados.loc[indice_entrada, 'descricao']
                    else:
                        continue
                        
                    descricao = str(descricao)
                    alimentador = descricao.replace("-EAE", "")
                    indice_saida = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == EAR, 'Codigo'].index[0]
                    descricao_saida = df_atributos_Dados.loc[indice_saida, 'descricao']
                    descricao_saida = str(descricao_saida)
                    indice_entrada_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERE, 'Codigo'].index[0]
                    descricao_Q = df_atributos_Dados.loc[indice_entrada_Q, 'descricao']
                    descricao_Q = str(descricao_Q)
                    indice_saida_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERR, 'Codigo'].index[0]
                    descricao_saida_Q = df_atributos_Dados.loc[indice_saida_Q, 'descricao']
                    descricao_saida_Q = str(descricao_saida_Q)
                    
                    # Criar DataFrame filtrado
                    df_base_filtrado = pd.DataFrame(df_filtrado_mes, columns=['DATA_HORA', descricao, descricao_saida, descricao_Q, descricao_saida_Q])
                    df_base_filtrado.fillna(0.0)
                    
                    # Calcular valores
                    df_base_filtrado[descricao] = df_base_filtrado[descricao].astype(float)
                    df_base_filtrado[descricao_saida] = df_base_filtrado[descricao_saida].astype(float)
                    df_base_filtrado[descricao_Q] = df_base_filtrado[descricao_Q].astype(float)
                    df_base_filtrado[descricao_saida_Q] = df_base_filtrado[descricao_saida_Q].astype(float)
                    
                    df_base_filtrado['P'] = df_base_filtrado[descricao] - df_base_filtrado[descricao_saida]
                    df_base_filtrado['Q'] = df_base_filtrado[descricao_Q] - df_base_filtrado[descricao_saida_Q]
                    df_base_filtrado['S'] = np.sqrt((df_base_filtrado['P'] ** 2) + (df_base_filtrado['Q'] ** 2))
                    
                    # Calcular valores máximos
                    valor_maximo_P = max(df_base_filtrado['P'])
                    valor_media_p = np.mean(df_base_filtrado['P'])
                    valor_media_Q = np.mean(df_base_filtrado['Q'])
                    valor_maximo_S = max(df_base_filtrado['S'])
                    valor_maximo_S = round(float(valor_maximo_S), 2)
                    valor_minimo_P = min(df_base_filtrado['P'])
                    
                    # Calcular desvio padrão
                    dados_sem_zeros = np.array([x for x in df_base_filtrado['P'] if x != 0]).astype(float)
                    if len(dados_sem_zeros) > 0:
                        valor_minimo_P_sem_zero = np.min(dados_sem_zeros)
                        desvio_padrao_P = np.nanstd(dados_sem_zeros)
                    else:
                        valor_minimo_P_sem_zero = 0
                        desvio_padrao_P = 0
                        
                    # Calcular desvios
                    minimo_desvio = round(valor_media_p - 2 * desvio_padrao_P, 0)
                    maximo_desvio = round((valor_media_p + 3 * desvio_padrao_P), 0)
                    
                    # Calcular fator de potência
                    fp = round(valor_media_p / np.mean(df_base_filtrado['S']), 2)
                    
                    # Calcular potência máxima considerada
                    base_filtrada = np.where(df_base_filtrado['P'] < valor_media_p + 3.0 * desvio_padrao_P, 
                                           df_base_filtrado['P'], np.nan)
                    valor_maximo_filtrado = np.nanmax(base_filtrada)
                    
                    if valor_maximo_P / valor_maximo_filtrado > 1.10:
                        valor_maximo_P = valor_maximo_filtrado
                        
                    # Encontrar data e hora do valor máximo
                    max_date_time_index = df_base_filtrado.index[df_base_filtrado['P'] == valor_maximo_P].tolist()
                    
                    if max_date_time_index and not pd.isna(max_date_time_index[0]):
                        valor_Q_correspondente = df_base_filtrado.at[max_date_time_index[0], 'Q']
                    else:
                        valor_Q_correspondente = 0
                        
                    potencia_aparente_maxima = np.sqrt((valor_maximo_P ** 2) + (valor_Q_correspondente ** 2))
                    
                    # Adicionar resultados
                    resultados.append({
                        'Cód. do Trafo/Alimentador': selecao,
                        'Ano': ano,
                        'Mês': mes,
                        'Valor Máximo P': valor_maximo_P,
                        'Valor Máximo S': valor_maximo_S,
                        'Valor Mínimo P': valor_minimo_P,
                        'Valor Mínimo P sem Zero': valor_minimo_P_sem_zero,
                        'Desvio Padrão': desvio_padrao_P,
                        'Fator de Potência': fp,
                        'Potência Aparente Máxima': potencia_aparente_maxima,
                        'Data/Hora Máximo': str(max_date_time_index[0]) if max_date_time_index else None
                    })

        return jsonify({'message': 'Cálculo concluído com sucesso', 'resultados': resultados})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/exportar', methods=['GET'])
def exportar_resultados():
    try:
        if not resultados:
            return jsonify({'error': 'Não há resultados para exportar'}), 400

        # Criar DataFrame com os resultados
        df_resultados = pd.DataFrame(resultados)
        
        # Criar buffer para o arquivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultados.to_excel(writer, index=False)
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'resultados_demanda_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True) 