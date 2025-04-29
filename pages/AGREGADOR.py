import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime, timedelta
import sys
sys.path.append('/mount/src/aspo_2/')
from DEMANDA import importa_base


# Configuração da página do Streamlit
st.set_page_config(page_title="Agregador de Demandas", page_icon='icone', layout='wide')
st.header('Energisa Mato Grosso - ASPO')
st.subheader("Agregador de Demandas")
#st.divider()

base = importa_base()
if base is None:
    st.stop()


# Leitura dos dados técnicos e de informações
try:
    df_dados_tecnicos = pd.read_excel('Tabela informativa.xlsx', sheet_name='Dados Técnicos')
    df_dados_tecnicos.dropna(subset=['Cód. do Trafo/Alimentador'], inplace=True)
    df_dados_tecnicos_lista = [''] + df_dados_tecnicos['Cód. do Trafo/Alimentador'].tolist()
    df_dados = pd.read_excel('Tabela informativa.xlsx', sheet_name='Dados')
except Exception as e:
    st.error(f"Erro ao carregar as tabelas de dados técnicos: {e}")
    st.stop()

# Interface do Streamlit
st.subheader('Equipamentos')
Coluna_1, Coluna_2, Coluna_3, Coluna_4, Coluna_5, Coluna_6, Coluna_7, Coluna_8, Coluna_9, Coluna_10, Coluna_11 = st.columns(11)
with Coluna_1:
    selecao_1 = st.selectbox("1º Equipamento", df_dados_tecnicos_lista, key='equip_1')
with Coluna_2:
    Operador_1_2 = st.selectbox("Operação 1", ["Somar", "Subtrair"], key='oper_1')
with Coluna_3:
    selecao_2 = st.selectbox("2º Equipamento", df_dados_tecnicos_lista, key='equip_2')
with Coluna_4:
    Operador_2_3 = st.selectbox("Operação 2", ["Somar", "Subtrair"], key='oper_2')
with Coluna_5:
    selecao_3 = st.selectbox("3º Equipamento", df_dados_tecnicos_lista, key='equip_3')
with Coluna_6:
    Operador_3_4 = st.selectbox("Operação 3", ["Somar", "Subtrair"], key='oper_3')
with Coluna_7:
    selecao_4 = st.selectbox("4º Equipamento", df_dados_tecnicos_lista, key='equip_4')
with Coluna_8:
    Operador_4_5 = st.selectbox("Operação 4", ["Somar", "Subtrair"], key='oper_4')
with Coluna_9:
    selecao_5 = st.selectbox("5º Equipamento", df_dados_tecnicos_lista, key='equip_5')
with Coluna_10:
    Operador_5_6 = st.selectbox("Operação 5", ["Somar", "Subtrair"], key='oper_5')
with Coluna_11:
    selecao_6 = st.selectbox("6º Equipamento", df_dados_tecnicos_lista, key='equip_6')
def obter_descricoes(selecao):
    if selecao:
        codigos = [f"{selecao}-EAE", f"{selecao}-EAR", f"{selecao}-ERE", f"{selecao}-ERR"]
        descricoes = []
        print(descricoes)
        for codigo in codigos:
            if df_dados[df_dados['Codigo'] == codigo].empty:
                descricoes.append(None)
            else:
                descricoes.append(df_dados.loc[df_dados['Codigo'] == codigo, 'descricao'].iloc[0])
                print(codigo)
        return descricoes
    return [None] * 10


def iniciar_calculo(selecao_1, selecao_2, selecao_3, selecao_4,selecao_5,selecao_6,Operador_1_2, Operador_2_3, Operador_3_4,Operador_4_5,Operador_5_6):

    descricoes_1 = obter_descricoes(selecao_1)
    descricoes_2 = obter_descricoes(selecao_2)
    descricoes_3 = obter_descricoes(selecao_3)
    descricoes_4 = obter_descricoes(selecao_4)
    descricoes_5 = obter_descricoes(selecao_5)
    descricoes_6 = obter_descricoes(selecao_6)

    base_descricoes = descricoes_1 + descricoes_2 + descricoes_3 + descricoes_4 + descricoes_5 + descricoes_6
    base_filtered = base.copy()

    for descricao in base_descricoes:
        if descricao and descricao not in base.columns:
            base_filtered[descricao] = 0.0

    def calcular_potencia(descricoes):
        P = (base_filtered[descricoes[0]] if descricoes[0] else 0) - (
            base_filtered[descricoes[1]] if descricoes[1] else 0)
        Q = (base_filtered[descricoes[2]] if descricoes[2] else 0) - (
            base_filtered[descricoes[3]] if descricoes[3] else 0)
        S = np.sqrt(P ** 2 + Q ** 2)
        fp = P / S if (S != 0).all() else 0
        return P, Q, S, fp

    P1, Q1, S1, fp1 = calcular_potencia(descricoes_1)
    P2, Q2, S2, fp2 = calcular_potencia(descricoes_2)
    P3, Q3, S3, fp3 = calcular_potencia(descricoes_3)
    P4, Q4, S4, fp4 = calcular_potencia(descricoes_4)
    P5, Q5, S5, fp5 = calcular_potencia(descricoes_5)
    P6, Q6, S6, fp6 = calcular_potencia(descricoes_6)

    if Operador_1_2 == "Subtrair":
        P2, Q2 = -P2, -Q2
    if Operador_2_3 == "Subtrair":
        P3, Q3 = -P3, -Q3
    if Operador_3_4 == "Subtrair":
        P4, Q4 = -P4, -Q4
    if Operador_4_5 == "Subtrair":
        P5, Q5 = -P5, -Q5
    if Operador_5_6 == "Subtrair":
        P6, Q6 = -P6, -Q6

    P_total = P1 + P2 + P3 + P4 + P5 + P6
    Q_total = Q1 + Q2 + Q3 + Q4 + Q5 + Q6
    S_total = np.sqrt(P_total ** 2 + Q_total ** 2)
    fp_total = P_total / S_total if (S_total != 0).all() else 0

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=base_filtered.index, y=P_total, name='Potencia Ativa - kW', line=dict(color='Navy')))
    fig.add_trace(
        go.Scatter(x=base_filtered.index, y=Q_total, name='Potencia Reativa - kVAr', line=dict(color='Tomato')))
    fig.add_trace(go.Scatter(x=base.index, y=S_total, name='Potencia Aparente - kVA', line=dict(color='DarkCyan')))
    fig.update_layout(title="Gráfico Com Demandas Agregadas", xaxis_title="Data", yaxis_title="Valor", width=1480,
                      height=600)
    st.plotly_chart(fig)

if st.button('Calcular'):
    iniciar_calculo(selecao_1, selecao_2, selecao_3, selecao_4,selecao_5,selecao_6,Operador_1_2,Operador_2_3,Operador_3_4,Operador_4_5,Operador_5_6)
