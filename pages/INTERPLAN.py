import matplotlib.pyplot as plt
import csv
import numpy as np
import openpyxl
import os
import pandas as pd
import streamlit as st
import datetime as dt
from datetime import datetime
import tempfile
import base64

contador =0
horario =0

#################################### CABEÇALHO ####################################################
st.set_page_config(layout="wide")
st.header('Energisa Mato Grosso - ASPO')
st.subheader('Importação da Demanda Base em Lote - Interplan')
st.markdown("Assessoria de Planejamento e Orçamento")
st.markdown("Demanda base para Alimentadores e Subestações")

########################### IMPORTAÇÃO DA BASE DE EQUIPAMENTOS ###################################
st.subheader('Definir Alimentadores ou Transformadores')
print("Importando a tabela de tributos:")
df_atributos = pd.read_excel('Tabela informativa.xlsx', sheet_name='Dados Técnicos')
df_atributos = df_atributos["Cód. do Trafo/Alimentador"]
print(df_atributos)
df_equipamentos = st.multiselect("Equipamentos", df_atributos)
quantidade_elementos = len(df_equipamentos)
print(df_equipamentos)
print("quantidade de elementos:",quantidade_elementos)
######################### SELEÇÃO DO SIDERBAR #########################################################
data_atual = datetime.now()
ano = st.sidebar.selectbox("Selecione o ano:", [2023, 2024])
opcoes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho (Demanda Mínima)","Julho", "Agosto", "Setembro", "Outubro (Demanda Máxima)", "Novembro", "Dezembro"]
selecao_mes = st.sidebar.selectbox("Selecione o mês:", opcoes_meses, index=9)

mes_inteiro = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho (Demanda Mínima)": 6,
               "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro (Demanda Máxima)": 10, "Novembro": 11, "Dezembro": 12}
################################################################################################################
st.divider()
st.subheader('Definir Horários dos Patamares de Carga')
Coluna_8, Coluna_9, Coluna_10, Coluna_11 = st.columns(4)
with Coluna_8:
    Madrugada = st.radio('Madrugada', [1,2,3,4,5,6],index=3)
with Coluna_9:
    Manha = st.radio("Manhã",[7,8,9,10,11,12],index=5)
with Coluna_10:
    Tarde= st.radio("Tarde",[13,14,15,16,17,18],index=4)
with Coluna_11:
    Noite = st.radio("Noite",[19,20,21,22,23,24],index=3)

st.divider()
st.subheader('Parâmetros de Calculo de Fluxo de Potência')

Coluna_1, Coluna_2,Coluna_3,Coluna_4,Coluna_5,Coluna_6,Coluna_7 = st.columns(7)
with Coluna_1:
    calcular_demanda = st.radio('Cálculo de Demanda', ["Não Calcular","Demanda sem Correção","Com correção de EPs e ETs","Com Correção de EPs","Com Correção de ETs"],index=4)
with Coluna_2:
    patamar = st.radio("Patamar", ["Madrugada", "Manhã", "Tarde", "Noite", "Maior Demanda"],index=3)
with Coluna_3:
    corrigir_curva = st.radio("Corrigir a Curva", ["Manter o Perfil da Curva", "Apenas Patamar Selecionado"],index=0)
with Coluna_4:
    curvas_demanda_bt = st.radio("Curvas de Demanda BT", ["Utilizar", "Não utilizar"],index=1)
with Coluna_5:
    tipo_correcao = st.radio("Tipo de Correção", ["Corrente","Potência"],index=1)
with Coluna_6:
    valores_medidos = st.radio("Valores Medidos", ["Não", "Sim"],index=0)
with Coluna_7:
    valores_por_fase = st.radio("Valores por Fase", ["Não", "Sim"],index=0)
# url_atributos = r"C:\Users\Engeselt\Documents\Tabela informativa.xlsx"
df_atributos_Dados = pd.read_excel('Tabela informativa.xlsx', sheet_name="Dados")
df_atributos_Dados.dropna(subset=['Codigo'], inplace=True)
print(df_atributos_Dados)

######################### SELECÃO DOS EQUIPAMENTOS ###################################################
if st.button("Processar Informações"):
    if calcular_demanda == "Não Calcular":
        calcular_demanda=int(0)
    if calcular_demanda == "Demanda Sem Correção":
        calcular_demanda=int(1)
    if calcular_demanda == "Com Correção de EPs":
        calcular_demanda=int(2)
    if calcular_demanda == "Com Correção de EPs":
        calcular_demanda=int(3)
    if calcular_demanda == str("Com Correção de ETs"):
        calcular_demanda=int(4)
    print(calcular_demanda)

    if patamar == "Madrugada":
        patamar=0
    elif patamar == "Manhã":
        patamar=1
    elif patamar == "Tarde":
        patamar=2
    elif patamar == "Noite":
        patamar=3
    elif patamar == "Maior Demanda":
        patamar=4
    print(patamar)

    if corrigir_curva == "Manter o Perfil da Curva":
        corrigir_curva = int(0)
    else:corrigir_curva = int(1)
    print(corrigir_curva)
    if curvas_demanda_bt == "Utilizar":
        curvas_demanda_bt = int(0)
    else: curvas_demanda_bt= int(1)
    print(curvas_demanda_bt)
    if tipo_correcao == "Corrente":
        tipo_correcao = int(0)
    else: tipo_correcao = int(1)
    print(tipo_correcao)
    if valores_medidos == "Não":
        valores_medidos = int(0)
    else: valores_medidos= int(1)
    print(valores_medidos)
    if valores_por_fase =="Não":
        valores_por_fase=int(0)
    else: valores_por_fase = int(1)
    print(valores_por_fase)

    #########################################  CALCULOS DOS PARÂMETROS #####################################################
    valores_maximos_P_meses = []
    valores_maximos_Q_meses = []
    lista_dados = []
    df = pd.DataFrame(lista_dados)
    def importa_base():
        print("Importando Base de Dados Agrupada\n")
        #url_base = r"C:\Users\Engeselt\Documents\GitHub\ASPO_2\Medição Agrupada.csv"
        df_base = pd.read_csv("Medição Agrupada.csv", sep=";", encoding='latin-1').set_index(['DATA_HORA'])
        print(df_base)
        print("\nImportação Concluída")
        return df_base
        #lista_dados = []
    df_base =importa_base()
    df_base.index = pd.to_datetime(df_base.index)
    print("Corrigir Curva:",corrigir_curva)
    if corrigir_curva == int (1):
        print("Patamar:",patamar)
        if patamar==0:
            df_filtrado = df_base.loc[(df_base.index.year == ano) & (df_base.index.month == mes_inteiro[selecao_mes]) & (df_base.index.hour == Madrugada)]
        if patamar==1:
            df_filtrado = df_base.loc[(df_base.index.year == ano) & (df_base.index.month == mes_inteiro[selecao_mes]) & (df_base.index.hour == Manha)]
        if patamar==2:
            df_filtrado = df_base.loc[(df_base.index.year == ano) & (df_base.index.month == mes_inteiro[selecao_mes]) & (df_base.index.hour == Tarde)]
        if patamar==3:
            df_filtrado = df_base.loc[(df_base.index.year == ano) & (df_base.index.month == mes_inteiro[selecao_mes]) & (df_base.index.hour == Noite)]
    else:
        df_filtrado = df_base.loc[(df_base.index.year == ano) & (df_base.index.month == mes_inteiro[selecao_mes]) & ((df_base.index.hour == Madrugada)|(df_base.index.hour ==Manha)|(df_base.index.hour == Tarde)|(df_base.index.hour == Noite))]

    for selecao in df_equipamentos:
        print(selecao)
        EAE = str(selecao) + '-EAE'
        EAR = str(selecao) + '-EAR'
        ERE = str(selecao) + '-ERE'
        ERR = str(selecao) + '-ERR'
        filtro = df_atributos_Dados["Codigo"] == EAE
        indices_encontrados = df_atributos_Dados.loc[filtro, "Codigo"].index
        print('Indice encontrado:',indices_encontrados)

        if len(indices_encontrados) > 0:
            indice_entrada = indices_encontrados[0]
            descricao = df_atributos_Dados.loc[indice_entrada, 'descricao']
        else:
            print("Chave não encontrada em df_atributos_Dados:", EAE)
            continue

        #descricao = (df_atributos_Dados.loc[indice_entrada, 'descricao'])
        print(descricao)
        descricao = str(descricao)
        alimentador = descricao.replace("-EAE", "")
        indice_saida = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == EAR, 'Codigo'].index[0]
        descricao_saida = df_atributos_Dados.loc[indice_saida, 'descricao']
        descricao_saida = str(descricao_saida)
        print(descricao_saida)
        indice_entrada_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERE, 'Codigo'].index[0]
        descricao_Q = (df_atributos_Dados.loc[indice_entrada_Q, 'descricao'])
        descricao_Q = str(descricao_Q)
        print(descricao_Q)
        indice_saida_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERR, 'Codigo'].index[0]
        descricao_saida_Q = df_atributos_Dados.loc[indice_saida_Q, 'descricao']
        descricao_saida_Q = str(descricao_saida_Q)
        print(descricao_saida_Q)
        ########################################################################################################
        ###################################### BASE FILTRADA ###################################################
        df_base_filtrado = pd.DataFrame(df_filtrado, columns=['DATA_HORA',descricao, descricao_saida, descricao_Q, descricao_saida_Q])
        df_base_filtrado.fillna(0.0)
        df_base_filtrado[descricao] = df_base_filtrado[descricao].astype(float)
        df_base_filtrado[descricao_saida] = df_base_filtrado[descricao_saida].astype(float)
        df_base_filtrado[descricao_Q] = df_base_filtrado[descricao_Q].astype(float)
        df_base_filtrado[descricao_saida_Q] = df_base_filtrado[descricao_saida_Q].astype(float)
        ########################################################################################################

        ####################################### CALCULO DOS PARAMETROS##########################################

        df_base_filtrado['P'] = df_base_filtrado[descricao] - df_base_filtrado[descricao_saida]
        df_base_filtrado['Q'] = df_base_filtrado[descricao_Q] - df_base_filtrado[descricao_saida_Q]
        df_base_filtrado['S'] = np.sqrt((df_base_filtrado['P'] ** 2) + (df_base_filtrado['Q'] ** 2))
        df_base_filtrado['fp'] = df_base_filtrado['P'] / df_base_filtrado['S']


        #######################################################################################################
        ####################### CALCULO DOS VALORES MÁXIMOS  E MINIMOS ########################################
        valor_maximo_P = max(df_base_filtrado['P'])
        valor_media_p = np.mean(df_base_filtrado['P'])
        valor_media_Q = np.mean(df_base_filtrado['Q'])
        valor_maximo_S = max(df_base_filtrado['S'])
        valor_maximo_S = round(float(valor_maximo_S), 2)
        valor_minimo_P = min(df_base_filtrado['P'])

        #######################################################################################################

        #########################CALCULA VALOR MINIMO E O DESVIO PADRÃO SEM O ZERO ###########################

        dados_sem_zeros = np.array([x for x in df_base_filtrado['P'] if x != 0]).astype(float)
        if len(dados_sem_zeros) > 0:
            valor_minimo_P_sem_zero = np.min(dados_sem_zeros)
            desvio_padrao_P = np.nanstd(dados_sem_zeros)
        else:
            valor_minimo_P_sem_zero = 0
            desvio_padrao_P = 0
        #print("Dados sem o zero\n",dados_sem_zeros)
        #valor_minimo_P_sem_zero = min(dados_sem_zeros)
        print("Valor Minimo sem o zero:\n", valor_minimo_P_sem_zero)
        desvio_padrao_P = np.nanstd(df_base_filtrado['P'])
        print("desvio Padrão: ", desvio_padrao_P)
        ########################################################################################################

        ################################# CALCULO DO DESVIO PADRÃO #############################################
        minimo_desvio = round(valor_media_p - 3 * desvio_padrao_P, 0)
        maximo_desvio = round((valor_media_p + 3 * desvio_padrao_P), 0)
        print("Valor minimo por desvio padrão:", minimo_desvio)
        print("Valor maximo por desvio padrão:", maximo_desvio)
        ########################################################################################################
        ############################### CALCULO DO FATOR DE POTENCIA ###########################################

        fp = round(valor_media_p / (np.mean(df_base_filtrado['S'])), 2)
        ########################################################################################################
        ###########################  Potencia máxima considerada ###############################################
        base_filtrada = np.where(df_base_filtrado['P'] < valor_media_p + 3.0 * desvio_padrao_P, df_base_filtrado['P'], np.nan)
        valor_maximo_filtrado = np.nanmax(base_filtrada)

        if valor_maximo_P / valor_maximo_filtrado > 1.10:
            valor_maximo_P = valor_maximo_filtrado
        else:
            valor_maximo_P = valor_maximo_P

        # Obtenha a lista de índices onde 'P' é igual ao valor mínimo
        max_date_time_index = df_base_filtrado.index[df_base_filtrado['P'] == valor_maximo_P].tolist()

        print("Data e hora do valor máximo: ", max_date_time_index)

        # Certifique-se de que há pelo menos um índice encontrado
        if max_date_time_index and not pd.isna(max_date_time_index[0]):
            # Pegue o primeiro índice encontrado (você pode ajustar conforme necessário)
            primeiro_indice = max_date_time_index[0]

            # Acesse diretamente o valor da coluna 'Q' correspondente ao índice mínimo de 'P'
            valor_Q_correspondente = df_base_filtrado.at[primeiro_indice, 'Q']
            print(f"Valor de 'Q' correspondente ao mínimo de 'P': {valor_Q_correspondente}")
        else:
            print("Nenhum índice encontrado para o valor mínimo de 'P'")

        potencia_aparente_maxima = np.sqrt((valor_maximo_P ** 2) + (valor_Q_correspondente ** 2))
        ########################################################################################################################
        data = {
            "Alimentador": str(selecao),
            "Imax (A)": 0,
            "P3F (kW)": round(valor_maximo_P,0),
            "Q3F (kVAr)": round(valor_Q_correspondente,0),
            "Cálculo de Demanda": calcular_demanda,
            "Patamar": patamar,
            "Corrigir a Curva": int(corrigir_curva),
            "Curvas de Demanda BT": curvas_demanda_bt,
            "Tipo de Correção": tipo_correcao,
            "Valores Medidos": valores_medidos,
            "Valores por Fase": valores_por_fase
        }
        lista_dados.append(data)
        df = pd.DataFrame(lista_dados)
        # Preenchendo a string com zeros à esquerda para que tenha 6 caracteres
        df['Alimentador'] = df['Alimentador'].str.zfill(6)
        print("inicio da exportação")
        #caminho_destino = r"C:\Users\Engeselt\Documents\GitHub\ASPO_2"
        df.to_csv("INTERPLAN.txt",index=False, sep=" ",encoding='latin-1')
        print("Arquivo Exportado")
        print(df)
# Verifica se o arquivo INTERPLAN.csv existe
if os.path.isfile("INTERPLAN.txt"):
    # Carrega o arquivo CSV
    df = pd.read_csv("INTERPLAN.txt", sep=" ", encoding='latin-1')

    # Adiciona um botão para baixar o arquivo
    if st.button("Baixar arquivo INTERPLAN.csv"):
        # Cria um link para download do arquivo
        csv = df.to_csv(index=False, sep=";", encoding='latin-1')
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="INTERPLAN.csv">Baixar arquivo CSV</a>'
        st.markdown(href, unsafe_allow_html=True)
else:
    st.warning("O arquivo INTERPLAN.csv não foi encontrado.")


