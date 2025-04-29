import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import numpy as np
import sqlite3 as sql
from st_aggrid import AgGrid, GridOptionsBuilder
import os

st.set_page_config(page_title="Operação",page_icon='icone',layout='wide')
# Carregar os dados do Excel
url = r"Demanda_Máxima_Semana.xlsx"
df = pd.read_excel(url, sheet_name="Potência Ativa")

# Carregar dados de demanda máxima
url_demanda_maxima = r"Demanda_Máxima_Não_Coincidente.xlsx"
df_demanda_maxima = pd.read_excel(url_demanda_maxima, sheet_name="Potência Aparente")

# Título da aplicação
st.title("Potência Máxima Semanal - kVA")

# Obter lista de códigos e ajustar formato se necessário
codigos_unicos = df['Cód. do Trafo/Alimentador'].unique()
codigos_ajustados = [f"AL-{cod}" if not str(cod).startswith('AL-') else str(cod) for cod in codigos_unicos]
codigos_ajustados.sort()

# Selecionar o transformador ou alimentador
selecao = st.selectbox("Selecione o Cód. do Trafo/Alimentador:", codigos_ajustados, index=1)

# Remover o prefixo AL- para buscar no DataFrame original
selecao_original = selecao.replace('AL-', '') if selecao.startswith('AL-') else selecao

# Filtrar os dados com base na seleção
df_filtrado = df[df['Cód. do Trafo/Alimentador'] == selecao_original]

# Verificar se há dados para o transformador ou alimentador selecionado
if not df_filtrado.empty:
    # Extrair os dados da linha (colunas 8 até antepenúltimas 5)
    dados_semanais = df_filtrado.iloc[0, 7:-5]
    
    # Converter os valores para float, ignorando erros
    valores_convertidos = pd.to_numeric(dados_semanais, errors='coerce')
    
    # Manter apenas valores válidos
    week_labels = dados_semanais.index[valores_convertidos.notna()]
    values = valores_convertidos[valores_convertidos.notna()]

    # Verifica se existem dados válidos
    if not values.empty:
        # Criar o gráfico
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=week_labels,
            y=values,
            mode='lines+markers',
            name='Potência KVA',
            line=dict(color='royalblue', width=2),
            marker=dict(size=6)
        ))

        # Verificar as colunas disponíveis
        print("Colunas disponíveis:", df_demanda_maxima.columns.tolist())
        
        # Adicionar linha horizontal da potência instalada
        try:
            # Ajustar o código do alimentador para incluir o prefixo AL-
            codigo_ajustado = f"AL-{selecao}" if not selecao.startswith('AL-') else selecao
            
            # Verificar se existem dados para o transformador/alimentador selecionado
            dados_potencia = df_demanda_maxima.loc[df_demanda_maxima['Cód. do Trafo/Alimentador'] == codigo_ajustado, 'Potencia Instalada']
            
            if not dados_potencia.empty:
                potencia_instalada = dados_potencia.iloc[0]
                fig.add_shape(
                    type="line",
                    x0=week_labels[0],
                    y0=potencia_instalada,
                    x1=week_labels[-1],
                    y1=potencia_instalada,
                    line=dict(color="red", width=3.0, dash="dash"),
                    name="Potência Instalada"
                )
            else:
                st.warning(f"Não foram encontrados dados de potência instalada para o transformador/alimentador {codigo_ajustado}")
                
        except KeyError:
            st.warning("Coluna 'Potencia Instalada' não encontrada. Verifique o nome da coluna no arquivo de dados.")

        fig.update_layout(
            title=f'Potência Máxima Semanal - {selecao}',
            xaxis_title='Semana',
            yaxis_title='Potência Máxima (kW)',
            xaxis_tickangle=-45,
            template='plotly_white',
            margin=dict(l=40, r=40, t=80, b=100),
            height=500,
            yaxis=dict(
                tickformat=",.0f",
                separatethousands=True
            )
        )

        st.plotly_chart(fig, use_container_width=True)
        
        # Adicionar resumo estatístico
        st.subheader("Resumo Estatístico")
        
        # Calcular estatísticas
        max_value = values.max()
        min_value = values.min()
        mean_value = values.mean()
        std_value = values.std()
        
        # Encontrar datas dos valores máximos e mínimos
        max_date = week_labels[values == max_value].tolist()
        min_date = week_labels[values == min_value].tolist()
        
        # Criar colunas para exibir as métricas
        col1, col2, col3 = st.columns(3)
        
        # Função para formatar números no padrão brasileiro
        def format_br(value):
            return '{:,.0f}'.format(value).replace(',', '.').replace('.', ',', 1)
        
        with col1:
            st.metric("Valor Máximo", f"{format_br(max_value)} kVA")
            st.text(f"Data: {max_date[0] if max_date else 'N/A'}")
        
        with col2:
            st.metric("Valor Mínimo", f"{format_br(min_value)} kVA")
            st.text(f"Data: {min_date[0] if min_date else 'N/A'}")
        
        with col3:
            st.metric("Média", f"{format_br(mean_value)} kVA")

    else:
        st.warning("Não há dados numéricos válidos para plotar o gráfico.")
else:
    st.warning("Nenhum dado encontrado para o Cód. do Trafo/Alimentador selecionado.")

st.markdown("---")

# Exibir tabela de demanda máxima
st.subheader("Demanda Máxima Não Coincidente")

# Filtrar os dados com o código original (sem AL-)
df_demanda_maxima_filtrado = df_demanda_maxima[df_demanda_maxima['Cód. do Trafo/Alimentador'] == selecao].copy()

# Verificar se existem dados para o equipamento selecionado
if not df_demanda_maxima_filtrado.empty:
    # Criar um novo DataFrame com as colunas necessárias
    colunas_numericas = [col for col in df_demanda_maxima_filtrado.columns 
                        if col not in ['Cód. do Trafo/Alimentador', 'Descrição']]
    colunas_texto = ['Cód. do Trafo/Alimentador', 'Descrição']

    # Criar DataFrame formatado
    df_formatado = pd.DataFrame()

    # Processar colunas de texto
    for col in colunas_texto:
        df_formatado[col] = df_demanda_maxima_filtrado[col].fillna('').astype(str)

    # Processar colunas numéricas
    for col in colunas_numericas:
        valores = pd.to_numeric(df_demanda_maxima_filtrado[col], errors='coerce')
        df_formatado[col] = valores.apply(lambda x: '{:,.0f}'.format(x).replace(',', '.').replace('.', ',', 1) if pd.notnull(x) else "")

    # Criar HTML para a tabela
    html = """
    <style>
        .dataframe {
            width: 90%;
            border-collapse: collapse;
        }
        .dataframe th, .dataframe td {
            padding: 4px;
            text-align: center;
            border: 1px solid #ddd;
        }
        .dataframe th {
            background-color: #f2f2f2;
        }
        .dataframe tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .dataframe td {
            white-space: nowrap;
        }
    </style>
    <table class="dataframe">
    """

    # Adicionar cabeçalho
    html += "<tr>"
    for col in df_formatado.columns:
        html += f"<th>{col}</th>"
    html += "</tr>"

    # Adicionar dados
    for _, row in df_formatado.iterrows():
        html += "<tr>"
        for col in df_formatado.columns:
            html += f"<td>{row[col]}</td>"
        html += "</tr>"

    html += "</table>"

    # Exibir a tabela HTML
    st.markdown(html, unsafe_allow_html=True)
else:
    st.warning(f"Não foram encontrados dados para o equipamento {selecao}")

# Botão de exportação
buffer = io.StringIO()
df.to_csv(buffer, sep=';', encoding='latin-1', index=False)
buffer.seek(0)
file_name = f"{selecao}_Demanda_Maxima_Semanal.csv"

st.download_button(
    label="📥 Downloads - Todos os dados",
    data=buffer.getvalue(),
    file_name=file_name,
    mime="text/csv")

st.markdown("---")


st.title("Ajustes de proteção dos equipamentos - Pickup")

# abrir o banco de dados sqlite convertendo para dataframe
conn = sql.connect('equipamentos.db')
cursor = conn.cursor()
df_equipamentos = pd.read_sql_query("SELECT * FROM equipamentos", conn)

# criar multiselect com a coluna Subestação, inicialmente exibir todos os dados
subestacao = st.multiselect(
    "Selecione a Subestação",
    df_equipamentos['Subestação'].unique(),
    placeholder='Selecione a Subestação'
)

# opções para tipo de equipamento
opcoes = ['Todos'] + df_equipamentos['Tipo'].unique().tolist()
tipo_equipamento = st.radio(
    "Selecione o tipo de equipamento",
    opcoes,
    horizontal=True
)

# filtro por tipo de equipamento
if tipo_equipamento != 'Todos':
    df_equipamentos = df_equipamentos[df_equipamentos['Tipo'] == tipo_equipamento]

# filtro por subestação
if subestacao:
    df_equipamentos = df_equipamentos[df_equipamentos['Subestação'].isin(subestacao)]

# -- MOSTRAR O AGGRID --
st.subheader("Tabela de Equipamentos")

# configurando AgGrid
gb = GridOptionsBuilder.from_dataframe(df_equipamentos)

gb.configure_pagination(paginationAutoPageSize=True)  # paginação automática
gb.configure_side_bar()  # barra lateral com filtros
gb.configure_default_column(
    groupable=True,
    value=True,
    enableRowGroup=True,
    editable=False,
    filter=True,
)

gridOptions = gb.build()

AgGrid(
    df_equipamentos,
    gridOptions=gridOptions,
    fit_columns_on_grid_load=True,
    height=500,
    width='100%',
    theme='streamlit'  # opções: 'streamlit', 'light', 'dark', 'blue', 'fresh', 'material'
)

# -- ATUALIZAR A TABELA DE PICKUP --
st.subheader('Atualizar Tabela de Pickup')

# Criar um botão para carregar o arquivo excel
arquivo_pickup = st.file_uploader("Importar Tabela de Pickup", type="xlsx")

def carregar_pickup():
    if arquivo_pickup is not None:
        try:
            # Criar um dataframe com os dados do excel
            df_pickup = pd.read_excel(arquivo_pickup)

            # Extrair SE_ALIMENTADOR e concatenar com Alimentador
            df_pickup['SE_ALIMENTADOR'] = df_pickup['Subestação'].str.extract(r'^2(\d{3})')[0].astype(str) + df_pickup['Alimentador'].astype(str)

            # Exibir o dataframe
            st.dataframe(df_pickup)

            # Criar banco de dados sqlite
            conn = sql.connect('equipamentos.db')
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS equipamentos (
                    Tipo TEXT,
                    Subestação TEXT,
                    Alimentador TEXT,
                    Religador TEXT,
                    Marca_Modelo TEXT,
                    Pickup FLOAT,
                    SE_ALIMENTADOR TEXT,
                    Tensão FLOAT,
                    Potência_limite FLOAT
                )
            """)

            # Salvar em banco de dados sqlite
            df_pickup.to_sql('equipamentos', conn, if_exists='replace', index=False)

            # Fechar a conexão
            conn.close()

            st.success("Dados carregados e salvos com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {str(e)}")
    else:
        st.info("Por favor, carregue um arquivo Excel para continuar.")

if arquivo_pickup is not None:
    carregar_pickup()

st.text("\nDesenvolvido por Arthur Williams")