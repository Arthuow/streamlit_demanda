import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import numpy as np
import sqlite3 as sql
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import os

st.set_page_config(page_title="Demanda Coincidente", page_icon='icone', layout='wide')

st.header('Demanda Coincidente')
st.markdown("Assessoria de Planejamento e Orçamento")

# Cache para dados que não mudam frequentemente
@st.cache_data
def load_initial_data():
    try:
        # Ajustando o caminho para a pasta raiz
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        excel_path = os.path.join(root_dir, 'input', 'Tabela informativa.xlsx')
        
        # Explicitly specify the engine as 'openpyxl'
        df_atributos = pd.read_excel(excel_path, sheet_name="Dados", engine='openpyxl')
        df_atributos.dropna(subset=['Codigo'], inplace=True)
        
        df_dados_tecnicos = pd.read_excel(excel_path, sheet_name='Dados Técnicos', engine='openpyxl')
        df_dados_tecnicos['Cód. do Trafo/Alimentador'] = df_dados_tecnicos['Cód. do Trafo/Alimentador'].astype(str)
        
        df_atributos_Dados = pd.read_excel(excel_path, sheet_name="Dados", engine='openpyxl')
        df_atributos_Dados.dropna(subset=['Codigo'], inplace=True)
        
        return df_atributos, df_dados_tecnicos, df_atributos_Dados
    except FileNotFoundError:
        st.error("Arquivo 'Tabela informativa.xlsx' não encontrado. Por favor, certifique-se de que o arquivo está na pasta input do projeto.")
        return None, None, None
    except Exception as e:
        st.error(f"Erro ao carregar arquivos: {str(e)}")
        return None, None, None

# Carregar dados iniciais
df_atributos, df_dados_tecnicos, df_atributos_Dados = load_initial_data()
if df_atributos is None:
    st.stop()

# Cache para dados de medição
@st.cache_data
def load_measurement_data(data_hora):
    try:
        # Ajustando o caminho para a pasta raiz
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        csv_path = os.path.join(root_dir, "input", "Medição Agrupada.csv")
        
        # Ler apenas a linha específica do CSV
        df_base_original = pd.read_csv(
            csv_path, 
            sep=";", 
            encoding='latin-1',
            parse_dates=['DATA_HORA']
        )
        # Filtrar apenas a data/hora selecionada
        df_base_original = df_base_original[df_base_original['DATA_HORA'] == data_hora]
        if not df_base_original.empty:
            df_base_original = df_base_original.set_index(['DATA_HORA'])
        return df_base_original
    except Exception as e:
        st.error(f"Erro ao carregar dados de medição: {str(e)}")
        return pd.DataFrame()

# Multiselect para Cód. do Trafo/Alimentador
codigos = df_dados_tecnicos['Cód. do Trafo/Alimentador'].unique()
selecao = st.multiselect("Selecione os Cód. do Equipamento:", codigos)

#seleção unica de data e hora
col1, col2 = st.columns(2)
with col1:
    data = st.date_input("Selecione a data", value=pd.Timestamp('2024-01-01'))
with col2:
    hora = st.selectbox("Selecione a hora", options=range(24), format_func=lambda x: f"{x:02d}:00")
data_hora_coincidente = pd.Timestamp.combine(data, pd.Timestamp(f"{hora:02d}:00:00").time())

if st.button("Calcular"):
    try:
        # Carregar dados de medição apenas para a data/hora selecionada
        df_base = load_measurement_data(data_hora_coincidente)
        
        if df_base.empty:
            st.warning("Não há dados disponíveis para a data e hora selecionadas.")
            st.stop()
            
        # Criar DataFrame para armazenar todos os resultados
        todos_resultados = []
            
        # Processar cada equipamento selecionado
        for equipamento in selecao:
            try:
                # Verificar se o equipamento existe na tabela de atributos
                EAE = f"{equipamento}-EAE"
                EAR = f"{equipamento}-EAR"
                ERE = f"{equipamento}-ERE"
                ERR = f"{equipamento}-ERR"
                
                # Buscar descrições na tabela de atributos
                filtro = df_atributos_Dados["Codigo"] == EAE
                indices_encontrados = df_atributos_Dados.loc[filtro, "Codigo"].index
                
                if len(indices_encontrados) == 0:
                    st.warning(f"Equipamento {equipamento} não encontrado na tabela de atributos")
                    continue
                
                # Obter descrições
                indice_entrada = indices_encontrados[0]
                descricao = str(df_atributos_Dados.loc[indice_entrada, 'descricao'])
                
                # Verificar se as descrições existem no DataFrame
                if descricao not in df_base.columns:
                    st.warning(f"Coluna {descricao} não encontrada nos dados de medição")
                    continue
                
                # Obter outras descrições
                indice_saida = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == EAR, 'Codigo'].index[0]
                descricao_saida = str(df_atributos_Dados.loc[indice_saida, 'descricao'])
                
                indice_entrada_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERE, 'Codigo'].index[0]
                descricao_Q = str(df_atributos_Dados.loc[indice_entrada_Q, 'descricao'])
                
                indice_saida_Q = df_atributos_Dados.loc[df_atributos_Dados['Codigo'] == ERR, 'Codigo'].index[0]
                descricao_saida_Q = str(df_atributos_Dados.loc[indice_saida_Q, 'descricao'])
                
                # Criar DataFrame filtrado
                df_base_filtrado = pd.DataFrame(df_base, columns=['DATA_HORA', descricao, descricao_saida, descricao_Q, descricao_saida_Q])
                df_base_filtrado.fillna(0.0, inplace=True)
                
                # Converter para float
                for col in [descricao, descricao_saida, descricao_Q, descricao_saida_Q]:
                    df_base_filtrado[col] = df_base_filtrado[col].astype(float)
                
                # Calcular potências
                df_base_filtrado['P'] = df_base_filtrado[descricao] - df_base_filtrado[descricao_saida]
                df_base_filtrado['Q'] = df_base_filtrado[descricao_Q] - df_base_filtrado[descricao_saida_Q]
                df_base_filtrado['S'] = np.sqrt((df_base_filtrado['P'] ** 2) + (df_base_filtrado['Q'] ** 2))
                
                # Obter potência instalada
                potencia_instalada = df_dados_tecnicos.loc[
                    df_dados_tecnicos['Cód. do Trafo/Alimentador'] == equipamento,
                    'Potencia Instalada'
                ].values[0]
                
                if pd.isna(potencia_instalada) or potencia_instalada == 0:
                    potencia_instalada = 1
                
                # Calcular carregamento
                carregamento = (df_base_filtrado['S'].max() / potencia_instalada) * 100
                
                # Adicionar resultados à lista
                todos_resultados.append({
                    'Equipamento': equipamento,
                    'Potência Ativa (kW)': '{:,.2f}'.format(df_base_filtrado['P'].max()).replace(',', 'X').replace('.', ',').replace('X', '.'),
                    'Potência Reativa (kvar)': '{:,.2f}'.format(df_base_filtrado['Q'].max()).replace(',', 'X').replace('.', ',').replace('X', '.'),
                    'Potência Aparente (kVA)': '{:,.2f}'.format(df_base_filtrado['S'].max()).replace(',', 'X').replace('.', ',').replace('X', '.'),
                    'Carregamento (%)': '{:,.2f}'.format(carregamento).replace(',', 'X').replace('.', ',').replace('X', '.')
                })
                
            except Exception as e:
                st.error(f"Erro ao processar equipamento {equipamento}: {str(e)}")
                continue
        
        # Exibir todos os resultados em uma única tabela
        if todos_resultados:
            resultados_df = pd.DataFrame(todos_resultados)
            st.table(resultados_df)          
            # Exportação para Excel
            output = io.BytesIO()
            resultados_df.to_excel(output, index=False, sheet_name='Resultados')
            
            # Criar diretório exportado se não existir
            export_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'exportado')
            os.makedirs(export_dir, exist_ok=True)
            
            # Salvar arquivo no diretório exportado
            export_path = os.path.join(export_dir, f"demanda_coincidente_{data_hora_coincidente.strftime('%Y%m%d_%H%M')}.xlsx")
            resultados_df.to_excel(export_path, index=False, sheet_name='Resultados')
            
            # Botão de download
            st.download_button(
                label="Exportar Dados (Excel)",
                data=output.getvalue(),
                file_name=f"demanda_coincidente_{data_hora_coincidente.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum resultado foi calculado.")
                
    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")

