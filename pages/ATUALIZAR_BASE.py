import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import numpy as np
import sqlite3 as sql
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import os
import zipfile
import datetime
import logging
from datetime import datetime
from github import Github
import base64

# Configurar o logging
def setup_logger():
    # Configurar o logger
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    return logging.getLogger()

def atualizar_github(arquivo_path, mensagem_commit):
    """Atualiza o arquivo no GitHub"""
    try:
        # Configurar o repositório do GitHub
        if 'github_repo' not in st.secrets:
            st.error("Repositório do GitHub não configurado. Configure em .streamlit/secrets.toml")
            return False
            
        repo_name = st.secrets['github_repo']
        
        # Inicializar o cliente do GitHub
        g = Github(st.secrets['github_token'])
        repo = g.get_repo(repo_name)
        
        # Ler o conteúdo do arquivo
        with open(arquivo_path, 'r', encoding='latin-1') as file:
            content = file.read()
        
        # Obter o SHA do arquivo atual
        try:
            contents = repo.get_contents(arquivo_path)
            sha = contents.sha
        except:
            sha = None
        
        # Atualizar o arquivo
        repo.update_file(
            path=arquivo_path,
            message=mensagem_commit,
            content=content,
            sha=sha
        )
        
        return True
    except Exception as e:
        st.error(f"Erro ao atualizar o GitHub: {str(e)}")
        return False

st.subheader("Atualizar Base de Dados")

def processar_arquivo_zip(uploaded_file, logger):
    """Processa um único arquivo ZIP"""
    try:
        logger.info(f"Processando arquivo ZIP: {uploaded_file.name}")
        
        # Ler o conteúdo do ZIP diretamente da memória
        zip_content = uploaded_file.read()
        zip_file = io.BytesIO(zip_content)
        
        # Lista para armazenar os DataFrames
        lista_dataFrame = []
        
        # Processar arquivos do ZIP
        with zipfile.ZipFile(zip_file) as zip_ref:
            # Listar arquivos no ZIP
            arquivos = zip_ref.namelist()
            logger.info(f"Arquivos encontrados no ZIP: {arquivos}")
            
            if not arquivos:
                logger.error("ZIP está vazio")
                st.error("O arquivo ZIP está vazio")
                return None
            
            # Processar cada arquivo
            for arquivo in arquivos:
                if arquivo.endswith('.txt') or arquivo.endswith('.csv'):
                    try:
                        # Ler o conteúdo do arquivo
                        with zip_ref.open(arquivo) as f:
                            # Determinar a extensão correta
                            extensao = '.csv' if arquivo.endswith('.txt') else '.csv'
                            nome_arquivo = arquivo.replace('.txt', extensao) if arquivo.endswith('.txt') else arquivo
                            
                            logger.info(f"Processando arquivo: {nome_arquivo}")
                            
                            # Ler o CSV diretamente da memória
                            df = pd.read_csv(f, sep=';', encoding='latin-1')
                            logger.info(f"Arquivo {nome_arquivo} lido com {len(df)} linhas")
                            
                            # Verificar colunas
                            logger.info(f"Colunas encontradas: {df.columns.tolist()}")
                            
                            # Verificar se o DataFrame tem dados
                            if len(df) > 0:
                                # Verificar se tem a coluna DATA/Hora
                                if 'DATA/Hora' not in df.columns:
                                    logger.error(f"Arquivo {nome_arquivo} não contém a coluna 'DATA/Hora'")
                                    st.error(f"Arquivo {nome_arquivo} não contém a coluna 'DATA/Hora'")
                                    continue
                                
                                # Verificar valores únicos na coluna DATA/Hora
                                logger.info(f"Valores únicos em DATA/Hora: {df['DATA/Hora'].unique()}")
                                
                                lista_dataFrame.append(df)
                                logger.info(f"Arquivo {nome_arquivo} processado com sucesso")
                                st.success(f"Arquivo {nome_arquivo} lido com sucesso!")
                            else:
                                logger.warning(f"Arquivo {nome_arquivo} está vazio")
                                st.warning(f"Arquivo {nome_arquivo} está vazio")
                    except Exception as e:
                        logger.error(f"Erro ao processar arquivo {arquivo}: {str(e)}")
                        st.error(f"Erro ao processar arquivo {arquivo}: {str(e)}")
                        continue
        
        if lista_dataFrame:
            logger.info(f"Total de DataFrames processados: {len(lista_dataFrame)}")
            # Verificar o total de linhas em todos os DataFrames
            total_linhas = sum(len(df) for df in lista_dataFrame)
            logger.info(f"Total de linhas em todos os DataFrames: {total_linhas}")
            return lista_dataFrame
        else:
            logger.warning(f"Nenhum arquivo CSV ou TXT encontrado em {uploaded_file.name}")
            st.warning(f"Nenhum arquivo CSV ou TXT encontrado em {uploaded_file.name}")
            return None
            
    except Exception as e:
        logger.error(f"Erro ao processar ZIP {uploaded_file.name}: {str(e)}")
        st.error(f"Erro ao processar ZIP {uploaded_file.name}: {str(e)}")
        return None

def processar_e_atualizar_base():
    logger = setup_logger()
    logger.info("Iniciando processo de atualização da base de dados")
    
    try:
        # Upload de múltiplos arquivos zip
        uploaded_files = st.file_uploader("Escolha os arquivos ZIP contendo os dados", type="zip", accept_multiple_files=True)
        
        if uploaded_files:
            logger.info(f"Recebidos {len(uploaded_files)} arquivos ZIP")
            
            # Lista para armazenar todos os DataFrames
            todos_dataframes = []
            
            # Processar cada arquivo ZIP
            for uploaded_file in uploaded_files:
                dataframes = processar_arquivo_zip(uploaded_file, logger)
                if dataframes:
                    todos_dataframes.extend(dataframes)
            
            # Processar os dados se houver DataFrames
            if todos_dataframes:
                logger.info(f"Iniciando processamento dos dados. Total de DataFrames: {len(todos_dataframes)}")
                df_principal = pd.concat(todos_dataframes)
                logger.info(f"DataFrame principal criado com {len(df_principal)} linhas")
                
                # Limpar dados
                logger.info("Iniciando limpeza dos dados")
                df_principal = df_principal[df_principal['DATA/Hora'] != 'TOTAL']
                logger.info(f"Dados após remover TOTAL: {len(df_principal)} linhas")
                
                # Verificar valores antes da conversão
                logger.info(f"Valores únicos em DATA/Hora antes da conversão: {df_principal['DATA/Hora'].unique()}")
                
                # Primeiro converter para datetime sem especificar formato
                df_principal['DATA/Hora'] = pd.to_datetime(df_principal['DATA/Hora'], errors='coerce')
                
                # Converter para o formato da base existente (DD/MM/YYYY HH:MM:SS)
                df_principal['DATA/Hora'] = df_principal['DATA/Hora'].dt.strftime('%d/%m/%Y %H:%M:%S')
                
                # Agora converter para datetime com o formato correto
                df_principal['DATA/Hora'] = pd.to_datetime(df_principal['DATA/Hora'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
                
                # Removendo linhas com datas inválidas
                df_principal = df_principal.dropna(subset=['DATA/Hora'])
                logger.info(f"Dados após limpeza de datas: {len(df_principal)} linhas")
                
                if len(df_principal) == 0:
                    logger.error("Todos os dados foram perdidos durante a limpeza de datas")
                    st.error("Erro: Todos os dados foram perdidos durante a limpeza de datas. Verifique o formato das datas nos arquivos.")
                    return
                
                # Processar datas e horas
                df_principal['DATA'] = df_principal['DATA/Hora'].dt.date
                df_principal['HORA'] = df_principal['DATA/Hora'].dt.hour
                df_principal['HORA'] = df_principal['HORA'].apply(lambda x: f'{x:02}')
                df_principal['DATA_HORA'] = pd.to_datetime(df_principal['DATA'].astype(str) + ' ' + df_principal['HORA'].astype(str), format='%Y-%m-%d %H')
                df_principal.set_index('DATA_HORA', inplace=True)
                logger.info("Dados limpos e formatados")
                
                # Limpar colunas numéricas
                logger.info("Iniciando limpeza das colunas numéricas")
                df_principal.drop('DATA/Hora', axis=1, inplace=True)
                colunas_numericas = df_principal.columns[1:-1]
                df_principal[colunas_numericas] = df_principal[colunas_numericas].replace(',', '.', regex=True).replace('', '0')
                df_principal[colunas_numericas] = df_principal[colunas_numericas].replace(' ', '0')
                df_principal[colunas_numericas] = df_principal[colunas_numericas].fillna('')
                df_principal[colunas_numericas] = df_principal[colunas_numericas].replace(' ', '', regex=True)
                df_principal[colunas_numericas] = df_principal[colunas_numericas].replace('', np.nan)
                logger.info("Colunas numéricas limpas")
                
                # Converter colunas para numérico
                logger.info("Convertendo colunas para numérico")
                for coluna in df_principal.columns:
                    df_principal[coluna] = pd.to_numeric(df_principal[coluna], errors='coerce').fillna(0).astype(int)
                logger.info("Conversão numérica concluída")
                
                # Agrupar dados
                logger.info("Iniciando agrupamento dos dados")
                df_agrupado = df_principal.groupby('DATA_HORA').sum()
                df_agrupado.to_csv("Medição Agrupada_diária.csv", sep=";", encoding='latin-1')
                logger.info(f"Dados agrupados: {len(df_agrupado)} linhas")
                
                # Carregar base existente se disponível
                try:
                    logger.info("Tentando carregar base de dados existente")
                    df_existente = pd.read_csv("Medição Agrupada.csv", sep=";", encoding='latin-1')
                    logger.info(f"Base existente lida com {len(df_existente)} linhas")
                    
                    # Verificar valores antes da conversão
                    logger.info(f"Exemplo de valores em DATA_HORA antes da conversão: {df_existente['DATA_HORA'].head()}")
                    
                    # Convertendo a coluna DATA_HORA para datetime com tratamento de erros
                    df_existente['DATA_HORA'] = pd.to_datetime(df_existente['DATA_HORA'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
                    
                    # Verificar valores após a conversão
                    logger.info(f"Exemplo de valores em DATA_HORA após conversão: {df_existente['DATA_HORA'].head()}")
                    
                    # Verificar quantos valores NaT foram gerados
                    nat_count = df_existente['DATA_HORA'].isna().sum()
                    if nat_count > 0:
                        logger.warning(f"Encontrados {nat_count} valores de data inválidos")
                        # Mostrar exemplos de valores que não puderam ser convertidos
                        invalid_dates = df_existente[df_existente['DATA_HORA'].isna()]['DATA_HORA'].head()
                        logger.warning(f"Exemplos de datas inválidas: {invalid_dates}")
                    
                    # Remover linhas com datas inválidas antes de remover duplicatas
                    df_existente = df_existente.dropna(subset=['DATA_HORA'])
                    logger.info(f"Base existente após remover datas inválidas: {len(df_existente)} linhas")
                    
                    # Remover duplicatas antes de definir o índice
                    df_existente = df_existente.drop_duplicates(subset=['DATA_HORA'], keep='last')
                    logger.info(f"Base existente após remoção de duplicatas: {len(df_existente)} linhas")
                    
                    # Verificar valores únicos após remoção de duplicatas
                    logger.info(f"Exemplo de valores únicos em DATA_HORA após remoção de duplicatas: {df_existente['DATA_HORA'].unique()[:5]}")
                    
                    df_existente.set_index('DATA_HORA', inplace=True)
                    logger.info(f"Base existente carregada: {len(df_existente)} linhas")
                    
                    # Mesclar dados
                    logger.info("Iniciando mesclagem dos dados")
                    # Garantir que os índices estão no mesmo formato
                    df_agrupado.index = pd.to_datetime(df_agrupado.index)
                    df_existente.index = pd.to_datetime(df_existente.index)
                    
                    # Verificar índices antes da mesclagem
                    logger.info(f"Exemplo de índices em df_agrupado: {df_agrupado.index[:5]}")
                    logger.info(f"Exemplo de índices em df_existente: {df_existente.index[:5]}")
                    
                    # Remover duplicatas em ambos os DataFrames
                    df_agrupado = df_agrupado[~df_agrupado.index.duplicated(keep='last')]
                    df_existente = df_existente[~df_existente.index.duplicated(keep='last')]
                    
                    # Combinar os índices
                    todos_indices = df_agrupado.index.union(df_existente.index)
                    logger.info(f"Total de índices únicos: {len(todos_indices)}")
                    
                    # Reindexar ambos os DataFrames para ter o mesmo índice
                    df_existente = df_existente.reindex(todos_indices)
                    df_agrupado = df_agrupado.reindex(todos_indices)
                    
                    # Atualizar valores existentes com novos dados
                    df_existente.update(df_agrupado)
                    
                    # Preencher valores NaN com 0
                    df_existente = df_existente.fillna(0)
                    
                    df_final = df_existente
                    logger.info(f"Mesclagem concluída. Total de registros: {len(df_final)}")
                    
                    # Verificar se há dados após a mesclagem
                    if df_final.empty:
                        logger.warning("DataFrame final está vazio após a mesclagem")
                        st.warning("Não foi possível mesclar os dados. Verifique os arquivos de entrada.")
                        return
                        
                except Exception as e:
                    logger.warning(f"Base existente não encontrada ou erro ao carregar: {str(e)}")
                    df_final = df_agrupado
                    logger.info("Usando apenas os novos dados carregados")
                
                # Adicionar patamar de carga
                logger.info("Iniciando cálculo do patamar de carga")
                def determinar_patamar_de_carga(data):
                    mes = data.month
                    dia_semana = data.weekday()
                    hora = data.hour

                    if 5 <= mes <= 8:  # maio a agosto
                        if 0 <= dia_semana <= 4:  # Segunda a Sexta-feira
                            if 1 <= hora <= 8:
                                return "Leve"
                            elif 9 <= hora <= 15:
                                return "Média"
                            elif 23 <= hora <= 24 or hora == 0:
                                return "Média"
                            else:
                                return "Pesada"
                        else:  # Sábado, Domingo e Feriado
                            if 1 <= hora <= 18:
                                return "Leve"
                            else:
                                return "Média"
                    elif mes == 4 or mes == 9 or mes == 10:  # abril, setembro e outubro
                        if 0 <= dia_semana <= 4:  # Segunda a Sexta-feira
                            if 1 <= hora <= 8:
                                return "Leve"
                            elif 9 <= hora <= 14:
                                return "Média"
                            elif 23 <= hora <= 24 or hora == 0:
                                return "Média"
                            elif 19 <= hora <= 22:
                                return "Leve"
                            else:
                                return "Pesada"
                        else:  # Sábado, Domingo e Feriado
                            if 1 <= hora <= 18:
                                return "Leve"
                            elif 19 <= hora <= 22:
                                return "Média"
                            else:
                                return "Leve"
                    else:  # novembro a março
                        if 0 <= dia_semana <= 4:  # Segunda a Sexta-feira
                            if 1 <= hora <= 8:
                                return "Leve"
                            elif 9 <= hora <= 13:
                                return "Média"
                            elif 23 <= hora <= 24 or hora == 0:
                                return "Média"
                            else:
                                return "Pesada"
                        else:  # Sábado, Domingo e Feriado
                            if 1 <= hora <= 18:
                                return "Leve"
                            else:
                                return "Média"
                
                df_final['Patamar de Carga'] = df_final.index.to_series().apply(determinar_patamar_de_carga)
                logger.info("Patamar de carga calculado")
                
                # Filtrar dados até o dia atual
                logger.info("Filtrando dados até o dia atual")
                current_date = datetime.now().date()
                filtered_df = df_final[df_final.index.to_series().dt.date < current_date]
                logger.info(f"Dados filtrados: {len(filtered_df)} linhas")
                
                # Remover duplicatas
                logger.info("Removendo duplicatas")
                filtered_df = filtered_df[~filtered_df.index.duplicated(keep='first')]
                logger.info(f"Dados após remoção de duplicatas: {len(filtered_df)} linhas")
                
                # Salvar arquivo final
                logger.info("Salvando arquivo final")
                arquivo_path = "Medição Agrupada.csv"
                filtered_df.to_csv(arquivo_path, sep=";", encoding='latin-1')
                logger.info("Arquivo final salvo com sucesso")
                
                # Atualizar no GitHub
                mensagem_commit = f"Atualização automática da base de dados - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                if atualizar_github(arquivo_path, mensagem_commit):
                    st.success("Base de dados atualizada com sucesso e enviada para o GitHub!")
                else:
                    st.warning("Base de dados atualizada localmente, mas houve erro ao enviar para o GitHub.")
            
        else:
            logger.warning("Nenhum arquivo ZIP foi carregado")
            st.warning("Por favor, carregue um ou mais arquivos ZIP contendo os dados")
            
    except Exception as e:
        logger.error(f"Erro durante o processamento: {str(e)}")
        st.error(f"Erro ao processar os dados: {str(e)}")

# Executar a função principal
processar_e_atualizar_base()