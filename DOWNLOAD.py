import os
import streamlit as st
import pandas as pd
import base64
import tempfile

# Diretório onde os arquivos Excel estão localizados
diretorio_excel = "data"  # Caminho relativo ao diretório do script

# Lista para armazenar os caminhos dos arquivos Excel
arquivos_excel = []

# Itera sobre os arquivos no diretório
for root, dirs, files in os.walk(diretorio_excel):
    for file in files:
        if file.endswith(".xlsx"):
            arquivos_excel.append(os.path.join(root, file))

# Exibindo a lista de arquivos Excel
if arquivos_excel:
    st.header('Lista de Arquivos Excel Disponíveis para Download:')
    for arquivo in arquivos_excel:
        st.write(f"- {os.path.basename(arquivo)}")

    # Seleção do arquivo para download
    arquivo_selecionado = st.selectbox('Selecione um arquivo para download:', arquivos_excel)

    # Verificar se o usuário clicou no botão de download
    if st.button('Download Excel'):
        try:
            # Leitura do DataFrame a partir do arquivo selecionado
            df = pd.read_excel(arquivo_selecionado)

            # Criando um arquivo temporário
            temp_dir = tempfile.mkdtemp()
            temp_excel_file = os.path.join(temp_dir, "temp_excel_file.xlsx")
            df.to_excel(temp_excel_file, index=False)

            # Download do arquivo Excel
            with open(temp_excel_file, "rb") as f:
                excel_data = f.read()
                b64 = base64.b64encode(excel_data).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(arquivo_selecionado)}">Baixar Dados</a>'
                st.markdown(href, unsafe_allow_html=True)

            # Removendo o diretório temporário
            os.remove(temp_excel_file)
            os.rmdir(temp_dir)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Excel: {e}")
else:
    st.warning('Nenhum arquivo Excel encontrado no diretório especificado.')
