import math
import time
import datetime
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")
import matplotlib.pyplot as plt
import csv
import numpy as np

# Constantes
FP_LIMITE = 0.92
FP_ALVO = 0.98
ANGULO_CORRECAO = 11.8
FATOR_SOBRECARGA = 1.00
DESVIO_PADRAO_MAX = 3.0
DESVIO_PADRAO_MIN = 2.0
TOLERANCIA_MAXIMO = 1.10

class AnalisadorFatorPotencia:
    def __init__(self, caminho_base, caminho_atributos):
        """
        Inicializa o analisador de fator de potência.
        
        Args:
            caminho_base (str): Caminho para o arquivo CSV com dados base
            caminho_atributos (str): Caminho para o arquivo Excel com atributos
        """
        self.caminho_base = caminho_base
        self.caminho_atributos = caminho_atributos
        self.df_base = None
        self.df_atributos = None
        self.df_atributos_Dados = None
        self.resultados = []
        
    def carregar_dados(self):
        """Carrega os dados dos arquivos de entrada."""
print("Importando Base de Dados Agrupada\n")
        self.df_base = pd.read_csv(self.caminho_base, sep=";", encoding='latin-1')
        self.df_base['DATA_HORA'] = pd.to_datetime(self.df_base['DATA_HORA'])
        self.df_base = self.df_base.set_index(['DATA_HORA'])
        self.df_base['MES'] = self.df_base.index.month
        self.df_base['ANO'] = self.df_base.index.year
        self.df_base['MES'] = self.df_base['MES'].astype(int)
        self.df_base['ANO'] = self.df_base['ANO'].astype(int)
print("\nImportação Concluída")

        self.df_atributos_Dados = pd.read_excel(self.caminho_atributos, sheet_name="Dados")
        self.df_atributos_Dados.dropna(subset=['Codigo'], inplace=True)

        self.df_atributos = pd.read_excel(self.caminho_atributos, sheet_name="Dados Técnicos")
        self.df_atributos.dropna(subset=['Cód. do Trafo/Alimentador'], inplace=True)

    def calcular_fator_potencia(self, df_filtrado, descricao, descricao_saida, descricao_Q, descricao_saida_Q):
        """
        Calcula o fator de potência para um conjunto de dados filtrado.
        
        Args:
            df_filtrado (DataFrame): DataFrame com dados filtrados
            descricao (str): Nome da coluna de entrada
            descricao_saida (str): Nome da coluna de saída
            descricao_Q (str): Nome da coluna de potência reativa de entrada
            descricao_saida_Q (str): Nome da coluna de potência reativa de saída
            
        Returns:
            DataFrame: DataFrame com cálculos de fator de potência
        """
        df_filtrado[descricao] = df_filtrado[descricao].astype(float)
        df_filtrado[descricao_saida] = df_filtrado[descricao_saida].astype(float)
        df_filtrado[descricao_Q] = df_filtrado[descricao_Q].astype(float)
        df_filtrado[descricao_saida_Q] = df_filtrado[descricao_saida_Q].astype(float)

        df_filtrado['P'] = df_filtrado[descricao] - df_filtrado[descricao_saida]
        df_filtrado['Q'] = df_filtrado[descricao_Q] - df_filtrado[descricao_saida_Q]
        df_filtrado['S'] = np.sqrt((df_filtrado['P'] ** 2) + (df_filtrado['Q'] ** 2))
        df_filtrado['fp'] = df_filtrado.apply(lambda row: row['P'] / row['S'] if row['P'] > 0 else -row['P'] / row['S'], axis=1)
        
        return df_filtrado

    def calcular_estatisticas_fp(self, df_filtrado):
        """
        Calcula estatísticas relacionadas ao fator de potência.
        
        Args:
            df_filtrado (DataFrame): DataFrame com dados filtrados
            
        Returns:
            dict: Dicionário com as estatísticas calculadas
        """
        df_filtrado['fp - condicao'] = pd.cut(df_filtrado['fp'],
                                           bins=[-float('inf'), FP_LIMITE, float('inf')],
                                           labels=['abaixo de 0.92', 'acima de 0.92'])
        
        contagem = df_filtrado['fp - condicao'].value_counts().to_dict()
                contagem_abaixo_092 = contagem.get('abaixo de 0.92', 0)
                contagem_acima_092 = contagem.get('acima de 0.92', 0)

        df_abaixo_092 = df_filtrado[df_filtrado['fp'] < FP_LIMITE]
        df_acima_092 = df_filtrado[df_filtrado['fp'] > FP_LIMITE]

                media_fp_abaixo_092 = df_abaixo_092['fp'].mean()
                media_fp_acima_092 = df_acima_092['fp'].mean()

        return {
            'contagem_abaixo_092': contagem_abaixo_092,
            'contagem_acima_092': contagem_acima_092,
            'media_fp_abaixo_092': media_fp_abaixo_092,
            'media_fp_acima_092': media_fp_acima_092
        }

    def calcular_correcao_fp(self, df_filtrado):
        """
        Calcula a correção necessária para o fator de potência.
        
        Args:
            df_filtrado (DataFrame): DataFrame com dados filtrados
            
        Returns:
            float: Valor da correção necessária
        """
        df_filtrado['S_corrigido'] = df_filtrado['P'] / FP_ALVO
        angulo_radianos = math.radians(ANGULO_CORRECAO)
        seno = math.sin(angulo_radianos)
        df_filtrado['Q_correção'] = df_filtrado['Q'] - (df_filtrado['S_corrigido'] * seno)
        return df_filtrado['Q_correção'].mean()

    def calcular_sobrecarga(self, df_filtrado, potencia_instalada):
        """
        Calcula as horas de sobrecarga.
        
        Args:
            df_filtrado (DataFrame): DataFrame com dados filtrados
            potencia_instalada (float): Potência instalada do equipamento
            
        Returns:
            int: Número de horas de sobrecarga
        """
        df_filtrado['ultrapassagem'] = df_filtrado['S'] / potencia_instalada >= FATOR_SOBRECARGA
        df_filtrado['ultrapassagem'] = df_filtrado['ultrapassagem'].astype(int)
        df_filtrado["Ultrapassagem"] = df_filtrado.apply(lambda row: row["ultrapassagem"].sum(), axis=1)
        return df_filtrado["Ultrapassagem"].sum()

    def calcular_valores_maximos(self, df_filtrado, valor_media_p, desvio_padrao_P):
        """
        Calcula os valores máximos considerando desvios padrão.
        
        Args:
            df_filtrado (DataFrame): DataFrame com dados filtrados
            valor_media_p (float): Média da potência ativa
            desvio_padrao_P (float): Desvio padrão da potência ativa
            
        Returns:
            tuple: (valor_maximo_P, valor_Q_correspondente, potencia_aparente_maxima)
        """
        minimo_desvio = round(valor_media_p - DESVIO_PADRAO_MIN * desvio_padrao_P, 0)
        maximo_desvio = round((valor_media_p + DESVIO_PADRAO_MAX * desvio_padrao_P), 0)

        base_filtrada = np.where(df_filtrado['P'] < valor_media_p + DESVIO_PADRAO_MAX * desvio_padrao_P, 
                               df_filtrado['P'], np.nan)
                valor_maximo_filtrado = np.nanmax(base_filtrada)
        valor_maximo_P = max(df_filtrado['P'])

        if valor_maximo_P / valor_maximo_filtrado > TOLERANCIA_MAXIMO:
                      valor_maximo_P = valor_maximo_filtrado

        max_date_time_index = df_filtrado.index[df_filtrado['P'] == valor_maximo_P].tolist()

                if max_date_time_index and not pd.isna(max_date_time_index[0]):
            valor_Q_correspondente = df_filtrado.at[max_date_time_index[0], 'Q']
                else:
            valor_Q_correspondente = 0

                potencia_aparente_maxima = np.sqrt((valor_maximo_P ** 2) + (valor_Q_correspondente ** 2))

        return valor_maximo_P, valor_Q_correspondente, potencia_aparente_maxima

    def processar_dados(self, ano, mes):
        """
        Processa os dados para um determinado ano e mês.
        
        Args:
            ano (int): Ano a ser processado
            mes (int): Mês a ser processado
        """
        df_filtrado_ano = self.df_base[self.df_base["ANO"] == ano]
        if mes not in df_filtrado_ano['MES'].values:
            return

        df_filtrado_mes = df_filtrado_ano[df_filtrado_ano["MES"] == mes]
        
        for selecao in self.df_atributos['Cód. do Trafo/Alimentador']:
            if selecao == 0:
                    continue

            try:
                potencia_instalada = self.df_atributos.loc[
                    self.df_atributos['Cód. do Trafo/Alimentador'] == selecao,
                    'Potencia Instalada'
                ].values[0]
            except IndexError:
                potencia_instalada = 0

            # Processamento dos dados para cada alimentador
            EAE = str(selecao) + '-EAE'
            EAR = str(selecao) + '-EAR'
            ERE = str(selecao) + '-ERE'
            ERR = str(selecao) + '-ERR'

            filtro = self.df_atributos_Dados["Codigo"] == EAE
            indices_encontrados = self.df_atributos_Dados.loc[filtro, "Codigo"].index

            if len(indices_encontrados) > 0:
                indice_entrada = indices_encontrados[0]
                descricao = self.df_atributos_Dados.loc[indice_entrada, 'descricao']
            else:
                print("Chave não encontrada em df_atributos_Dados:", EAE)
                continue

            descricao = str(descricao)
            alimentador = descricao.replace("-EAE", "")
            indice_saida = self.df_atributos_Dados.loc[self.df_atributos_Dados['Codigo'] == EAR, 'Codigo'].index[0]
            descricao_saida = self.df_atributos_Dados.loc[indice_saida, 'descricao']
            descricao_saida = str(descricao_saida)
            indice_entrada_Q = self.df_atributos_Dados.loc[self.df_atributos_Dados['Codigo'] == ERE, 'Codigo'].index[0]
            descricao_Q = self.df_atributos_Dados.loc[indice_entrada_Q, 'descricao']
            descricao_Q = str(descricao_Q)
            indice_saida_Q = self.df_atributos_Dados.loc[self.df_atributos_Dados['Codigo'] == ERR, 'Codigo'].index[0]
            descricao_saida_Q = self.df_atributos_Dados.loc[indice_saida_Q, 'descricao']
            descricao_saida_Q = str(descricao_saida_Q)

            # Criar DataFrame filtrado
            df_base_filtrado = pd.DataFrame(df_filtrado_mes, columns=['DATA_HORA', descricao, descricao_saida, descricao_Q, descricao_saida_Q])
            df_base_filtrado.fillna(0.0)

            # Calcular fator de potência
            df_base_filtrado = self.calcular_fator_potencia(df_base_filtrado, descricao, descricao_saida, descricao_Q, descricao_saida_Q)

            # Calcular estatísticas
            estatisticas = self.calcular_estatisticas_fp(df_base_filtrado)
            
            # Calcular correção
            Q_correcao_media = self.calcular_correcao_fp(df_base_filtrado)
            
            # Calcular sobrecarga
            horas_sobrecarga = self.calcular_sobrecarga(df_base_filtrado, potencia_instalada)
            
            # Calcular valores máximos
            valor_maximo_P, valor_Q_correspondente, potencia_aparente_maxima = self.calcular_valores_maximos(
                df_base_filtrado, 
                np.mean(df_base_filtrado['P']), 
                np.nanstd(df_base_filtrado['P'])
            )

            # Determinar se é capacitivo ou indutivo
            valor_media_Q = np.mean(df_base_filtrado['Q'])
            valor_fp_médio = np.mean(df_base_filtrado['fp'])
            cap_indut = 'Capacitivo' if valor_media_Q < 0 else 'Indutivo'

            # Adicionar resultados
            self.resultados.append({
                'Cód. do Trafo/Alimentador': selecao,
                'Ano': ano,
                'Mês': mes,
                'Contagem_abaixo_0.92': estatisticas['contagem_abaixo_092'],
                'Contagem_acima_0.92': estatisticas['contagem_acima_092'],
                'Média_fp_abaixo_092': estatisticas['media_fp_abaixo_092'],
                'Média_fp_acima_092': estatisticas['media_fp_acima_092'],
                'BC necessário': Q_correcao_media,
                'valor_fp_médio': valor_fp_médio,
                'Capacitivo_ou_indutivo': cap_indut,
                'Horas Sobrecarga': horas_sobrecarga
            })

    def executar_analise(self, anos, meses):
        """
        Executa a análise completa para os anos e meses especificados.
        
        Args:
            anos (list): Lista de anos a serem analisados
            meses (list): Lista de meses a serem analisados
        """
        self.carregar_dados()
        
        for ano in anos:
            for mes in meses:
                self.processar_dados(ano, mes)
        
        return self.resultados

def main():
    """
    Função principal que executa a análise de fator de potência.
    """
    # Configuração dos caminhos dos arquivos
    caminho_base = r"C:\Users\Engeselt\Documents\GitHub\ASPO_2\Medição Agrupada.csv"
    caminho_atributos = r"C:\Users\Engeselt\Documents\GitHub\ASPO_2\Tabela informativa.xlsx"
    
    # Configuração dos anos e meses a serem analisados
    anos = range(2024, 2025)
    meses = range(1, 13)
    
    # Criação e execução do analisador
    analisador = AnalisadorFatorPotencia(caminho_base, caminho_atributos)
    resultados = analisador.executar_analise(anos, meses)
    
    # Salvamento dos resultados
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel("Resultados_Fator_Potencia.xlsx", index=False)
    
    print("Análise concluída. Resultados salvos em 'Resultados_Fator_Potencia.xlsx'")

if __name__ == "__main__":
    main()
