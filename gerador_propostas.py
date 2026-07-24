import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Gerador de Propostas - Leves", layout="wide")

st.title("📦 Gerador de Propostas e Movimentação de Leves")
st.markdown("Bem-vindo! Faça o upload das bases extraídas do Looker para começar.")

# --- BARRA LATERAL: UPLOADS ---
st.sidebar.header("1. Upload de Bases de Dados")
file_frete = st.sidebar.file_uploader("Tabelas frete peso praticadas", type=["xlsx", "csv"])
file_abrangencia = st.sidebar.file_uploader("Abrangências atuais", type=["xlsx", "csv"])
file_slos = st.sidebar.file_uploader("SLOs globais das cidades", type=["xlsx", "csv"])
file_volume = st.sidebar.file_uploader("Volume de pacotes (30 dias)", type=["xlsx", "csv"])

# --- FUNÇÕES DE TRATAMENTO DE DADOS ---
@st.cache_data
def load_data(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    return pd.read_excel(file)

@st.cache_data
def processar_frete(df_frete):
    # A base já vem com a data mais recente no topo. 
    # Precisamos pegar apenas o 'table name' mais recente de cada 'LMC name'
    df_recentes = df_frete.drop_duplicates(subset=['LMC name'], keep='first')
    tabelas_validas = df_recentes[['LMC name', 'table name']]
    
    # Fazemos um merge para manter apenas as linhas que correspondem ao table name mais recente
    df_filtrado = df_frete.merge(tabelas_validas, on=['LMC name', 'table name'], how='inner')
    return df_filtrado

@st.cache_data
def processar_slos(df_slos):
    # Regra: Se tem Loggi express, pega o menor SLO. Ignora Loggi Hoje/Redespacho.
    # Se só tem um serviço, pega ele.
    
    # 1. Filtramos apenas 'Loggi express' e os que são os únicos serviços da cidade
    # Para simplificar na Fase 1, vamos ordenar para dar preferência ao Loggi express e depois pegar o menor SLO
    
    # Criamos uma coluna de prioridade para o serviço (Express = 1, outros = 2)
    df_slos['prioridade'] = np.where(df_slos['Tipo de serviço'].str.contains('express', case=False, na=False), 1, 2)
    
    # Ordenamos por Cidade, Prioridade do serviço e SLO (do menor pro maior)
    df_slos = df_slos.sort_values(by=['Cidade', 'prioridade', 'SLO'], ascending=[True, True, True])
    
    # Drop duplicates pela cidade, mantendo o primeiro (que será o express com menor SLO, ou o único disponível)
    df_slos_clean = df_slos.drop_duplicates(subset=['Cidade'], keep='first')
    return df_slos_clean

# --- FLUXO PRINCIPAL ---
if file_frete and file_abrangencia and file_slos:
    
    with st.spinner("Carregando e processando bases..."):
        df_frete = load_data(file_frete)
        df_abrangencia = load_data(file_abrangencia)
        df_slos = load_data(file_slos)
        
        df_frete_clean = processar_frete(df_frete)
        df_slos_clean = processar_slos(df_slos)
        
    st.success("Bases carregadas com sucesso!")
    st.divider()

    st.header("2. Seleção de Leves e Parâmetros")
    
    col1, col2 = st.columns(2)
    
    with col1:
        leves_disponiveis = df_frete_clean['LMC name'].dropna().unique().tolist()
        leves_selecionados = st.multiselect("Selecione os Leves envolvidos na negociação:", leves_disponiveis)
    
    with col2:
        cidades_disponiveis = df_slos_clean['Cidade'].dropna().unique().tolist()
        cidade_base = st.selectbox("Qual é a Cidade Base para cálculo do SLO Local?", cidades_disponiveis)

    if leves_selecionados and cidade_base:
        st.divider()
        st.header("3. Dados Atuais dos Leves Selecionados")
        
        # Filtra os dados dos Leves selecionados
        df_frete_leves = df_frete_clean[df_frete_clean['LMC name'].isin(leves_selecionados)]
        df_abrangencia_leves = df_abrangencia[df_abrangencia['LMC Name'].isin(leves_selecionados)]
        
        # --- CÁLCULO DO SLO LOCAL ---
        # SLO Local = SLO Global da Cidade atendida - SLO Local da Cidade base do Leve
        try:
            slo_cidade_base = df_slos_clean[df_slos_clean['Cidade'] == cidade_base]['SLO'].values[0]
            
            # Traz o SLO Global para as cidades de abrangência
            df_abrangencia_leves = df_abrangencia_leves.merge(
                df_slos_clean[['Cidade', 'SLO']], on='Cidade', how='left'
            )
            df_abrangencia_leves.rename(columns={'SLO': 'SLO Global'}, inplace=True)
            
            # Calcula o SLO Local
            df_abrangencia_leves['SLO Local'] = df_abrangencia_leves['SLO Global'] - slo_cidade_base
            
            # Limita a 0 caso o cálculo dê negativo
            df_abrangencia_leves['SLO Local'] = df_abrangencia_leves['SLO Local'].apply(lambda x: x if x > 0 else 0)
            
        except IndexError:
            st.warning("Não foi possível encontrar a Cidade Base na tabela de SLOs Globais.")

        tab1, tab2 = st.tabs(["Tabela Frete Peso Atual", "Abrangência e Prazos"])
        
        with tab1:
            st.subheader("Tabela Frete Peso (Mais recente)")
            st.dataframe(df_frete_leves[['LMC name', 'table name', 'Faixa de peso (g/m³)', 'on time amount', 'out of time amount']], use_container_width=True)
            
        with tab2:
            st.subheader("Abrangência e SLO Local Calculado")
            colunas_exibicao = ['LMC Name', 'Região de preço 2023', 'Cidade', 'State', 'SLO Global', 'SLO Local']
            # Verifica se as colunas existem antes de exibir para evitar erros caso os nomes variem
            colunas_presentes = [col for col in colunas_exibicao if col in df_abrangencia_leves.columns]
            st.dataframe(df_abrangencia_leves[colunas_presentes], use_container_width=True)

else:
    st.info("Por favor, faça o upload de pelo menos as tabelas de **Frete**, **Abrangência** e **SLOs Globais** na barra lateral para prosseguir.")