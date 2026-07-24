import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Gerador de Propostas - Leves", layout="wide")

st.title("📦 Gerador de Propostas e Movimentação de Leves")
st.markdown("Bem-vindo! Faça o upload das bases extraídas do Looker para começar.")

# --- BARRA LATERAL: UPLOADS E LINKS LOOKER ---
st.sidebar.header("1. Upload de Bases de Dados")

st.sidebar.markdown("**Tabelas frete peso praticadas**")
st.sidebar.markdown("[Link Looker: 26300](https://loggi.looker.com/looks/26300)")
file_frete = st.sidebar.file_uploader("Upload Frete Peso", type=["xlsx", "csv"], label_visibility="collapsed")

st.sidebar.markdown("**Abrangências atuais**")
st.sidebar.markdown("[Link Looker: 26301](https://loggi.looker.com/looks/26301)")
file_abrangencia = st.sidebar.file_uploader("Upload Abrangência", type=["xlsx", "csv"], label_visibility="collapsed")

st.sidebar.markdown("**SLOs globais das cidades**")
st.sidebar.markdown("[Link Looker: 26303](https://loggi.looker.com/looks/26303)")
file_slos = st.sidebar.file_uploader("Upload SLOs", type=["xlsx", "csv"], label_visibility="collapsed")

st.sidebar.markdown("**Volume de pacotes (30 dias)**")
st.sidebar.markdown("[Link Looker: 26302](https://loggi.looker.com/looks/26302)")
file_volume = st.sidebar.file_uploader("Upload Volume", type=["xlsx", "csv"], label_visibility="collapsed")

# --- FUNÇÕES DE TRATAMENTO DE DADOS ---
@st.cache_data
def load_data(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    return pd.read_excel(file)

@st.cache_data
def processar_frete(df_frete):
    df_recentes = df_frete.drop_duplicates(subset=['LMC name'], keep='first')
    tabelas_validas = df_recentes[['LMC name', 'table name']]
    df_filtrado = df_frete.merge(tabelas_validas, on=['LMC name', 'table name'], how='inner')
    return df_filtrado

@st.cache_data
def processar_slos(df_slos):
    df_slos['prioridade'] = np.where(df_slos['Tipo de serviço'].str.contains('express', case=False, na=False), 1, 2)
    df_slos = df_slos.sort_values(by=['Cidade', 'prioridade', 'SLO'], ascending=[True, True, True])
    df_slos_clean = df_slos.drop_duplicates(subset=['Cidade'], keep='first')
    return df_slos_clean

@st.cache_data
def processar_nomes_leves(df_volume):
    mapping = df_volume[['Leve', 'Routing Code']].drop_duplicates().dropna()
    mapping['nome_completo'] = mapping['Leve'] + " (" + mapping['Routing Code'] + ")"
    return mapping

def formatar_moeda(valor):
    try:
        # Formata para 2 casas decimais e troca ponto por vírgula no padrão BR
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return valor

# --- FLUXO PRINCIPAL ---
if file_frete and file_abrangencia and file_slos and file_volume:
    
    with st.spinner("Carregando e processando bases..."):
        df_frete = load_data(file_frete)
        df_abrangencia = load_data(file_abrangencia)
        df_slos = load_data(file_slos)
        df_volume = load_data(file_volume)
        
        df_frete_clean = processar_frete(df_frete)
        df_slos_clean = processar_slos(df_slos)
        df_nomes_leves = processar_nomes_leves(df_volume)
        
    st.success("Bases carregadas com sucesso!")
    st.divider()

    st.header("2. Seleção de Leves")
    
    # Prepara lista de leves e dicionários de mapeamento
    leves_disponiveis = df_frete_clean['LMC name'].dropna().unique().tolist()
    
    mapa_nomes = {}
    mapa_routing = {} # Novo dicionário para pegar a sigla (Routing Code)
    
    for lmc in leves_disponiveis:
        match = df_nomes_leves[df_nomes_leves['Leve'] == lmc]
        if not match.empty:
            nome_formatado = match['nome_completo'].values[0]
            routing_code = match['Routing Code'].values[0]
            mapa_nomes[nome_formatado] = lmc
            mapa_routing[lmc] = routing_code
        else:
            mapa_nomes[lmc] = lmc
            mapa_routing[lmc] = "-"
            
    lista_nomes_exibicao = list(mapa_nomes.keys())

    leves_selecionados_formatados = st.multiselect("Selecione os Leves envolvidos na negociação:", lista_nomes_exibicao)
    
    leves_selecionados = [mapa_nomes[nome] for nome in leves_selecionados_formatados]

    if leves_selecionados:
        st.divider()
        st.header("3. Definição das Cidades Base")
        
        cidades_disponiveis = df_slos_clean['Cidade'].dropna().unique().tolist()
        cidades_base_dict = {}
        
        cols = st.columns(len(leves_selecionados))
        
        for idx, leve in enumerate(leves_selecionados):
            nome_exibicao = leves_selecionados_formatados[idx]
            with cols[idx]:
                cidade = st.selectbox(f"Cidade Base para:\n{nome_exibicao}", cidades_disponiveis, key=f"cidade_{leve}")
                cidades_base_dict[leve] = cidade

        st.divider()
        st.header("4. Dados Atuais dos Leves Selecionados")
        
        df_abrangencia.rename(columns={'Região de preço 2023': 'Região de preço'}, inplace=True)

        tab1, tab2 = st.tabs(["Tabela Frete Peso Atual", "Abrangência e Prazos"])
        
        with tab1:
            for idx, leve in enumerate(leves_selecionados):
                nome_exibicao = leves_selecionados_formatados[idx]
                st.subheader(f"Tabela Frete Peso: {nome_exibicao}")
                
                df_frete_leve = df_frete_clean[df_frete_clean['LMC name'] == leve].copy()
                
                # Adiciona o Routing Code
                df_frete_leve['Routing Code'] = mapa_routing.get(leve, "-")
                
                # Renomeia colunas
                df_frete_leve.rename(columns={
                    'label': 'Região de preço',
                    'on time amount': 'Valor do pacote dentro do prazo',
                    'out of time amount': 'Valor do pacote fora do prazo'
                }, inplace=True)
                
                # Formata como moeda
                df_frete_leve['Valor do pacote dentro do prazo'] = df_frete_leve['Valor do pacote dentro do prazo'].apply(formatar_moeda)
                df_frete_leve['Valor do pacote fora do prazo'] = df_frete_leve['Valor do pacote fora do prazo'].apply(formatar_moeda)
                
                # Define a ordem das colunas
                colunas_frete_exibicao = [
                    'LMC name', 'Routing Code', 'Região de preço', 'Faixa de peso (g/m³)',
                    'Valor do pacote dentro do prazo', 'Valor do pacote fora do prazo', 'table name'
                ]
                
                # Filtra apenas as colunas que existem para evitar quebra caso alguma falte
                colunas_presentes_frete = [col for col in colunas_frete_exibicao if col in df_frete_leve.columns]
                
                st.dataframe(df_frete_leve[colunas_presentes_frete], use_container_width=True, hide_index=True)
                st.markdown("<br>", unsafe_allow_html=True)
            
        with tab2:
            for idx, leve in enumerate(leves_selecionados):
                nome_exibicao = leves_selecionados_formatados[idx]
                cidade_base = cidades_base_dict[leve]
                
                st.subheader(f"Abrangência e Prazos: {nome_exibicao}")
                
                df_abrangencia_leve = df_abrangencia[df_abrangencia['LMC Name'] == leve].copy()
                
                # Adiciona o Routing Code
                df_abrangencia_leve['Routing Code'] = mapa_routing.get(leve, "-")
                df_abrangencia_leve['SLO Local (Arquivo)'] = df_abrangencia_leve['Prazo adicional']
                
                # Define a ordem das colunas
                colunas_abrangencia_exibicao = ['LMC Name', 'Routing Code', 'Região de preço', 'Cidade', 'State', 'SLO Local (Arquivo)']
                colunas_presentes_abrangencia = [col for col in colunas_abrangencia_exibicao if col in df_abrangencia_leve.columns]
                
                st.dataframe(df_abrangencia_leve[colunas_presentes_abrangencia], use_container_width=True, hide_index=True)
                st.markdown("<br>", unsafe_allow_html=True)

else:
    st.info("Por favor, faça o upload de **todas as 4 bases** na barra lateral para prosseguir.")