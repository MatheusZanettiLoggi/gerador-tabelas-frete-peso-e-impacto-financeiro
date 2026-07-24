import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(page_title="Gerador de Propostas - Leves", layout="wide")

st.title("📦 Gerador de Propostas e Movimentação de Leves")
st.markdown("Bem-vindo! Faça o upload das bases extraídas do Looker para começar.")

# --- BARRA LATERAL: UPLOADS E LINKS LOOKER ---
with st.sidebar.expander("1. Upload de Bases de Dados", expanded=True):
    st.markdown("**Tabelas frete peso praticadas**")
    st.markdown("[Link Looker: 26300](https://loggi.looker.com/looks/26300)")
    file_frete = st.file_uploader("Upload Frete Peso", type=["xlsx", "csv"], label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Abrangências atuais**")
    st.markdown("[Link Looker: 26301](https://loggi.looker.com/looks/26301)")
    file_abrangencia = st.file_uploader("Upload Abrangência", type=["xlsx", "csv"], label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**SLOs globais das cidades**")
    st.markdown("[Link Looker: 26303](https://loggi.looker.com/looks/26303)")
    file_slos = st.file_uploader("Upload SLOs", type=["xlsx", "csv"], label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Volume de pacotes (30 dias)**")
    st.markdown("[Link Looker: 26302](https://loggi.looker.com/looks/26302)")
    file_volume = st.file_uploader("Upload Volume", type=["xlsx", "csv"], label_visibility="collapsed")

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
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return valor

def extrair_estado(nome_leve):
    match = re.search(r'-\s*([A-Z]{2})\b', nome_leve)
    if match:
        return match.group(1)
    return None

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
        
    st.sidebar.success("Bases carregadas com sucesso!")

    with st.expander("2. Seleção de Leves", expanded=True):
        leves_disponiveis = df_frete_clean['LMC name'].dropna().unique().tolist()
        
        mapa_nomes = {}
        mapa_routing = {} 
        
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
        
        with st.expander("3. Definição das Cidades Base", expanded=True):
            cidades_base_dict = {}
            cols = st.columns(len(leves_selecionados))
            
            for idx, leve in enumerate(leves_selecionados):
                nome_exibicao = leves_selecionados_formatados[idx]
                estado_do_leve = extrair_estado(leve)
                
                if estado_do_leve:
                    df_cidades_estado = df_slos_clean[df_slos_clean['State'] == estado_do_leve]
                else:
                    df_cidades_estado = df_slos_clean 
                    
                opcoes_cidades = {}
                for _, row in df_cidades_estado.iterrows():
                    cid = str(row['Cidade'])
                    est = str(row['State'])
                    txt_display = f"{cid} - {est}"
                    opcoes_cidades[txt_display] = cid
                    
                lista_opcoes_display = sorted(list(opcoes_cidades.keys()))
    
                with cols[idx]:
                    cidade_escolhida_display = st.selectbox(
                        f"Cidade Base para:\n{nome_exibicao}", 
                        lista_opcoes_display, 
                        key=f"cidade_{leve}"
                    )
                    cidades_base_dict[leve] = opcoes_cidades[cidade_escolhida_display]

        with st.expander("4. Dados Atuais dos Leves Selecionados", expanded=True):
            df_abrangencia.rename(columns={'Região de preço 2023': 'Região de preço'}, inplace=True)
    
            tab1, tab2 = st.tabs(["Tabela Frete Peso Atual", "Abrangência e Prazos"])
            
            with tab1:
                for idx, leve in enumerate(leves_selecionados):
                    nome_exibicao = leves_selecionados_formatados[idx]
                    st.subheader(f"Tabela Frete Peso: {nome_exibicao}")
                    
                    df_frete_leve = df_frete_clean[df_frete_clean['LMC name'] == leve].copy()
                    df_frete_leve['Routing Code'] = mapa_routing.get(leve, "-")
                    
                    df_frete_leve.rename(columns={
                        'label': 'Região de preço',
                        'on time amount': 'Valor do pacote dentro do prazo',
                        'out of time amount': 'Valor do pacote fora do prazo'
                    }, inplace=True)
                    
                    df_frete_leve['Valor do pacote dentro do prazo'] = df_frete_leve['Valor do pacote dentro do prazo'].apply(formatar_moeda)
                    df_frete_leve['Valor do pacote fora do prazo'] = df_frete_leve['Valor do pacote fora do prazo'].apply(formatar_moeda)
                    
                    colunas_frete_exibicao = [
                        'LMC name', 'Routing Code', 'Região de preço', 'Faixa de peso (g/m³)',
                        'Valor do pacote dentro do prazo', 'Valor do pacote fora do prazo', 'table name'
                    ]
                    
                    colunas_presentes_frete = [col for col in colunas_frete_exibicao if col in df_frete_leve.columns]
                    
                    st.dataframe(df_frete_leve[colunas_presentes_frete], use_container_width=True, hide_index=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                
            with tab2:
                for idx, leve in enumerate(leves_selecionados):
                    nome_exibicao = leves_selecionados_formatados[idx]
                    cidade_base = cidades_base_dict[leve]
                    
                    st.subheader(f"Abrangência e Prazos: {nome_exibicao}")
                    
                    df_abrangencia_leve = df_abrangencia[df_abrangencia['LMC Name'] == leve].copy()
                    
                    df_abrangencia_leve['Routing Code'] = mapa_routing.get(leve, "-")
                    df_abrangencia_leve['SLO Local (Arquivo)'] = df_abrangencia_leve['Prazo adicional']
                    
                    colunas_abrangencia_exibicao = ['LMC Name', 'Routing Code', 'Região de preço', 'Cidade', 'State', 'SLO Local (Arquivo)']
                    colunas_presentes_abrangencia = [col for col in colunas_abrangencia_exibicao if col in df_abrangencia_leve.columns]
                    
                    st.dataframe(df_abrangencia_leve[colunas_presentes_abrangencia], use_container_width=True, hide_index=True)
                    st.markdown("<br>", unsafe_allow_html=True)

else:
    st.info("Por favor, faça o upload de **todas as 4 bases** na barra lateral para prosseguir.")