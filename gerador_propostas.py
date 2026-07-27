import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import io
import unicodedata

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
def load_local_excel(filename):
    if not os.path.exists(filename):
        return None
    return pd.read_excel(filename)

def normalize_string(s):
    """Remove acentos, espaços extras e transforma em minúsculo para cruzamento seguro."""
    if pd.isna(s): return ""
    s = str(s).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def padronizar_colunas_volume(df):
    mapa_colunas = {
        "Package Charge Leve Last Mile Company Name": "Leve",
        "Distribution and Expedition Center Locations Routing Code": "Routing Code",
        "Package Charge Leve Region Label": "Region label",
        "Package Charge Leve Service Charge Type": "Service Charge Type",
        "Faixa Pesos": "Faixa pesos",
        "Package Charge Leve # Packages": "# Total Packages",
        "Package Destination City": "Cidade"
    }
    df = df.rename(columns=mapa_colunas)
    renames_fallback = {}
    for col in df.columns:
        col_str = str(col)
        if "Leve" not in df.columns and col_str.endswith("Leve"):
            renames_fallback[col_str] = "Leve"
        elif "Routing Code" not in df.columns and col_str.endswith("Routing Code"):
            renames_fallback[col_str] = "Routing Code"
        elif "Region label" not in df.columns and col_str.endswith("Region label"):
            renames_fallback[col_str] = "Region label"
        elif "# Total Packages" not in df.columns and col_str.endswith("Total Packages"):
            renames_fallback[col_str] = "# Total Packages"
        elif "Faixa pesos" not in df.columns and col_str.lower().endswith("faixa pesos"):
            renames_fallback[col_str] = "Faixa pesos"
        elif "Cidade" not in df.columns and (col_str.endswith("Destination City") or col_str.lower().endswith("cidade")):
            renames_fallback[col_str] = "Cidade"
    return df.rename(columns=renames_fallback)

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

@st.cache_data
def processar_price_var(df_price):
    df_price.columns = ['Faixa de peso (g/m³)', 'Multiplicador']
    return df_price

def formatar_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return valor

def extrair_estado(nome_leve):
    match = re.search(r'-\s*([A-Z]{2})', nome_leve)
    if match:
        return match.group(1)
    return None

# --- FLUXO PRINCIPAL ---
df_price_var_raw = load_local_excel("Price variation.xlsx")

if file_frete and file_abrangencia and file_slos and file_volume:
    if df_price_var_raw is None:
        st.error("Erro: O arquivo 'Price variation.xlsx' não foi encontrado. Por favor, certifique-se de que ele foi subido para o repositório do GitHub.")
    else:
        with st.spinner("Carregando e processando bases..."):
            df_frete = load_data(file_frete)
            df_abrangencia = load_data(file_abrangencia)
            df_slos = load_data(file_slos)
            
            df_volume = load_data(file_volume)
            df_volume = padronizar_colunas_volume(df_volume)
            
            df_volume['Cidade_Normalizada'] = df_volume['Cidade'].apply(normalize_string)
            df_abrangencia['Cidade_Normalizada'] = df_abrangencia['Cidade'].apply(normalize_string)
            
            df_price_var_clean = processar_price_var(df_price_var_raw)
            
        if 'Leve' not in df_volume.columns or 'Routing Code' not in df_volume.columns:
            st.error("🚨 **Colunas básicas ausentes no arquivo de Volume!**")
            st.stop()
            
        if 'Cidade' not in df_volume.columns or 'Faixa pesos' not in df_volume.columns:
            st.error("🚨 **Atenção: Base de Volume Desatualizada!**")
            st.markdown("Faça o upload da base mais recente com as colunas de Cidade e Faixa Pesos.")
            st.stop()
            
        with st.spinner("Finalizando processamento..."):
            df_frete_clean = processar_frete(df_frete)
            df_slos_clean = processar_slos(df_slos)
            df_nomes_leves = processar_nomes_leves(df_volume)
            
        st.sidebar.success("Todas as bases carregadas!")
    
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
    
            with st.expander("4. Dados Atuais dos Leves Selecionados", expanded=False):
                df_abrangencia.rename(columns={'Região de preço 2023': 'Região de preço'}, inplace=True)
                
                tab1, tab2 = st.tabs(["Tabela Frete Peso Atual", "Abrangência e Prazos"])
                
                with tab1:
                    for idx, leve in enumerate(leves_selecionados):
                        st.subheader(f"Tabela Frete Peso: {leves_selecionados_formatados[idx]}")
                        df_frete_leve = df_frete_clean[df_frete_clean['LMC name'] == leve].copy()
                        df_frete_leve['Routing Code'] = mapa_routing.get(leve, "-")
                        df_frete_leve.rename(columns={
                            'label': 'Região de preço',
                            'on time amount': 'Valor do pacote dentro do prazo',
                            'out of time amount': 'Valor do pacote fora do prazo'
                        }, inplace=True)
                        df_frete_leve['Valor do pacote dentro do prazo'] = df_frete_leve['Valor do pacote dentro do prazo'].apply(formatar_moeda)
                        df_frete_leve['Valor do pacote fora do prazo'] = df_frete_leve['Valor do pacote fora do prazo'].apply(formatar_moeda)
                        col_exib = ['LMC name', 'Routing Code', 'Região de preço', 'Faixa de peso (g/m³)', 'Valor do pacote dentro do prazo', 'Valor do pacote fora do prazo', 'table name']
                        st.dataframe(df_frete_leve[[c for c in col_exib if c in df_frete_leve.columns]], use_container_width=True, hide_index=True)
                    
                with tab2:
                    for idx, leve in enumerate(leves_selecionados):
                        st.subheader(f"Abrangência e Prazos: {leves_selecionados_formatados[idx]}")
                        df_abrangencia_leve = df_abrangencia[df_abrangencia['LMC Name'] == leve].copy()
                        df_abrangencia_leve['Routing Code'] = mapa_routing.get(leve, "-")
                        df_abrangencia_leve['SLO Local (Arquivo)'] = df_abrangencia_leve['Prazo adicional']
                        col_ab = ['LMC Name', 'Routing Code', 'Região de preço', 'Cidade', 'State', 'SLO Local (Arquivo)']
                        st.dataframe(df_abrangencia_leve[[c for c in col_ab if c in df_abrangencia_leve.columns]], use_container_width=True, hide_index=True)
                        
                # Adicionado botão para baixar propostas atuais
                st.divider()
                st.markdown("### 📥 Download das Propostas Atuais")
                st.markdown("Baixe um Excel com a abrangência e tabela frete peso atual de cada um dos Leves selecionados.")
                
                cols_download = st.columns(len(leves_selecionados))
                for idx, leve in enumerate(leves_selecionados):
                    with cols_download[idx]:
                        # Prepara dados para download
                        df_abr_dl = df_abrangencia[df_abrangencia['LMC Name'] == leve].copy()
                        df_abr_dl['Routing Code'] = mapa_routing.get(leve, "-")
                        df_abr_dl['SLO Local'] = df_abr_dl['Prazo adicional']
                        colunas_abr_dl = ['Cidade', 'State', 'Região de preço', 'SLO Local']
                        if not df_abr_dl.empty:
                            df_abr_dl = df_abr_dl[colunas_abr_dl]
                        
                        df_frete_dl = df_frete_clean[df_frete_clean['LMC name'] == leve].copy()
                        df_frete_dl.rename(columns={
                            'label': 'Região de Preço',
                            'on time amount': 'Valor dentro do prazo',
                            'out of time amount': 'Valor fora do prazo'
                        }, inplace=True)
                        colunas_frete_dl = ['Região de Preço', 'Faixa de peso (g/m³)', 'Valor dentro do prazo', 'Valor fora do prazo']
                        if not df_frete_dl.empty:
                            df_frete_dl = df_frete_dl[colunas_frete_dl]
                            
                        output_atual = io.BytesIO()
                        with pd.ExcelWriter(output_atual, engine='openpyxl') as writer:
                            df_abr_dl.to_excel(writer, sheet_name='Abrangência e Prazos', index=False)
                            df_frete_dl.to_excel(writer, sheet_name='Tabela Frete Peso', index=False)
                        
                        st.download_button(
                            label=f"Baixar Proposta Atual: {mapa_routing.get(leve, '')}",
                            data=output_atual.getvalue(),
                            file_name=f"Proposta_Atual_{leve.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_atual_{idx}"
                        )

    
            with st.expander("5. Definição do Leve/Lead de Destino", expanded=True):
                tipo_destino = st.radio("O destino da movimentação será para:", ["Um Leve Existente (já selecionado)", "Um Novo Lead"])
                
                nome_destino_final = ""
                cidade_base_destino = ""
                
                if tipo_destino == "Um Leve Existente (já selecionado)":
                    nome_destino_display = st.selectbox("Selecione o Leve de Destino:", leves_selecionados_formatados)
                    nome_destino_final = mapa_nomes.get(nome_destino_display)
                    cidade_base_destino = cidades_base_dict.get(nome_destino_final)
                    st.info(f"**Cidade Base do Destino:** {cidade_base_destino}")
                else:
                    col_n1, col_n2, col_n3 = st.columns(3)
                    with col_n1:
                        nome_destino_final = st.text_input("Nome do Novo Lead:", placeholder="Ex: Lead - SP Sorocaba...")
                    with col_n2:
                        estados_disponiveis = sorted(df_slos_clean['State'].dropna().unique().tolist())
                        estado_lead = st.selectbox("Estado do Novo Lead:", estados_disponiveis)
                    with col_n3:
                        cidades_estado_lead = df_slos_clean[df_slos_clean['State'] == estado_lead]['Cidade'].tolist()
                        cidade_base_destino = st.selectbox("Cidade Base do Novo Lead:", sorted(cidades_estado_lead))
    
            if nome_destino_final and cidade_base_destino:
                with st.expander("6. Manipulação de Abrangência (Tabela Interativa)", expanded=True):
                    
                    config_key = f"{','.join(leves_selecionados)}_{nome_destino_final}"
                    if "mov_config_key" not in st.session_state or st.session_state.mov_config_key != config_key:
                        df_abrangencia_alvo = df_abrangencia[df_abrangencia['LMC Name'].isin(leves_selecionados)].copy()
                        df_mov = df_abrangencia_alvo[['LMC Name', 'Região de preço', 'Cidade', 'State']].copy()
                        df_mov['Destino'] = "Manter no Leve Atual"
                        st.session_state.df_movimentacao = df_mov
                        st.session_state.mov_config_key = config_key

                    opcoes_destino = ["Manter no Leve Atual", nome_destino_final]

                    st.markdown("⚡ **Ações em Massa (Mover vários municípios de uma vez)**")
                    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns([2.5, 2.5, 2.5, 1.5, 1.5])
                    
                    with col_m1:
                        bulk_lmc = st.selectbox("1. Filtrar por Origem:", ["Selecione..."] + leves_selecionados, key="bulk_lmc")
                    with col_m2:
                        opcoes_reg = ["Todas as Regiões"]
                        if bulk_lmc != "Selecione...":
                            opcoes_reg += sorted(list(st.session_state.df_movimentacao[st.session_state.df_movimentacao['LMC Name'] == bulk_lmc]['Região de preço'].unique()))
                        bulk_reg = st.selectbox("2. Filtrar por Região:", opcoes_reg, key="bulk_reg")
                    with col_m3:
                        bulk_dest = st.selectbox("3. Escolher Destino:", opcoes_destino, key="bulk_dest")
                        
                    with col_m4:
                        st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                        def aplicar_em_massa():
                            l = st.session_state.bulk_lmc
                            r = st.session_state.bulk_reg
                            d = st.session_state.bulk_dest
                            if l != "Selecione...":
                                mask = st.session_state.df_movimentacao['LMC Name'] == l
                                if r != "Todas as Regiões":
                                    mask &= st.session_state.df_movimentacao['Região de preço'] == r
                                st.session_state.df_movimentacao.loc[mask, 'Destino'] = d
                        st.button("Aplicar Ação", on_click=aplicar_em_massa, type="secondary", use_container_width=True)
                        
                    with col_m5:
                        st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                        def resetar_massa():
                            st.session_state.df_movimentacao['Destino'] = "Manter no Leve Atual"
                        st.button("🔄 Resetar Tudo", on_click=resetar_massa, use_container_width=True)

                    st.divider()
                    st.markdown("🖱️ **Seleção Manual (Município a Município)**")
                    st.markdown("Selecione na coluna **'Destino'** para onde cada município deve ir de forma manual.")
                    
                    df_editado = st.data_editor(
                        st.session_state.df_movimentacao,
                        column_config={
                            "Destino": st.column_config.SelectboxColumn(
                                "Destino (Clique para alterar)",
                                help="Selecione o destino deste município",
                                options=opcoes_destino,
                                required=True,
                            )
                        },
                        disabled=["LMC Name", "Região de preço", "Cidade", "State"],
                        hide_index=True,
                        use_container_width=True,
                        height=400
                    )
                    
                    st.session_state.df_movimentacao = df_editado
                    df_movidos = df_editado[df_editado['Destino'] == nome_destino_final].copy()

                with st.expander("7. Estratégia de Precificação da Proposta", expanded=True):
                    if df_movidos.empty:
                        st.warning("Nenhum município foi movimentado para o destino ainda.")
                    else:
                        st.markdown("Como você deseja gerar a nova Tabela Frete Peso para as cidades movidas?")
                        estrategia_preco = st.radio(
                            "Escolha a estratégia:", 
                            ["Gerar Tabela Equivalente (Focada em manter Impacto Neutro)", 
                             "Utilizar uma Tabela Existente (Manter tabela do destino atual)"]
                        )
                        
                        tabela_base_selecionada = None
                        if estrategia_preco == "Utilizar uma Tabela Existente (Manter tabela do destino atual)":
                            if tipo_destino == "Um Leve Existente (já selecionado)":
                                st.info(f"O sistema utilizará a tabela atual do Leve **{nome_destino_final}** para precificar as cidades absorvidas.")
                                tabela_base_selecionada = nome_destino_final
                            else:
                                nome_base_display = st.selectbox("Qual tabela atual devemos usar como base para o Novo Lead?", leves_selecionados_formatados)
                                tabela_base_selecionada = mapa_nomes.get(nome_base_display)
                
                # --- PROCESSAMENTO DOS RESULTADOS ---
                if not df_movidos.empty:
                    if tipo_destino == "Um Leve Existente (já selecionado)":
                        df_abrangencia_existente = df_abrangencia[df_abrangencia['LMC Name'] == nome_destino_final].copy()
                        df_abrangencia_existente['Observação'] = "Abrangência Atual"
                    else:
                        df_abrangencia_existente = pd.DataFrame(columns=['LMC Name', 'Região de preço', 'Cidade', 'State', 'Observação'])

                    df_movidos_fmt = df_movidos.copy()
                    df_movidos_fmt['Observação'] = "Migrado de: " + df_movidos_fmt['LMC Name']

                    df_escopo_final = pd.concat([
                        df_abrangencia_existente[['Cidade', 'State', 'Região de preço', 'Observação']],
                        df_movidos_fmt[['Cidade', 'State', 'Região de preço', 'Observação']]
                    ], ignore_index=True).drop_duplicates(subset=['Cidade'])
                    
                    slo_base_dest = df_slos_clean[df_slos_clean['Cidade'] == cidade_base_destino]['SLO'].values
                    slo_base_dest_val = slo_base_dest[0] if len(slo_base_dest) > 0 else 0
                    
                    df_escopo_final = df_escopo_final.merge(df_slos_clean[['Cidade', 'SLO']], on='Cidade', how='left')
                    df_escopo_final['Novo SLO Local'] = df_escopo_final['SLO'] - slo_base_dest_val
                    df_escopo_final['Novo SLO Local'] = df_escopo_final['Novo SLO Local'].apply(lambda x: x if pd.notnull(x) and x > 0 else 0).astype(int)
                    
                    colunas_finais_abrangencia = ['Cidade', 'State', 'Região de preço', 'Novo SLO Local', 'Observação']
                    
                    regioes_finais_destino = df_escopo_final['Região de preço'].unique()
                    regioes_movimentadas = df_movidos['Região de preço'].unique()

                    for regiao in regioes_finais_destino:
                        if f"ajuste_{regiao}" not in st.session_state:
                            st.session_state[f"ajuste_{regiao}"] = 0.0

                    def aplicar_ajuste_global():
                        val = st.session_state.ajuste_global_input
                        for r in regioes_finais_destino:
                            st.session_state[f"ajuste_{r}"] = val

                    def zerar_ajuste_regiao(r):
                        st.session_state[f"ajuste_{r}"] = 0.0

                    with st.expander("8. Ajustes Comerciais (Opcional)", expanded=True):
                        st.markdown("Insira um percentual de reajuste (positivo ou negativo) para as tabelas do Leve de destino. Deixe 0 para manter o valor calculado.")
                        
                        st.markdown("**Ajuste em Massa:**")
                        cg1, cg2, cg3 = st.columns([1.5, 2, 4])
                        with cg1:
                            st.number_input("Todas as regiões (%)", step=0.5, format="%.2f", key="ajuste_global_input")
                        with cg2:
                            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                            st.button("Aplicar a todas", on_click=aplicar_ajuste_global)
                            
                        st.divider()
                        
                        st.markdown("**Ajustes Individuais:**")
                        ajustes_dict = {}
                        cols_ajuste = st.columns(3)
                        for i, regiao in enumerate(sorted(regioes_finais_destino)):
                            with cols_ajuste[i % 3]:
                                c_input, c_btn = st.columns([2, 1])
                                with c_input:
                                    ajustes_dict[regiao] = st.number_input(f"Ajuste {regiao} (%)", step=0.5, format="%.2f", key=f"ajuste_{regiao}")
                                with c_btn:
                                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                                    st.button("Zerar", key=f"btn_zerar_{regiao}", on_click=zerar_ajuste_regiao, args=(regiao,))

                    with st.expander("9. Resultados e Impacto Financeiro (Proposta)", expanded=True):
                        st.success(f"Proposta gerada para **{nome_destino_final}**!")
                        
                        lista_novas_tabelas = []
                        registros_auditoria = []
                        texto_explicativo = []
                        texto_explicativo.append(f"=== MEMORIAL DE CÁLCULO - PROPOSTA FRETE PESO ===\n")
                        texto_explicativo.append(f"Destino: {nome_destino_final}")
                        texto_explicativo.append(f"Estratégia Escolhida: {estrategia_preco}\n")
                        
                        global_vol_destino_original = 0
                        global_custo_destino_original = 0
                        global_vol_migrado = 0
                        global_custo_migrado_original = 0
                        global_custo_novo_total = 0
                        
                        detalhes_regioes = {}
                        
                        for regiao in regioes_finais_destino:
                            texto_explicativo.append(f"--------------------------------------------------")
                            texto_explicativo.append(f"📍 REGIÃO: {regiao}")
                            texto_explicativo.append(f"--------------------------------------------------")
                            
                            cidades_mov_regiao = df_movidos[df_movidos['Região de preço'] == regiao]
                            
                            if regiao in regioes_movimentadas:
                                if estrategia_preco == "Gerar Tabela Equivalente (Focada em manter Impacto Neutro)":
                                    custo_alvo_on = 0
                                    custo_alvo_out = 0
                                    soma_vol_mult = 0
                                    
                                    mult_dict = dict(zip(df_price_var_clean['Faixa de peso (g/m³)'], df_price_var_clean['Multiplicador']))
                                    
                                    texto_explicativo.append(">> PASSO 1: Levantamento do Custo Antigo Global e Pesos (Análise de todas as 23 faixas)\n")
                                    texto_explicativo.append("Cidades Migradas (Origem):")
                                    
                                    for _, row in cidades_mov_regiao.iterrows():
                                        leve_origem = row['LMC Name']
                                        cidade_movida = row['Cidade']
                                        
                                        vol_data = df_volume[(df_volume['Leve'] == leve_origem) & 
                                                             (df_volume['Cidade_Normalizada'] == normalize_string(cidade_movida))]
                                        
                                        vol_cidade = 0
                                        custo_cidade = 0
                                        
                                        for _, v_row in vol_data.iterrows():
                                            fx = v_row['Faixa pesos']
                                            qtd = v_row['# Total Packages']
                                            
                                            tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == leve_origem) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == fx)]
                                            preco_on = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                            preco_out = tb_antiga['out of time amount'].values[0] if not tb_antiga.empty else 0
                                            
                                            mult = mult_dict.get(fx, 1.0)
                                            
                                            custo_alvo_on += qtd * preco_on
                                            custo_alvo_out += qtd * preco_out
                                            soma_vol_mult += qtd * mult
                                            
                                            vol_cidade += qtd
                                            custo_cidade += qtd * preco_on
                                            
                                        if vol_cidade > 0:
                                            texto_explicativo.append(f" - {str(cidade_movida).title()} ({leve_origem}): {int(vol_cidade)} pct | Custo Atual: {formatar_moeda(custo_cidade)}")
                                    
                                    if tipo_destino == "Um Leve Existente (já selecionado)":
                                        texto_explicativo.append("\nCidades Atuais do Destino:")
                                        cidades_exist = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == regiao]
                                        
                                        for _, row in cidades_exist.iterrows():
                                            cid_ext = row['Cidade']
                                            vol_dest_data = df_volume[(df_volume['Leve'] == nome_destino_final) & 
                                                                      (df_volume['Cidade_Normalizada'] == normalize_string(cid_ext))]
                                            
                                            vol_cidade = 0
                                            custo_cidade = 0
                                            for _, v_row in vol_dest_data.iterrows():
                                                fx = v_row['Faixa pesos']
                                                qtd = v_row['# Total Packages']
                                                
                                                tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == fx)]
                                                preco_on = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                                preco_out = tb_antiga['out of time amount'].values[0] if not tb_antiga.empty else 0
                                                
                                                mult = mult_dict.get(fx, 1.0)
                                                
                                                custo_alvo_on += qtd * preco_on
                                                custo_alvo_out += qtd * preco_out
                                                soma_vol_mult += qtd * mult
                                                
                                                vol_cidade += qtd
                                                custo_cidade += qtd * preco_on
                                                
                                            if vol_cidade > 0:
                                                texto_explicativo.append(f" - {str(cid_ext).title()}: {int(vol_cidade)} pct | Custo Atual: {formatar_moeda(custo_cidade)}")

                                    base_on = custo_alvo_on / soma_vol_mult if soma_vol_mult > 0 else 0
                                    base_out = custo_alvo_out / soma_vol_mult if soma_vol_mult > 0 else 0
                                    
                                    texto_explicativo.append(f"\n>> PASSO 2: Cálculo da Base Matemática Ponderada Global")
                                    texto_explicativo.append(f"Custo Alvo Total (Todas as 23 faixas) = {formatar_moeda(custo_alvo_on)}")
                                    texto_explicativo.append(f"Fator de Volume Ponderado (Soma Vol * Mult) = {soma_vol_mult:.2f}")
                                    texto_explicativo.append(f"Base de Cálculo (Índice 1.00) = Custo Alvo / Fator = {formatar_moeda(base_on)}")
                                    
                                else:
                                    frete_base = df_frete_clean[(df_frete_clean['LMC name'] == tabela_base_selecionada) & (df_frete_clean['label'] == regiao)]
                                    faixa1_base = frete_base[frete_base['Faixa de peso (g/m³)'].str.contains('01 0 to 300', na=False, case=False)]
                                    nova_faixa1_on = faixa1_base['on time amount'].values[0] if not faixa1_base.empty else 0
                                    nova_faixa1_out = faixa1_base['out of time amount'].values[0] if not faixa1_base.empty else 0
                                    
                                    texto_explicativo.append(f">> Tabela Existente Selecionada: {tabela_base_selecionada}")
                                    texto_explicativo.append(f"Valor copiado da Faixa 1: {formatar_moeda(nova_faixa1_on)}")
                                    
                                    base_on = nova_faixa1_on / 0.83
                                    base_out = nova_faixa1_out / 0.83
                                
                                texto_explicativo.append(f"\n>> PASSO 3: Geração das 23 Faixas")
                                
                                df_regiao_tabela = df_price_var_clean.copy()
                                df_regiao_tabela['Região de Preço'] = regiao
                                df_regiao_tabela['Valor dentro do prazo'] = df_regiao_tabela['Multiplicador'] * base_on
                                df_regiao_tabela['Valor fora do prazo'] = df_regiao_tabela['Multiplicador'] * base_out
                                
                                ajuste_perc = ajustes_dict.get(regiao, 0.0)
                                if ajuste_perc != 0.0:
                                    fator = 1 + (ajuste_perc / 100)
                                    df_regiao_tabela['Valor dentro do prazo'] = df_regiao_tabela['Valor dentro do prazo'] * fator
                                    df_regiao_tabela['Valor fora do prazo'] = df_regiao_tabela['Valor fora do prazo'] * fator
                                    texto_explicativo.append(f"\n>> PASSO 4: Ajuste Comercial")
                                    texto_explicativo.append(f"Aplicado reajuste de {ajuste_perc:+.2f}% negociado para a região {regiao}.\n")
                                
                                for _, prow in df_regiao_tabela.iterrows():
                                    fx = prow['Faixa de peso (g/m³)']
                                    mult = prow['Multiplicador']
                                    val = prow['Valor dentro do prazo']
                                    texto_explicativo.append(f"{fx}: Mult {mult:.2f} x {formatar_moeda(base_on)} = {formatar_moeda(val)}")
                                
                                reg_vol_antigo_loggi = 0
                                reg_custo_antigo_loggi = 0
                                reg_custo_novo_loggi = 0
                                
                                for _, row in cidades_mov_regiao.iterrows():
                                    leve_orig = row['LMC Name']
                                    cid_movida = row['Cidade']
                                    
                                    vols_cidade = df_volume[(df_volume['Leve'] == leve_orig) & (df_volume['Cidade_Normalizada'] == normalize_string(cid_movida))]
                                    for _, v_row in vols_cidade.iterrows():
                                        faixa_peso = v_row['Faixa pesos']
                                        qtd_pacotes = v_row['# Total Packages']
                                        
                                        tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == leve_orig) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == faixa_peso)]
                                        preco_antigo = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                        
                                        tb_nova = df_regiao_tabela[df_regiao_tabela['Faixa de peso (g/m³)'] == faixa_peso]
                                        preco_novo = tb_nova['Valor dentro do prazo'].values[0] if not tb_nova.empty else 0
                                        
                                        global_vol_migrado += qtd_pacotes
                                        global_custo_migrado_original += (qtd_pacotes * preco_antigo)
                                        global_custo_novo_total += (qtd_pacotes * preco_novo)
                                        
                                        reg_vol_antigo_loggi += qtd_pacotes
                                        reg_custo_antigo_loggi += (qtd_pacotes * preco_antigo)
                                        reg_custo_novo_loggi += (qtd_pacotes * preco_novo)
                                        
                                        registros_auditoria.append({
                                            'Tipo': 'Cidades Migradas (Origem)',
                                            'LMC Atual / Origem': leve_orig,
                                            'Região de Preço': regiao,
                                            'Cidade': str(cid_movida).title(),
                                            'Faixa de Peso': faixa_peso,
                                            'Pacotes (30 dias)': qtd_pacotes,
                                            'Tarifa Antiga': preco_antigo,
                                            'Ajuste Comercial (%)': ajuste_perc / 100,
                                            'Tarifa Nova Projetada': preco_novo,
                                            'Custo Antigo Total': qtd_pacotes * preco_antigo,
                                            'Novo Custo Total': qtd_pacotes * preco_novo,
                                            'Diferença (R$)': (qtd_pacotes * preco_novo) - (qtd_pacotes * preco_antigo)
                                        })
                                        
                                if tipo_destino == "Um Leve Existente (já selecionado)":
                                    cidades_exist = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == regiao]
                                    for _, row in cidades_exist.iterrows():
                                        cid_ext = row['Cidade']
                                        vols_cidade = df_volume[(df_volume['Leve'] == nome_destino_final) & (df_volume['Cidade_Normalizada'] == normalize_string(cid_ext))]
                                        
                                        for _, v_row in vols_cidade.iterrows():
                                            faixa_peso = v_row['Faixa pesos']
                                            qtd_pacotes = v_row['# Total Packages']
                                            
                                            tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == faixa_peso)]
                                            preco_antigo = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                            
                                            tb_nova = df_regiao_tabela[df_regiao_tabela['Faixa de peso (g/m³)'] == faixa_peso]
                                            preco_novo = tb_nova['Valor dentro do prazo'].values[0] if not tb_nova.empty else 0
                                            
                                            global_vol_destino_original += qtd_pacotes
                                            global_custo_destino_original += (qtd_pacotes * preco_antigo)
                                            global_custo_novo_total += (qtd_pacotes * preco_novo)
                                            
                                            reg_vol_antigo_loggi += qtd_pacotes
                                            reg_custo_antigo_loggi += (qtd_pacotes * preco_antigo)
                                            reg_custo_novo_loggi += (qtd_pacotes * preco_novo)
                                            
                                            registros_auditoria.append({
                                                'Tipo': 'Cidades Atuais do Destino',
                                                'LMC Atual / Origem': nome_destino_final,
                                                'Região de Preço': regiao,
                                                'Cidade': str(cid_ext).title(),
                                                'Faixa de Peso': faixa_peso,
                                                'Pacotes (30 dias)': qtd_pacotes,
                                                'Tarifa Antiga': preco_antigo,
                                                'Ajuste Comercial (%)': ajuste_perc / 100,
                                                'Tarifa Nova Projetada': preco_novo,
                                                'Custo Antigo Total': qtd_pacotes * preco_antigo,
                                                'Novo Custo Total': qtd_pacotes * preco_novo,
                                                'Diferença (R$)': (qtd_pacotes * preco_novo) - (qtd_pacotes * preco_antigo)
                                            })
                                
                                detalhes_regioes[regiao] = {
                                    'vol': reg_vol_antigo_loggi,
                                    'custo_antigo': reg_custo_antigo_loggi,
                                    'custo_novo': reg_custo_novo_loggi
                                }
                                texto_explicativo.append("\n")

                            else:
                                texto_explicativo.append("Nenhuma movimentação nesta região. Tabela atual do Leve mantida.\n")
                                df_regiao_tabela = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao)].copy()
                                df_regiao_tabela.rename(columns={'label': 'Região de Preço', 'on time amount': 'Valor dentro do prazo', 'out of time amount': 'Valor fora do prazo'}, inplace=True)
                                
                                reg_vol_antigo_loggi = 0
                                reg_custo_antigo_loggi = 0
                                reg_custo_novo_loggi = 0
                                
                                ajuste_perc = ajustes_dict.get(regiao, 0.0)
                                if ajuste_perc != 0.0:
                                    fator = 1 + (ajuste_perc / 100)
                                    df_regiao_tabela['Valor dentro do prazo'] = df_regiao_tabela['Valor dentro do prazo'] * fator
                                    df_regiao_tabela['Valor fora do prazo'] = df_regiao_tabela['Valor fora do prazo'] * fator
                                    texto_explicativo.append(f">> Ajuste Comercial: Aplicado reajuste de {ajuste_perc:+.2f}% negociado para a região {regiao}.\n")
                                    
                                if tipo_destino == "Um Leve Existente (já selecionado)":
                                    cidades_exist = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == regiao]
                                    for _, row in cidades_exist.iterrows():
                                        cid_ext = row['Cidade']
                                        vols_cidade = df_volume[(df_volume['Leve'] == nome_destino_final) & (df_volume['Cidade_Normalizada'] == normalize_string(cid_ext))]
                                        
                                        for _, v_row in vols_cidade.iterrows():
                                            faixa_peso = v_row['Faixa pesos']
                                            qtd_pacotes = v_row['# Total Packages']
                                            
                                            tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == faixa_peso)]
                                            preco_antigo = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                            
                                            tb_nova = df_regiao_tabela[df_regiao_tabela['Faixa de peso (g/m³)'] == faixa_peso]
                                            preco_novo = tb_nova['Valor dentro do prazo'].values[0] if not tb_nova.empty else 0
                                            
                                            global_vol_destino_original += qtd_pacotes
                                            global_custo_destino_original += (qtd_pacotes * preco_antigo)
                                            global_custo_novo_total += (qtd_pacotes * preco_novo)
                                            
                                            reg_vol_antigo_loggi += qtd_pacotes
                                            reg_custo_antigo_loggi += (qtd_pacotes * preco_antigo)
                                            reg_custo_novo_loggi += (qtd_pacotes * preco_novo)
                                            
                                            registros_auditoria.append({
                                                'Tipo': 'Cidades Atuais do Destino',
                                                'LMC Atual / Origem': nome_destino_final,
                                                'Região de Preço': regiao,
                                                'Cidade': str(cid_ext).title(),
                                                'Faixa de Peso': faixa_peso,
                                                'Pacotes (30 dias)': qtd_pacotes,
                                                'Tarifa Antiga': preco_antigo,
                                                'Ajuste Comercial (%)': ajuste_perc / 100,
                                                'Tarifa Nova Projetada': preco_novo,
                                                'Custo Antigo Total': qtd_pacotes * preco_antigo,
                                                'Novo Custo Total': qtd_pacotes * preco_novo,
                                                'Diferença (R$)': (qtd_pacotes * preco_novo) - (qtd_pacotes * preco_antigo)
                                            })
                                            
                                if ajuste_perc != 0.0:
                                    detalhes_regioes[regiao] = {
                                        'vol': reg_vol_antigo_loggi,
                                        'custo_antigo': reg_custo_antigo_loggi,
                                        'custo_novo': reg_custo_novo_loggi
                                    }
                            
                            df_regiao_tabela = df_regiao_tabela[['Região de Preço', 'Faixa de peso (g/m³)', 'Valor dentro do prazo', 'Valor fora do prazo']]
                            lista_novas_tabelas.append(df_regiao_tabela)
                        
                        df_tabela_final = pd.concat(lista_novas_tabelas, ignore_index=True)
                        texto_completo_memorial = "\n".join(texto_explicativo)
                        df_auditoria_excel = pd.DataFrame(registros_auditoria)

                        # --- ALERTAS DE MISROUTES / PACOTES IGNORADOS ---
                        ignorado_info = []
                        total_ignorados_geral = 0
                        leves_para_checar = set(leves_selecionados)
                        if tipo_destino == "Um Leve Existente (já selecionado)":
                            leves_para_checar.add(nome_destino_final)

                        for leve in leves_para_checar:
                            vol_leve = df_volume[df_volume['Leve'] == leve]
                            cidades_oficiais = df_abrangencia[df_abrangencia['LMC Name'] == leve]['Cidade_Normalizada'].unique()
                            vol_fora = vol_leve[~vol_leve['Cidade_Normalizada'].isin(cidades_oficiais)]
                            
                            if not vol_fora.empty:
                                qtd_fora = vol_fora['# Total Packages'].sum()
                                total_ignorados_geral += qtd_fora
                                top_erros = vol_fora.groupby('Cidade')['# Total Packages'].sum().sort_values(ascending=False).head(3)
                                exemplos_str = ", ".join([f"{cid} ({qtd} pct)" for cid, qtd in top_erros.items()])
                                ignorado_info.append(f"**{leve}:** {int(qtd_fora)} pacotes descartados. *(Ex: {exemplos_str})*")

                        if total_ignorados_geral > 0:
                            with st.expander(f"⚠️ {int(total_ignorados_geral)} pacotes não contabilizados (Misroutes e Erros de Digitação)", expanded=False):
                                st.markdown("Os pacotes abaixo constam na base de volume, mas as cidades informadas **não existem** no contrato oficial (Abrangência) do Leve. Eles foram automaticamente isolados e desconsiderados para não sujar o cálculo financeiro das tabelas.")
                                for info in ignorado_info:
                                    st.markdown(f"- {info}")
                        
                        # --- EXIBIÇÃO DE FATURAMENTO DO PARCEIRO ---
                        st.subheader("🤝 Visão do Parceiro (Faturamento do Leve)")
                        st.markdown("Esta visão demonstra o tamanho da operação do parceiro **antes** (sem as cidades novas) e **depois** de absorver o novo escopo.")
                        
                        vol_parceiro_novo = global_vol_destino_original + global_vol_migrado
                        tk_parceiro_antigo = global_custo_destino_original / global_vol_destino_original if global_vol_destino_original > 0 else 0
                        tk_parceiro_novo = global_custo_novo_total / vol_parceiro_novo if vol_parceiro_novo > 0 else 0
                        crescimento_parceiro = global_custo_novo_total - global_custo_destino_original
                        perc_crescimento = (crescimento_parceiro / global_custo_destino_original) * 100 if global_custo_destino_original > 0 else 100
                        
                        cp1, cp2, cp3 = st.columns(3)
                        with cp1:
                            st.metric("Faturamento Atual (Sem Novas Cidades)", formatar_moeda(global_custo_destino_original))
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria: {int(global_vol_destino_original):,} pacotes</span>", unsafe_allow_html=True)
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio: {formatar_moeda(tk_parceiro_antigo)}</span>", unsafe_allow_html=True)
                        with cp2:
                            st.metric("Novo Faturamento Projetado", formatar_moeda(global_custo_novo_total))
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria: {int(vol_parceiro_novo):,} pacotes</span>", unsafe_allow_html=True)
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio: {formatar_moeda(tk_parceiro_novo)}</span>", unsafe_allow_html=True)
                        with cp3:
                            st.metric("Crescimento da Operação", formatar_moeda(crescimento_parceiro))
                            st.markdown(f"<span style='font-size: 0.9em; color: #09ab3b; font-weight: bold;'>▲ +{perc_crescimento:.2f}% de aumento no faturamento</span>", unsafe_allow_html=True)

                        st.divider()

                        # --- EXIBIÇÃO DE IMPACTO LOGGI ---
                        st.subheader("📉 Visão Loggi (Impacto Financeiro Real)")
                        st.markdown("Esta visão compara o custo da volumetria total afetada (Destino + Origem) na **tabela antiga** versus a **tabela nova projetada**.")
                        
                        custo_loggi_antigo = global_custo_destino_original + global_custo_migrado_original
                        vol_loggi_total = global_vol_destino_original + global_vol_migrado
                        tk_loggi_antigo = custo_loggi_antigo / vol_loggi_total if vol_loggi_total > 0 else 0
                        tk_loggi_novo = global_custo_novo_total / vol_loggi_total if vol_loggi_total > 0 else 0
                        
                        impacto_loggi = global_custo_novo_total - custo_loggi_antigo
                        perc_impacto_loggi = (impacto_loggi / custo_loggi_antigo) * 100 if custo_loggi_antigo > 0 else 0

                        cl1, cl2, cl3 = st.columns(3)
                        with cl1:
                            st.metric("Custo Antigo Global (Destino + Cidades Migradas)", formatar_moeda(custo_loggi_antigo))
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria Total: {int(vol_loggi_total):,} pacotes</span>", unsafe_allow_html=True)
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio Antigo: {formatar_moeda(tk_loggi_antigo)}</span>", unsafe_allow_html=True)
                        with cl2:
                            st.metric("Novo Custo Global Projetado", formatar_moeda(global_custo_novo_total))
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria Total: {int(vol_loggi_total):,} pacotes</span>", unsafe_allow_html=True)
                            st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio Novo: {formatar_moeda(tk_loggi_novo)}</span>", unsafe_allow_html=True)
                        with cl3:
                            if impacto_loggi > 0:
                                st.metric("Impacto Financeiro Loggi", formatar_moeda(impacto_loggi), f"+{formatar_moeda(impacto_loggi)} (Aumento de custo)", delta_color="inverse")
                                st.markdown(f"<span style='font-size: 0.9em; color: #ff4b4b; font-weight: bold;'>▲ +{perc_impacto_loggi:.2f}% de impacto no budget</span>", unsafe_allow_html=True)
                            elif impacto_loggi < 0:
                                st.metric("Impacto Financeiro Loggi", formatar_moeda(impacto_loggi), f"{formatar_moeda(impacto_loggi)} (Economia)", delta_color="normal")
                                st.markdown(f"<span style='font-size: 0.9em; color: #09ab3b; font-weight: bold;'>▼ {perc_impacto_loggi:.2f}% de economia no budget</span>", unsafe_allow_html=True)
                            else:
                                st.metric("Impacto Financeiro Loggi", "R$ 0,00", "Neutro")
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>0.00%</span>", unsafe_allow_html=True)
                        
                        st.markdown("<br>", unsafe_allow_html=True)

                        # --- DETALHAMENTO POR REGIÃO ---
                        if detalhes_regioes:
                            st.markdown("#### 🔍 Detalhamento das Regiões Afetadas")
                            st.markdown("Veja abaixo o impacto financeiro isolado para cada Região de Preço que sofreu movimentação de cidades ou reajuste comercial.")
                            
                            for reg, dados in detalhes_regioes.items():
                                st.markdown(f"##### 📍 {reg}")
                                cd1, cd2, cd3 = st.columns(3)
                                tk_r_ant = dados['custo_antigo'] / dados['vol'] if dados['vol'] > 0 else 0
                                tk_r_nov = dados['custo_novo'] / dados['vol'] if dados['vol'] > 0 else 0
                                imp_r = dados['custo_novo'] - dados['custo_antigo']
                                perc_r = (imp_r / dados['custo_antigo']) * 100 if dados['custo_antigo'] > 0 else 0
                                
                                with cd1:
                                    st.markdown(f"**Custo Antigo:** {formatar_moeda(dados['custo_antigo'])}")
                                    st.markdown(f"<span style='font-size: 0.8em; color: gray;'>Vol: {int(dados['vol']):,} | Tk: {formatar_moeda(tk_r_ant)}</span>", unsafe_allow_html=True)
                                with cd2:
                                    st.markdown(f"**Novo Custo:** {formatar_moeda(dados['custo_novo'])}")
                                    st.markdown(f"<span style='font-size: 0.8em; color: gray;'>Vol: {int(dados['vol']):,} | Tk: {formatar_moeda(tk_r_nov)}</span>", unsafe_allow_html=True)
                                with cd3:
                                    color = "gray"
                                    if imp_r > 0: color = "#ff4b4b"
                                    elif imp_r < 0: color = "#09ab3b"
                                    st.markdown(f"**Variação:** <span style='color: {color}; font-weight: bold;'>{formatar_moeda(imp_r)} ({perc_r:+.2f}%)</span>", unsafe_allow_html=True)
                                st.markdown("<br>", unsafe_allow_html=True)

                        with st.expander("🤔 Detalhes do Cálculo & Proporção das Tabelas", expanded=False):
                            st.markdown("O aplicativo avalia o volume completo nas 23 faixas de peso simultaneamente e encontra matematicamente o multiplicador base perfeito que zera o impacto financeiro na geração de tabelas equivalentes.")
                            st.download_button(
                                label="📄 Baixar Memorial de Cálculo (.txt)",
                                data=texto_completo_memorial,
                                file_name=f"Memorial_Calculo_{nome_destino_final.replace(' ', '_')}.txt",
                                mime="text/plain"
                            )

                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        # --- EXIBIÇÃO DAS TABELAS COMPLETAS ---
                        st.subheader("1. Abrangência Completa Projetada")
                        st.dataframe(df_escopo_final[colunas_finais_abrangencia], hide_index=True, use_container_width=True)
                        
                        st.subheader("2. Tabelas Frete Peso (Todas as Regiões)")
                        df_exibicao_tabela = df_tabela_final.copy()
                        df_exibicao_tabela['Valor dentro do prazo'] = df_exibicao_tabela['Valor dentro do prazo'].apply(formatar_moeda)
                        df_exibicao_tabela['Valor fora do prazo'] = df_exibicao_tabela['Valor fora do prazo'].apply(formatar_moeda)
                        st.dataframe(df_exibicao_tabela, hide_index=True, use_container_width=True)

                        st.divider()
                        
                        col_dw1, col_down2 = st.columns(2)
                        with col_dw1:
                            st.markdown("### 📥 Download da Proposta Final")
                            st.markdown("Proposta a ser apresentada ao Leve / Lead")
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_escopo_final[colunas_finais_abrangencia].to_excel(writer, sheet_name='Abrangência e Prazos', index=False)
                                df_exibicao_tabela.to_excel(writer, sheet_name='Tabela Frete Peso', index=False)
                            
                            st.download_button(
                                label="Baixar Proposta em Excel",
                                data=output.getvalue(),
                                file_name=f"Proposta_Movimentacao_{nome_destino_final.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                            
                        with col_down2:
                            if not df_auditoria_excel.empty:
                                st.markdown("### 📊 Download do Detalhamento")
                                st.markdown("Planilha detalhando com os cálculos e impactos")
                                
                                df_resumo = pd.DataFrame({
                                    'Custo Antigo Total': [custo_loggi_antigo],
                                    'Custo Novo Total': [global_custo_novo_total],
                                    'Diferença (R$)': [impacto_loggi]
                                })
                                
                                output_det = io.BytesIO()
                                with pd.ExcelWriter(output_det, engine='openpyxl') as writer:
                                    df_resumo.to_excel(writer, sheet_name='Detalhamento Impacto', index=False, startrow=0)
                                    df_auditoria_excel.to_excel(writer, sheet_name='Detalhamento Impacto', index=False, startrow=4)
                                
                                st.download_button(
                                    label="Baixar Detalhes do Cálculo (.xlsx)",
                                    data=output_det.getvalue(),
                                    file_name=f"Detalhamento_Calculo_{nome_destino_final.replace(' ', '_')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="secondary"
                                )

else:
    st.info("Por favor, faça o upload de **todas as 4 bases** na barra lateral para prosseguir.")
