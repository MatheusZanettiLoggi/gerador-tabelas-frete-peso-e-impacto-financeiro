import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import io

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
    match = re.search(r'-\s*([A-Z]{2})\b', nome_leve)
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
            
            df_price_var_clean = processar_price_var(df_price_var_raw)
            
        if 'Leve' not in df_volume.columns or 'Routing Code' not in df_volume.columns or 'Cidade' not in df_volume.columns:
            st.error("🚨 **Colunas ausentes no arquivo de Volume!**")
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
                    st.markdown("Selecione na coluna **'Destino'** para onde cada município deve ir.")
                    
                    df_abrangencia_alvo = df_abrangencia[df_abrangencia['LMC Name'].isin(leves_selecionados)].copy()
                    df_movimentacao = df_abrangencia_alvo[['LMC Name', 'Região de preço', 'Cidade', 'State']].copy()
                    df_movimentacao['Destino'] = "Manter no Leve Atual"
                    
                    opcoes_destino = ["Manter no Leve Atual", nome_destino_final]
                    
                    df_editado = st.data_editor(
                        df_movimentacao,
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
                    with st.expander("8. Resultados e Impacto Financeiro (Proposta)", expanded=True):
                        st.success(f"Proposta gerada para **{nome_destino_final}**!")
                        
                        # PREPARAÇÃO DA ABRANGÊNCIA TOTAL DO DESTINO (Objetivo 2)
                        if tipo_destino == "Um Leve Existente (já selecionado)":
                            df_abrangencia_existente = df_abrangencia[df_abrangencia['LMC Name'] == nome_destino_final].copy()
                            df_abrangencia_existente['Observação'] = "Abrangência Atual"
                        else:
                            df_abrangencia_existente = pd.DataFrame(columns=['LMC Name', 'Região de preço', 'Cidade', 'State', 'Observação'])

                        df_movidos_fmt = df_movidos.copy()
                        df_movidos_fmt['Observação'] = "Migrado de: " + df_movidos_fmt['LMC Name']

                        # Combina as cidades que ele já atendia com as que ele recebeu agora
                        df_escopo_final = pd.concat([
                            df_abrangencia_existente[['Cidade', 'State', 'Região de preço', 'Observação']],
                            df_movidos_fmt[['Cidade', 'State', 'Região de preço', 'Observação']]
                        ], ignore_index=True).drop_duplicates(subset=['Cidade'])
                        
                        # Calcula o novo SLO para toda a abrangência
                        slo_base_dest = df_slos_clean[df_slos_clean['Cidade'] == cidade_base_destino]['SLO'].values
                        slo_base_dest_val = slo_base_dest[0] if len(slo_base_dest) > 0 else 0
                        
                        df_escopo_final = df_escopo_final.merge(df_slos_clean[['Cidade', 'SLO']], on='Cidade', how='left')
                        df_escopo_final['Novo SLO Local'] = df_escopo_final['SLO'] - slo_base_dest_val
                        df_escopo_final['Novo SLO Local'] = df_escopo_final['Novo SLO Local'].apply(lambda x: x if pd.notnull(x) and x > 0 else 0).astype(int)
                        
                        colunas_finais_abrangencia = ['Cidade', 'State', 'Região de preço', 'Novo SLO Local', 'Observação']
                        
                        # --- CÁLCULO DAS TABELAS FRETE PESO E IMPACTO ---
                        regioes_finais_destino = df_escopo_final['Região de preço'].unique()
                        regioes_movimentadas = df_movidos['Região de preço'].unique()
                        
                        lista_novas_tabelas = []
                        texto_explicativo = []
                        texto_explicativo.append(f"=== MEMORIAL DE CÁLCULO - PROPOSTA FRETE PESO ===\n")
                        texto_explicativo.append(f"Destino: {nome_destino_final}")
                        texto_explicativo.append(f"Estratégia Escolhida: {estrategia_preco}\n")
                        
                        custo_anterior_total = 0
                        custo_novo_total = 0
                        
                        for regiao in regioes_finais_destino:
                            texto_explicativo.append(f"--------------------------------------------------")
                            texto_explicativo.append(f"📍 REGIÃO: {regiao}")
                            texto_explicativo.append(f"--------------------------------------------------")
                            
                            cidades_da_regiao = df_escopo_final[df_escopo_final['Região de preço'] == regiao]['Cidade'].unique()
                            municipios_str = ", ".join(sorted([str(c).title() for c in cidades_da_regiao]))
                            
                            # Se a região sofreu alteração (recebeu cidades novas)
                            if regiao in regioes_movimentadas:
                                if estrategia_preco == "Gerar Tabela Equivalente (Focada em manter Impacto Neutro)":
                                    soma_produto_on_time = 0
                                    soma_produto_out_time = 0
                                    soma_volumes = 0
                                    
                                    texto_explicativo.append(">> PASSO 1: Somatório de Volumes (Apenas Faixa 1: 0 a 300g)\n")
                                    
                                    # 1. Peso das cidades movidas (Origem)
                                    texto_explicativo.append("Cidades Migradas (Origem):")
                                    cidades_mov_regiao = df_movidos[df_movidos['Região de preço'] == regiao]
                                    for _, row in cidades_mov_regiao.iterrows():
                                        leve_origem = row['LMC Name']
                                        cidade_movida = row['Cidade']
                                        
                                        vol_data = df_volume[(df_volume['Leve'] == leve_origem) & 
                                                             (df_volume['Cidade'].astype(str).str.lower() == str(cidade_movida).lower()) & 
                                                             (df_volume['Faixa pesos'].astype(str).str.contains('01 0 to 300', case=False, na=False))]
                                        volume = vol_data['# Total Packages'].sum() if not vol_data.empty else 0
                                        
                                        frete_data = df_frete_clean[(df_frete_clean['LMC name'] == leve_origem) & (df_frete_clean['label'] == regiao)]
                                        faixa1 = frete_data[frete_data['Faixa de peso (g/m³)'].str.contains('01 0 to 300', na=False, case=False)]
                                        
                                        preco_on = faixa1['on time amount'].values[0] if not faixa1.empty else 0
                                        preco_out = faixa1['out of time amount'].values[0] if not faixa1.empty else 0
                                            
                                        peso_calc = volume if volume > 0 else 1
                                        soma_produto_on_time += (preco_on * peso_calc)
                                        soma_produto_out_time += (preco_out * peso_calc)
                                        soma_volumes += peso_calc
                                        
                                        texto_explicativo.append(f" - {str(cidade_movida).title()} ({leve_origem}): {volume} pct x {formatar_moeda(preco_on)} = {formatar_moeda(volume * preco_on)}")
                                    
                                    # 2. Peso das cidades existentes no Destino (se for Leve existente)
                                    if tipo_destino == "Um Leve Existente (já selecionado)":
                                        texto_explicativo.append("\nCidades Atuais do Destino:")
                                        cidades_exist = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == regiao]
                                        
                                        for _, row in cidades_exist.iterrows():
                                            cid_ext = row['Cidade']
                                            vol_dest_data = df_volume[(df_volume['Leve'] == nome_destino_final) & 
                                                                      (df_volume['Cidade'].astype(str).str.lower() == str(cid_ext).lower()) & 
                                                                      (df_volume['Faixa pesos'].astype(str).str.contains('01 0 to 300', case=False, na=False))]
                                            volume_dest = vol_dest_data['# Total Packages'].sum() if not vol_dest_data.empty else 0
                                            
                                            if volume_dest > 0:
                                                frete_dest = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao)]
                                                faixa1_dest = frete_dest[frete_dest['Faixa de peso (g/m³)'].str.contains('01 0 to 300', na=False, case=False)]
                                                
                                                preco_dest_on = faixa1_dest['on time amount'].values[0] if not faixa1_dest.empty else 0
                                                preco_dest_out = faixa1_dest['out of time amount'].values[0] if not faixa1_dest.empty else 0
                                                
                                                soma_produto_on_time += (preco_dest_on * volume_dest)
                                                soma_produto_out_time += (preco_dest_out * volume_dest)
                                                soma_volumes += volume_dest
                                                
                                                texto_explicativo.append(f" - {str(cid_ext).title()}: {volume_dest} pct x {formatar_moeda(preco_dest_on)} = {formatar_moeda(volume_dest * preco_dest_on)}")

                                    nova_faixa1_on = soma_produto_on_time / soma_volumes if soma_volumes > 0 else 0
                                    nova_faixa1_out = soma_produto_out_time / soma_volumes if soma_volumes > 0 else 0
                                    
                                    texto_explicativo.append(f"\n>> PASSO 2: Preço Equivalente (Faixa 1)")
                                    texto_explicativo.append(f"Custo Total Somado / Total Pacotes = {formatar_moeda(soma_produto_on_time)} / {soma_volumes}")
                                    texto_explicativo.append(f"Preço Equivalente Faixa 1: {formatar_moeda(nova_faixa1_on)}")
                                    
                                else:
                                    # Usar tabela existente
                                    frete_base = df_frete_clean[(df_frete_clean['LMC name'] == tabela_base_selecionada) & (df_frete_clean['label'] == regiao)]
                                    faixa1_base = frete_base[frete_base['Faixa de peso (g/m³)'].str.contains('01 0 to 300', na=False, case=False)]
                                    nova_faixa1_on = faixa1_base['on time amount'].values[0] if not faixa1_base.empty else 0
                                    nova_faixa1_out = faixa1_base['out of time amount'].values[0] if not faixa1_base.empty else 0
                                    
                                    texto_explicativo.append(f">> Tabela Existente Selecionada: {tabela_base_selecionada}")
                                    texto_explicativo.append(f"Valor copiado da Faixa 1: {formatar_moeda(nova_faixa1_on)}")
                                
                                # GERAÇÃO DA TABELA (23 FAIXAS)
                                base_on = nova_faixa1_on / 0.83
                                base_out = nova_faixa1_out / 0.83
                                
                                texto_explicativo.append(f"\n>> PASSO 3: Geração das 23 Faixas")
                                texto_explicativo.append(f"Base de Cálculo (Índice 1.00) = Faixa 1 / 0.83 = {formatar_moeda(base_on)}\n")
                                
                                df_regiao_tabela = df_price_var_clean.copy()
                                df_regiao_tabela['Região de Preço'] = regiao
                                df_regiao_tabela['Valor dentro do prazo'] = df_regiao_tabela['Multiplicador'] * base_on
                                df_regiao_tabela['Valor fora do prazo'] = df_regiao_tabela['Multiplicador'] * base_out
                                
                                for _, prow in df_regiao_tabela.iterrows():
                                    fx = prow['Faixa de peso (g/m³)']
                                    mult = prow['Multiplicador']
                                    val = prow['Valor dentro do prazo']
                                    texto_explicativo.append(f"{fx}: Mult {mult:.2f} x {formatar_moeda(base_on)} = {formatar_moeda(val)}")
                                
                                # --- CÁLCULO DE IMPACTO FINANCEIRO DESTA REGIÃO ---
                                # 1. Impacto das cidades movidas
                                for _, row in cidades_mov_regiao.iterrows():
                                    leve_orig = row['LMC Name']
                                    cid_movida = row['Cidade']
                                    
                                    vols_cidade = df_volume[(df_volume['Leve'] == leve_orig) & (df_volume['Cidade'].astype(str).str.lower() == str(cid_movida).lower())]
                                    for _, v_row in vols_cidade.iterrows():
                                        faixa_peso = v_row['Faixa pesos']
                                        qtd_pacotes = v_row['# Total Packages']
                                        
                                        tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == leve_orig) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == faixa_peso)]
                                        preco_antigo = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                        
                                        tb_nova = df_regiao_tabela[df_regiao_tabela['Faixa de peso (g/m³)'] == faixa_peso]
                                        preco_novo = tb_nova['Valor dentro do prazo'].values[0] if not tb_nova.empty else 0
                                        
                                        custo_anterior_total += (qtd_pacotes * preco_antigo)
                                        custo_novo_total += (qtd_pacotes * preco_novo)
                                        
                                # 2. Impacto das cidades que já existiam no destino (Elas também sofrem alteração de preço)
                                if tipo_destino == "Um Leve Existente (já selecionado)":
                                    cidades_exist = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == regiao]
                                    for _, row in cidades_exist.iterrows():
                                        cid_ext = row['Cidade']
                                        vols_cidade = df_volume[(df_volume['Leve'] == nome_destino_final) & (df_volume['Cidade'].astype(str).str.lower() == str(cid_ext).lower())]
                                        
                                        for _, v_row in vols_cidade.iterrows():
                                            faixa_peso = v_row['Faixa pesos']
                                            qtd_pacotes = v_row['# Total Packages']
                                            
                                            tb_antiga = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao) & (df_frete_clean['Faixa de peso (g/m³)'] == faixa_peso)]
                                            preco_antigo = tb_antiga['on time amount'].values[0] if not tb_antiga.empty else 0
                                            
                                            tb_nova = df_regiao_tabela[df_regiao_tabela['Faixa de peso (g/m³)'] == faixa_peso]
                                            preco_novo = tb_nova['Valor dentro do prazo'].values[0] if not tb_nova.empty else 0
                                            
                                            custo_anterior_total += (qtd_pacotes * preco_antigo)
                                            custo_novo_total += (qtd_pacotes * preco_novo)
                                
                                texto_explicativo.append("\n")

                            else:
                                # Região que já era do Leve e não recebeu nenhuma cidade nova (Tabela Original mantida)
                                texto_explicativo.append("Nenhuma movimentação nesta região. Tabela atual do Leve mantida.\n")
                                df_regiao_tabela = df_frete_clean[(df_frete_clean['LMC name'] == nome_destino_final) & (df_frete_clean['label'] == regiao)].copy()
                                df_regiao_tabela.rename(columns={'label': 'Região de Preço', 'on time amount': 'Valor dentro do prazo', 'out of time amount': 'Valor fora do prazo'}, inplace=True)
                            
                            df_regiao_tabela['Municípios'] = municipios_str
                            df_regiao_tabela = df_regiao_tabela[['Região de Preço', 'Municípios', 'Faixa de peso (g/m³)', 'Valor dentro do prazo', 'Valor fora do prazo']]
                            lista_novas_tabelas.append(df_regiao_tabela)
                        
                        df_tabela_final = pd.concat(lista_novas_tabelas, ignore_index=True)
                        texto_completo_memorial = "\n".join(texto_explicativo)
                        
                        # --- EXIBIÇÃO DO IMPACTO FINANCEIRO ---
                        impacto_financeiro = custo_novo_total - custo_anterior_total
                        
                        st.subheader("📊 Impacto Financeiro (30 dias)")
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Custo Antigo (Todas as cidades afetadas)", formatar_moeda(custo_anterior_total))
                        col2.metric("Novo Custo Projetado", formatar_moeda(custo_novo_total))
                        
                        if impacto_financeiro > 0:
                            col3.metric("Impacto Financeiro", formatar_moeda(impacto_financeiro), f"+{formatar_moeda(impacto_financeiro)} (Aumento de custo)", delta_color="inverse")
                        elif impacto_financeiro < 0:
                            col3.metric("Impacto Financeiro", formatar_moeda(impacto_financeiro), f"{formatar_moeda(impacto_financeiro)} (Economia)", delta_color="normal")
                        else:
                            col3.metric("Impacto Financeiro", "R$ 0,00", "Neutro")
                            
                        with st.expander("🤔 Detalhes do Cálculo & Proporção das Tabelas", expanded=False):
                            st.markdown("Baixe o memorial descritivo completo para ver exatamente os volumes utilizados, o cálculo da média ponderada e a aplicação da curva padrão nas 23 faixas.")
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
                        st.markdown("### 📥 Download da Proposta Final")
                        
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

else:
    st.info("Por favor, faça o upload de **todas as 4 bases** na barra lateral para prosseguir.")