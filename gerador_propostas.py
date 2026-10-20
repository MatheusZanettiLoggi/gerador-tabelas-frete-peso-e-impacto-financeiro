import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import io
import base64
import json
import zipfile
import unicodedata
import uuid
from datetime import datetime, timezone, timedelta
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Tenta importar as bibliotecas de PDF
try:
    import weasyprint
    HAS_PDF_GENERATOR = True
except ImportError:
    HAS_PDF_GENERATOR = False

st.set_page_config(page_title="Gerador de Propostas - Leves", layout="wide")

# --- INICIALIZAÇÃO DE ESTADOS E CORES ---
if "app_mode" not in st.session_state:
    st.session_state.app_mode = "start"
if "num_cenarios" not in st.session_state:
    st.session_state["num_cenarios"] = 1

CORES_CENARIOS = ['#e6f2ff', '#e6ffe6', '#fff2e6', '#f2e6ff', '#ffe6e6']

# --- TELA INICIAL (LANDING PAGE) ---
if st.session_state.app_mode == "start":
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; color: #002766;'>📦 Gerador de Propostas e Movimentação de Leves</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 1.1em; color: gray;'>Selecione o modo de inicialização do projeto abaixo.</p><br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 4, 1])
    with col2:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
            <div style="background-color: #111e2d; border-radius: 8px; padding: 20px; border-left: 5px solid #3b82f6; height: 120px;">
                <h4 style="color: #60a5fa; margin-top: 0;">✨ Nova Análise</h4>
                <p style="color: #bfdbfe; font-size: 0.95em;">Inicie um projeto do zero, importando as bases do Looker.</p>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Iniciar Nova Análise", use_container_width=True):
                st.session_state.app_mode = "new"
                st.rerun()
                
        with c2:
            st.markdown("""
            <div style="background-color: #112a1f; border-radius: 8px; padding: 20px; border-left: 5px solid #10b981; height: 120px;">
                <h4 style="color: #34d399; margin-top: 0;">📁 Carregar Análise Passada</h4>
                <p style="color: #a7f3d0; font-size: 0.95em;">Continue exatamente de onde parou importando seu backup (.zip).</p>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Carregar Backup (.zip)", use_container_width=True):
                st.session_state.app_mode = "load"
                st.rerun()
                
    st.markdown("<br><hr><div style='text-align: center;'><small>Desenvolvido por <b>Matheus Zanetti</b></small></div>", unsafe_allow_html=True)
    st.stop()

# --- TELA DE CARREGAMENTO DE BACKUP ---
if st.session_state.app_mode == "load":
    st.title("📁 Carregar Backup de Análise")
    st.warning("⚠️ **Aviso:** Caso a análise passada tenha sido realizada há alguns dias, fique atento a mudanças de abrangência e de tabela frete peso em leves que possam ter sido realizadas neste meio tempo. Caso alguma mudança tenha sido feita, recomenda-se começar uma nova análise com dados atualizados.")
    
    uploaded_zip = st.file_uploader("Selecione o arquivo de backup (.zip)", type=["zip"])
    
    if uploaded_zip is not None:
        with st.spinner("Restaurando ambiente..."):
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                st.session_state.df_frete_std = pd.read_csv(zip_ref.open("frete.csv"))
                st.session_state.df_abrangencia_std = pd.read_csv(zip_ref.open("abrangencia.csv"))
                st.session_state.df_volume_std = pd.read_csv(zip_ref.open("volume.csv"))
                
                if "movimentacao.csv" in zip_ref.namelist():
                    st.session_state.df_movimentacao = pd.read_csv(zip_ref.open("movimentacao.csv"))
                    
                state_dict = json.loads(zip_ref.read("state.json").decode("utf-8"))
                for k, v in state_dict.items():
                    if not str(k).startswith(("btn_", "dl_")):
                        st.session_state[k] = v
                    
        st.session_state.app_mode = "run"
        st.rerun()
        
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅️ Voltar"):
        st.session_state.app_mode = "start"
        st.rerun()
    st.stop()

# --- O CÓDIGO PRINCIPAL CONTINUA AQUI (MODOS 'NEW' E 'RUN') ---
st.title("📦 Gerador de Propostas e Movimentação de Leves")
st.markdown("Configure os parâmetros da sua negociação abaixo.")

diagnostic_log = [f"--- DIAGNÓSTICO DO SISTEMA LOGGI ---", f"Data/Hora: {datetime.now(timezone(timedelta(hours=-3)))}", f"Streamlit Versão: {st.__version__}\n"]
def add_log(msg): diagnostic_log.append(str(msg))

@st.cache_data
def load_data(file):
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file)

@st.cache_data
def load_local_excel(filename):
    if not os.path.exists(filename): return None
    return pd.read_excel(filename)

def normalize_string(s):
    if pd.isna(s): return ""
    s = str(s).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def extrair_estado(nome_leve):
    match = re.search(r'-\s*([A-Z]{2})\b', nome_leve)
    if match: return match.group(1)
    return None

def formatar_moeda(valor):
    try: return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except: return str(valor)

EQUIVALENCIAS_REGIOES = {
    "AC RED": "AC INT 2", "AL RED": "AL INT 3", "AM RED": "AM INT", "AP RED": "AP INT 3",
    "BA RED": "BA INT 5", "CE RED": "CE INT 2", "ES RED": "ES INT 3", "GO RED": "GO INT 4",
    "MA RED": "MA INT 2", "MG RED": "MG INT 4", "MS RED": "MS INT 2", "MT RED": "MT INT 2",
    "PA RED": "PA INT 3", "PE RED": "PE INT 3", "PE RED B": "PE INT 3", "PI RED": "PI INT 2",
    "PR RED": "PR INT 4", "RJ RED": "RJ INT 3", "RN RED": "RN INT 2", "RO RED": "RO INT 3",
    "RS RED": "RS INT 4", "SC RED": "SC INT 4", "SE RED": "SE INT 2", "SP RED": "SP INT 6",
    "TO RED": "TO INT 3"
}

def limpar_regiao(val):
    if pd.isna(val): return val
    s = str(val).strip()
    if s in EQUIVALENCIAS_REGIOES:
        s = EQUIVALENCIAS_REGIOES[s]
    if s.endswith(" B"):
        s = s[:-2].strip()
    return s

def padronizar_colunas_frete_raw(df):
    mapa = {"Leve Contract Region LMC name": "LMC name", "Leve Contract Region table name": "table name", "Leve Contract Region label": "label", "Leve Contract Region on time amount": "on time amount", "Leve Contract Region out of time amount": "out of time amount", "Leve Contract Region service type": "service type", "Leve Contract Region Faixa de peso (g/m³)": "Faixa de peso cubado (g)", "Faixa de peso (g/m³)": "Faixa de peso cubado (g)"}
    return df.rename(columns=mapa)

def padronizar_colunas_abrangencia_raw(df):
    mapa = {"Territorial Scope Pricing Regions LMC Name": "LMC Name", "Territorial Scope Pricing Regions Pricing Region": "Região de preço 2023", "Territorial Scope Pricing Regions City": "Cidade", "Territorial Scope Pricing Regions State": "State", "Territorial Scope Pricing Regions Service Type": "Tipo de serviço", "Territorial Scope Pricing Regions SLO Lastmile": "Prazo adicional"}
    return df.rename(columns=mapa)

def padronizar_colunas_volume_raw(df):
    mapa = {"Package Charge Leve Last Mile Company Name": "Leve", "Distribution and Expedition Center Locations Routing Code": "Routing Code", "Package Charge Leve Region label": "Região de preço", "Package Charge Leve Region Label": "Região de preço", "Package Destination City": "Cidade", "Package Charge Leve Service Charge Type": "Service Charge Type", "Package Charge Leve # Packages": "# Total Packages", "Faixa pesos": "Faixa de peso cubado (g)", "Faixa Pesos": "Faixa de peso cubado (g)", "Package Charge Leve Faixa pesos": "Faixa de peso cubado (g)"}
    return df.rename(columns=mapa)

def padronizar_colunas_frete(df):
    df = padronizar_colunas_frete_raw(df)
    if 'label' in df.columns: df['label'] = df['label'].apply(limpar_regiao)
    return df

def padronizar_colunas_abrangencia(df):
    df = padronizar_colunas_abrangencia_raw(df)
    if 'Região de preço 2023' in df.columns: df['Região de preço 2023'] = df['Região de preço 2023'].apply(limpar_regiao)
    return df

def padronizar_colunas_volume(df):
    df = padronizar_colunas_volume_raw(df)
    if 'Região de preço' in df.columns: df['Região de preço'] = df['Região de preço'].apply(limpar_regiao)
    return df

@st.cache_data
def processar_frete(df_frete):
    return df_frete.drop_duplicates(subset=['LMC name', 'label', 'Faixa de peso cubado (g)'], keep='first')

@st.cache_data
def processar_slos_novo(df):
    if len(df.columns) > 10:
        col_cidade = df.columns[3]; col_regiao = df.columns[5]; col_prazo = df.columns[10]
        mask_express = df.apply(lambda row: row.astype(str).str.contains('Loggi Express', case=False).any(), axis=1)
        df_filtered = df[mask_express].copy()
        df_filtered['State'] = df_filtered[col_regiao].astype(str).str.strip().str[:2].str.upper()
        df_filtered['Cidade_str'] = df_filtered[col_cidade].astype(str).str.strip()
        df_filtered['Chave'] = df_filtered.apply(lambda r: normalize_string(r['Cidade_str']) + str(r['State']).lower(), axis=1)
        df_filtered['SLO'] = pd.to_numeric(df_filtered[col_prazo], errors='coerce')
        df_grouped = df_filtered.groupby(['State', 'Cidade_str', 'Chave'], as_index=False)['SLO'].min()
        df_grouped.rename(columns={'Cidade_str': 'Cidade'}, inplace=True)
        return df_grouped
    return pd.DataFrame(columns=['State', 'Cidade', 'Chave', 'SLO'])

@st.cache_data
def processar_nomes_leves(df_volume):
    mapping = df_volume[['Leve', 'Routing Code']].drop_duplicates().dropna()
    mapping['nome_completo'] = mapping['Leve'] + " (" + mapping['Routing Code'] + ")"
    return mapping

@st.cache_data
def processar_price_var(df_price):
    df_price.columns = ['Faixa de peso cubado (g)', 'Multiplicador']
    df_price['Cod'] = df_price['Faixa de peso cubado (g)'].astype(str).str.strip().str[:2]
    return df_price

@st.cache_data
def get_backup_zip_cached(state_json, frete_csv, abr_csv, vol_csv, mov_csv):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        zip_file.writestr("frete.csv", frete_csv)
        zip_file.writestr("abrangencia.csv", abr_csv)
        zip_file.writestr("volume.csv", vol_csv)
        if mov_csv is not None:
            zip_file.writestr("movimentacao.csv", mov_csv)
        zip_file.writestr("state.json", state_json)
    return zip_buffer.getvalue()

def formatar_excel_proposta(writer):
    workbook = writer.book
    font_padrao = Font(name='Inter', size=10)
    font_cabecalho = Font(name='Inter', size=10, bold=True)
    alinhamento = Alignment(horizontal='center', vertical='center', wrap_text=False)
    borda_cinza = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'A2'
        col_formats = {}
        for row in ws.iter_rows(min_row=1, max_row=50):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val_str = cell.value.lower()
                    if "(r$)" in val_str or "valor" in val_str or "tarifa" in val_str: col_formats[cell.column_letter] = '"R$" #,##0.00'
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 18
            for cell in row:
                if cell.value is not None:
                    if cell.row == 1: cell.font = font_cabecalho
                    else: cell.font = font_padrao
                    if cell.column_letter in col_formats and type(cell.value) in [int, float]: cell.number_format = col_formats[cell.column_letter]
                    cell.alignment = alinhamento
                    cell.border = borda_cinza
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column].width = max((max_length + 2) * 1.15, 12)

def formatar_excel_resumo(writer, cenarios_nomes):
    workbook = writer.book
    font_padrao = Font(name='Inter', size=10)
    font_cabecalho = Font(name='Inter', size=10, bold=True, color='FFFFFF')
    fill_cabecalho = PatternFill(start_color='002766', end_color='002766', fill_type='solid')
    font_destaque_red = Font(name='Inter', size=10, color='9C0006', bold=True)
    fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    font_total = Font(name='Inter', size=10, color='FFFFFF', bold=True)
    fill_total = PatternFill(start_color='002766', end_color='002766', fill_type='solid')
    alinhamento = Alignment(horizontal='center', vertical='center', wrap_text=False)
    borda_cinza = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

    CORES_HEX = [c.replace('#', '') for c in CORES_CENARIOS]

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'A2'
        header_map = {} 
        for row in ws.iter_rows():
            r_idx = row[0].row
            first_val = str(row[0].value).strip() if row[0].value is not None else ""
            is_header = (first_val in ['Região de Preço', 'Cenário'])
            is_total = ("Total" in first_val or "Total Geral" in first_val)
            if is_header: header_map[r_idx] = {cell.column: str(cell.value) for cell in row}
            current_header_row = 1
            for h_r in sorted(header_map.keys(), reverse=True):
                if r_idx >= h_r:
                    current_header_row = h_r
                    break
            for cell in row:
                if cell.value is None: continue
                c_idx = cell.column
                col_name = header_map.get(current_header_row, {}).get(c_idx, "")
                col_name_lower = col_name.lower()
                cell.alignment = alinhamento
                cell.border = borda_cinza
                if any(x in col_name_lower for x in ['fat', 'ticket', 'tk', 'impacto', 'atual', '(r$)', 'tarifa', 'custo', 'diferença', 'projetado', '1ª fx', '1ª faixa']):
                    if "volumetria" not in col_name_lower and "pacotes" not in col_name_lower: cell.number_format = '"R$" #,##0.00'
                if "%" in col_name_lower or "aum" in col_name_lower or "comercial" in col_name_lower: cell.number_format = '0.00%'
                if "volumetria" in col_name_lower or "volume" in col_name_lower or "pacotes" in col_name_lower: cell.number_format = '#,##0'
                cell.font = font_padrao
                if is_header:
                    bg_applied = False
                    for i, cen in enumerate(cenarios_nomes):
                        if cen in col_name or f"Cenário {i+1}" in col_name:
                            cell.fill = PatternFill(start_color=CORES_HEX[i % len(CORES_HEX)], end_color=CORES_HEX[i % len(CORES_HEX)], fill_type='solid')
                            cell.font = Font(name='Inter', size=10, bold=True, color='000000')
                            bg_applied = True
                            break
                    if not bg_applied:
                        cell.fill = fill_cabecalho
                        cell.font = font_cabecalho
                elif is_total:
                    if "impacto" in col_name_lower or "aum" in col_name_lower:
                        cell.font = font_destaque_red
                        cell.fill = fill_red
                    else:
                        cell.font = font_total
                        cell.fill = fill_total
                else:
                    if sheet_name == 'Resumo de Cenários' or r_idx < 10:
                        for i, cen in enumerate(cenarios_nomes):
                            if cen in col_name or f"Cenário {i+1}" in col_name:
                                cell.fill = PatternFill(start_color=CORES_HEX[i % len(CORES_HEX)], end_color=CORES_HEX[i % len(CORES_HEX)], fill_type='solid')
                                break
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column].width = max((max_length + 2) * 1.15, 12)

# --- FUNÇÕES GLOBAIS DE CONSTRUÇÃO DE HTML PARA PDF ---
def generate_html_table_styled(df, cenarios_nomes):
    if df is None or df.empty: return "<p>Sem dados.</p>"
    html = "<table><thead><tr>"
    for col in df.columns:
        bg_color = "#002766" 
        text_color = "#ffffff"
        for i, cen in enumerate(cenarios_nomes):
            if cen in col or f"Cenário {i+1}" in col:
                bg_color = CORES_CENARIOS[i % len(CORES_CENARIOS)]
                text_color = "#000000"
                break
        html += f"<th style='background-color: {bg_color} !important; color: {text_color} !important;'>{col}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df.iterrows():
        is_total = str(row.iloc[0]).lower() == 'total geral'
        html += "<tr>"
        for col_idx, val in enumerate(row):
            col_name = df.columns[col_idx]
            bg_color = "#ffffff"
            text_color = "#333333"
            font_weight = "normal"
            if is_total:
                font_weight = "bold"
                if "Impacto" in col_name or "% Aum" in col_name:
                    bg_color = "#ffc7ce"
                    text_color = "#9c0006"
                else:
                    bg_color = "#002766"
                    text_color = "#ffffff"
            else:
                for i, cen in enumerate(cenarios_nomes):
                    if cen in col_name or f"Cenário {i+1}" in col_name:
                        bg_color = CORES_CENARIOS[i % len(CORES_CENARIOS)]
                        text_color = "#000000"
                        break
            style_str = f"background-color: {bg_color}; color: {text_color}; font-weight: {font_weight};" if bg_color or is_total else f"color: {text_color};"
            html += f"<td style='{style_str}'>{val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

def generate_simple_html_table(df):
    if df is None or df.empty: return "<p>Sem dados.</p>"
    html = "<table><thead><tr>"
    for col in df.columns: html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df.iterrows():
        html += "<tr>"
        for val in row: html += f"<td>{val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

def generate_html_pdf(nome_destino, estrategia, cidades_movimentadas_str, df_comparativo, cenario_metrics, df_abrangencia_out, dict_tabelas_out, tabelas_atuais_pdf, cenarios_nomes):
    fuso_brasilia = datetime.now(timezone(timedelta(hours=-3)))
    data_extracao = fuso_brasilia.strftime("%d/%m/%Y às %H:%M")
    
    def format_money_local(val): return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    def format_perc_local(val): return f"{val:+.2f}%"
    
    def get_indicator(val, is_cost=False):
        if val > 0: return f'<span class="{"arrow-up-red" if is_cost else "arrow-up-green"}">▲ +{format_money_local(val)}</span>'
        elif val < 0: return f'<span class="{"arrow-down-green" if is_cost else "arrow-down-red"}">▼ -{format_money_local(abs(val))}</span>'
        return f'<span class="arrow-neutral">■ R$ 0,00</span>'
        
    def get_perc_indicator(val, is_cost=False):
        if val > 0: return f'<span class="{"arrow-up-red" if is_cost else "arrow-up-green"}">(+{format_perc_local(val)})</span>'
        elif val < 0: return f'<span class="{"arrow-down-green" if is_cost else "arrow-down-red"}">({format_perc_local(val)})</span>'
        return f'<span class="arrow-neutral">(0.00%)</span>'

    logo_html = ""
    logo_path = "logo.png"
    if not os.path.exists(logo_path) and os.path.exists("logo.png.png"): logo_path = "logo.png.png"
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            logo_html = f'<img src="data:image/png;base64,{encoded_string}" style="position: fixed; top: -20mm; right: 0; width: 45px; height: auto; z-index: 1000;">'

    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Relatório de Simulação - Loggi</title>
        <style>
            @page {{ size: A4 landscape; margin: 30mm 15mm 20mm 15mm; background-color: #ffffff; @bottom-center {{ content: "Simulador de Movimentação de Leves - Desenvolvido por Matheus Zanetti | Página " counter(page); font-family: 'Montserrat', sans-serif; font-size: 8pt; color: #888888; font-style: italic; }} }}
            body {{ font-family: 'Montserrat', sans-serif; margin: 0; padding: 0; color: #000; background-color: #ffffff; font-size: 9.5pt; line-height: 1.4; }}
            *, *::before, *::after {{ box-sizing: border-box; }}
            h1 {{ color: #002766; font-size: 16pt; text-align: center; margin-top: 10px; margin-bottom: 5px; text-transform: uppercase; }}
            h2 {{ color: #006aff; font-size: 12pt; border-bottom: 2px solid #00baff; padding-bottom: 4px; margin-top: 25px; margin-bottom: 10px; }}
            h3 {{ color: #002766; font-size: 11pt; margin-top: 15px; margin-bottom: 8px; }}
            .header-meta {{ text-align: center; color: #666; font-size: 8.5pt; margin-bottom: 25px; }}
            .card-container {{ display: block; width: 100%; margin-bottom: 15px; page-break-inside: avoid; }}
            .card {{ background-color: #fff; border: 1px solid #e0e0e0; border-radius: 6px; padding: 12px; }}
            .metric-row {{ display: block; width: 100%; margin-bottom: 6px; }}
            .metric-label {{ display: inline-block; width: 180px; color: #555; }}
            .metric-value {{ display: inline-block; font-weight: bold; color: #000; }}
            .metric-sub {{ color: #777; font-size: 8pt; margin-left: 10px; }}
            .arrow-up-green {{ color: #09ab3b; font-weight: bold; }}
            .arrow-down-green {{ color: #09ab3b; font-weight: bold; }}
            .arrow-up-red {{ color: #ff4b4b; font-weight: bold; }}
            .arrow-down-red {{ color: #ff4b4b; font-weight: bold; }}
            .arrow-neutral {{ color: #888888; font-weight: bold; }}
            .region-block {{ background-color: #fff; border: 1px solid #e0e0e0; border-radius: 4px; padding: 10px; margin-bottom: 8px; page-break-inside: avoid; }}
            .region-title {{ font-weight: bold; color: #006aff; margin-bottom: 6px; font-size: 10pt; }}
            .region-alert {{ color: #e67e22; font-weight: bold; font-size: 8.5pt; margin-bottom: 6px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 8pt; text-align: center; page-break-inside: avoid; }}
            th {{ background-color: #002766; color: #fff; font-weight: bold; padding: 6px 4px; border: 1px solid #002766; }}
            td {{ padding: 5px 4px; border: 1px solid #e0e0e0; color: #333; }}
            tr:nth-child(even) {{ background-color: #f4f8fb; }}
            .page-break {{ page-break-before: always; }}
        </style>
    </head>
    <body>
    {logo_html}
        <h1>Relatório Comparativo de Cenários</h1>
        <div class="header-meta">Gerado em: {data_extracao} | Lead/Destino: {nome_destino}</div>
    """

    html_content += f"""<h2>1. RESUMO COMPARATIVO DE CENÁRIOS</h2>"""
    df_resumo_html = df_comparativo.copy()
    for c in df_resumo_html.columns:
        if any(term in c for term in ["Fat", "Ticket", "TK", "Impacto", "Atual", "1ª Fx", "1ª Faixa"]):
            if c != 'Região de Preço' and c != 'Volumetria' and "%" not in c:
                df_resumo_html[c] = df_resumo_html[c].apply(lambda x: format_money_local(x) if x != "-" and pd.notna(x) else "-")
        elif "%" in c or "Aum" in c: df_resumo_html[c] = df_resumo_html[c].apply(lambda x: format_perc_local(x * 100) if pd.notna(x) and isinstance(x, (int, float)) else x)
        elif "Vol" in c: df_resumo_html[c] = df_resumo_html[c].apply(lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "-")

    html_content += generate_html_table_styled(df_resumo_html, cenarios_nomes)

    for cenario_nome, metricas in cenario_metrics.items():
        fat_parc_ant = format_money_local(metricas.get('fat_parceiro_antigo', metricas['fat_antigo']))
        vol_parc_ant = int(metricas.get('vol_parceiro_antigo', metricas['vol_fat_antigo']))
        tk_parc_ant = format_money_local(metricas.get('tk_parceiro_antigo', metricas['tk_fat_antigo']))
        cresc_fat_val = metricas.get('cresc_fat_parceiro', metricas['cresc_fat'])
        perc_cresc_val = metricas.get('perc_cresc_parceiro', metricas['perc_cresc'])
        
        migrated_html = ""
        if metricas.get('municipios_migrados_nomes'):
            mun_str = ", ".join(metricas['municipios_migrados_nomes'])
            vol_mig = int(metricas['vol_migrado'])
            fat_mig = format_money_local(metricas['fat_migrado_projetado'])
            migrated_html = f"<div style='background-color: #e6f2ff; padding: 10px; border-left: 4px solid #006aff; margin-top: 10px; font-size: 8.5pt;'><b>Cidades Adicionadas:</b> Estamos adicionando os municípios de <b>{mun_str}</b>, que representam um volume de <b>{vol_mig:,} pacotes</b> a mais e um faturamento projetado adicional de <b>{fat_mig}</b>, compondo o Novo Faturamento Projetado.</div>"

        html_content += f"""<div class="page-break"></div><h2>DETALHAMENTO: {cenario_nome.upper()}</h2>"""
        html_content += f"""
            <h3>Visão do Parceiro (Faturamento do Leve)</h3>
            <div class="card-container">
                <div class="card">
                    <div class="metric-row"><span class="metric-label">Faturamento Atual:</span><span class="metric-value">{fat_parc_ant}</span><span class="metric-sub">(Vol: {vol_parc_ant:,} | TK: {tk_parc_ant})</span></div>
                    <div class="metric-row"><span class="metric-label">Novo Faturamento:</span><span class="metric-value">{format_money_local(metricas['fat_novo'])}</span><span class="metric-sub">(Vol: {int(metricas['vol_fat_novo']):,} | TK: {format_money_local(metricas['tk_fat_novo'])})</span></div>
                    <div class="metric-row" style="margin-top: 5px; padding-top: 5px; border-top: 1px dashed #e0e0e0;"><span class="metric-label">Crescimento da Operação:</span><span class="metric-value">{get_indicator(cresc_fat_val, False)} {get_perc_indicator(perc_cresc_val, False)}</span></div>
                </div>
                {migrated_html}
            </div>
            <h3>🐰 Visão LOGGI (Impacto Financeiro Real)</h3>
            <div class="card-container">
                <div class="card">
                    <div class="metric-row"><span class="metric-label">Custo Antigo Global:</span><span class="metric-value">{format_money_local(metricas['loggi_antigo'])}</span><span class="metric-sub">(Vol: {int(metricas['vol_loggi']):,} | TK: {format_money_local(metricas['tk_loggi_antigo'])})</span></div>
                    <div class="metric-row"><span class="metric-label">Novo Custo Projetado:</span><span class="metric-value">{format_money_local(metricas['loggi_novo'])}</span><span class="metric-sub">(Vol: {int(metricas['vol_loggi']):,} | TK: {format_money_local(metricas['tk_loggi_novo'])})</span></div>
                    <div class="metric-row" style="margin-top: 5px; padding-top: 5px; border-top: 1px dashed #e0e0e0;"><span class="metric-label">Impacto Financeiro:</span><span class="metric-value">{get_indicator(metricas['imp_loggi'], True)} {get_perc_indicator(metricas['perc_imp_loggi'], True)}</span></div>
                </div>
            </div>
        """
        
        detalhes_reg = metricas.get('detalhes_regioes', {})
        if detalhes_reg:
            html_content += """<h3>Impacto por Região</h3>"""
            for reg, dados in detalhes_reg.items():
                tk_ant = dados['custo_antigo'] / dados['vol'] if dados['vol'] > 0 else 0
                tk_nov = dados['custo_novo'] / dados['vol'] if dados['vol'] > 0 else 0
                imp_r = dados['custo_novo'] - dados['custo_antigo']
                perc_r = (imp_r / dados['custo_antigo']) * 100 if dados['custo_antigo'] > 0 else 0
                ajuste = dados.get('ajuste', 0.0)
                ajuste_html = f'<div class="region-alert">Aviso: Ajuste Comercial Aplicado.</div>' if ajuste != 0.0 else ''
                html_content += f"""
                <div class="region-block">
                    <div class="region-title">Região: {reg}</div>
                    {ajuste_html}
                    <div class="metric-row"><span class="metric-label">Custo Antigo:</span><span class="metric-value">{format_money_local(dados['custo_antigo'])}</span><span class="metric-sub">(TK: {format_money_local(tk_ant)})</span></div>
                    <div class="metric-row"><span class="metric-label">Novo Custo:</span><span class="metric-value">{format_money_local(dados['custo_novo'])}</span><span class="metric-sub">(TK: {format_money_local(tk_nov)})</span></div>
                    <div class="metric-row"><span class="metric-label">Variação no Budget:</span><span class="metric-value">{get_indicator(imp_r, True)} {get_perc_indicator(perc_r, True)}</span></div>
                </div>
                """
        html_content += f"<h3>Abrangência Completa Projetada ({cenario_nome})</h3>"
        colunas_abr_pdf = [c for c in df_abrangencia_out.columns if c != 'State']
        html_content += generate_simple_html_table(df_abrangencia_out[colunas_abr_pdf])
        html_content += f"<h3>Tabela Frete Peso Projetada ({cenario_nome})</h3>"
        df_ex = dict_tabelas_out[cenario_nome].copy()
        df_ex['Valor dentro do prazo'] = df_ex['Valor dentro do prazo'].apply(formatar_moeda)
        df_ex['Valor fora do prazo'] = df_ex['Valor fora do prazo'].apply(formatar_moeda)
        html_content += generate_simple_html_table(df_ex)
        html_content += f"""<div class="page-break"></div>"""

    if tabelas_atuais_pdf:
        html_content += f"""<h2>3. TABELAS FRETE PESO ATUAIS (ORIGENS ENVOLVIDAS)</h2>"""
        for leve_nome, df_tab in tabelas_atuais_pdf.items():
            html_content += f"<h3>Leve Atual: {leve_nome}</h3>"
            html_content += generate_simple_html_table(df_tab)

    html_content += """</body></html>"""
    pdf_bytes = weasyprint.HTML(string=html_content).write_pdf()
    return pdf_bytes

# --- DADOS E LÓGICA PRINCIPAL ---
data_ready = False
df_price_var_raw = load_local_excel("Price variation.xlsx")
df_slos_raw = load_local_excel("coverage_spreadsheet_tab_13_2026_04.xlsx")

if st.session_state.app_mode == "new":
    with st.sidebar:
        st.markdown("### Configuração Inicial")
        if st.button("⬅️ Voltar ao Início", use_container_width=True):
            st.session_state.app_mode = "start"
            st.rerun()
            
        with st.expander("1. Upload de Bases de Dados", expanded=True):
            st.markdown("**Tabelas frete peso praticadas**")
            st.markdown("[Link Looker: 26300](https://loggi.looker.com/looks/26300)")
            file_frete = st.file_uploader("Upload Frete Peso", type=["xlsx", "csv"], label_visibility="collapsed")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Abrangências atuais**")
            st.markdown("[Link Looker: 26301](https://loggi.looker.com/looks/26301)")
            file_abrangencia = st.file_uploader("Upload Abrangência", type=["xlsx", "csv"], label_visibility="collapsed")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**SLOs globais das cidades**")
            st.info("Arquivo carregado automaticamente (coverage_spreadsheet_tab_13_2026_04.xlsx)")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Volume de pacotes (30 dias)**")
            st.markdown("[Link Looker: 26302](https://loggi.looker.com/looks/26302)")
            file_volume = st.file_uploader("Upload Volume", type=["xlsx", "csv"], label_visibility="collapsed")
            
    if file_frete and file_abrangencia and file_volume:
        if df_price_var_raw is None:
            st.error("Erro: O arquivo 'Price variation.xlsx' não foi encontrado. Por favor, certifique-se de que ele foi subido para o repositório do GitHub.")
        elif df_slos_raw is None:
            st.error("Erro: O arquivo local 'coverage_spreadsheet_tab_13_2026_04.xlsx' não foi encontrado. Certifique-se de que ele está no repositório do GitHub.")
        else:
            with st.spinner("Carregando e processando bases..."):
                df_frete = load_data(file_frete)
                df_abrangencia = load_data(file_abrangencia)
                df_volume = load_data(file_volume)
                
                # Para carregar preservando os nomes " B" e salvar a base de dados raw na memoria:
                df_frete = padronizar_colunas_frete_raw(df_frete)
                df_abrangencia = padronizar_colunas_abrangencia_raw(df_abrangencia)
                df_volume = padronizar_colunas_volume_raw(df_volume)
                
                add_log("Bases carregadas com sucesso.")
                
            if 'LMC name' not in df_frete.columns:
                st.error("🚨 **Arquivo de Frete Incorreto!** Você enviou um arquivo diferente no primeiro campo.")
                st.stop()
            if 'Cidade' not in df_abrangencia.columns or 'LMC Name' not in df_abrangencia.columns:
                st.error("🚨 **Arquivo de Abrangência Incorreto!** Você enviou um arquivo diferente no segundo campo.")
                st.stop()
            if 'Leve' not in df_volume.columns or 'Routing Code' not in df_volume.columns:
                st.error("🚨 **Colunas básicas ausentes no arquivo de Volume!**")
                st.stop()
            if 'Cidade' not in df_volume.columns or 'Faixa de peso cubado (g)' not in df_volume.columns:
                st.error("🚨 **Atenção: Base de Volume Desatualizada ou Incorreta!**")
                st.stop()
                
            with st.spinner("Finalizando processamento..."):
                df_volume['Cidade_Normalizada'] = df_volume['Cidade'].apply(normalize_string)
                df_abrangencia['Cidade_Normalizada'] = df_abrangencia['Cidade'].apply(normalize_string)
                
                st.session_state.df_frete_std = df_frete
                st.session_state.df_abrangencia_std = df_abrangencia
                st.session_state.df_volume_std = df_volume
                data_ready = True
    else:
        st.info("Por favor, faça o upload de **todas as 3 bases** na barra lateral para prosseguir.")
        
elif st.session_state.app_mode == "run":
    data_ready = True
    with st.sidebar:
        st.success("✅ Projeto carregado do Backup.")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⬅️ Sair e Voltar ao Início", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state.app_mode = "start"
            st.rerun()

if data_ready:
    df_price_var_clean = processar_price_var(df_price_var_raw)
    
    # 1. RAW DFs (Preservando as "gambiarras")
    df_frete_raw_clean = processar_frete(st.session_state.df_frete_std)
    df_abrangencia_raw = st.session_state.df_abrangencia_std.copy()
    if 'Região de preço 2023' in df_abrangencia_raw.columns:
        df_abrangencia_raw.rename(columns={'Região de preço 2023': 'Região de preço'}, inplace=True)
    
    df_volume_raw_tmp = st.session_state.df_volume_std.copy()
    df_volume_raw_tmp['Faixa de peso cubado (g)'] = df_volume_raw_tmp['Faixa de peso cubado (g)'].astype(str).str.strip()
    df_volume_grouped_raw = df_volume_raw_tmp.groupby(
        ['Leve', 'Cidade_Normalizada', 'Região de preço', 'Faixa de peso cubado (g)'], as_index=False
    ).agg({'# Total Packages': 'sum', 'Cidade': 'first'})
    
    # 2. CLEAN DFs (Aplicando a remoção do sufixo " B")
    df_frete_clean_std = st.session_state.df_frete_std.copy()
    if 'label' in df_frete_clean_std.columns: df_frete_clean_std['label'] = df_frete_clean_std['label'].apply(limpar_regiao)
    df_frete_clean = processar_frete(df_frete_clean_std)
    
    df_volume_clean_std = st.session_state.df_volume_std.copy()
    if 'Região de preço' in df_volume_clean_std.columns: df_volume_clean_std['Região de preço'] = df_volume_clean_std['Região de preço'].apply(limpar_regiao)
    df_volume_tmp_clean = df_volume_clean_std.copy()
    df_volume_tmp_clean['Faixa de peso cubado (g)'] = df_volume_tmp_clean['Faixa de peso cubado (g)'].astype(str).str.strip()
    df_volume_grouped_clean = df_volume_tmp_clean.groupby(
        ['Leve', 'Cidade_Normalizada', 'Região de preço', 'Faixa de peso cubado (g)'], as_index=False
    ).agg({'# Total Packages': 'sum', 'Cidade': 'first'})
    
    df_abrangencia_clean_std = st.session_state.df_abrangencia_std.copy()
    if 'Região de preço 2023' in df_abrangencia_clean_std.columns: df_abrangencia_clean_std['Região de preço 2023'] = df_abrangencia_clean_std['Região de preço 2023'].apply(limpar_regiao)

    df_slos_clean = processar_slos_novo(df_slos_raw)
    df_nomes_leves = processar_nomes_leves(st.session_state.df_volume_std)
    
    st.sidebar.success("Bases ativas no sistema!")
    
    with st.expander("2. Seleção de Leves", expanded=True):
        leves_disponiveis = df_frete_clean['LMC name'].dropna().unique().tolist()
        mapa_nomes = {}
        mapa_routing = {} 
        for lmc in leves_disponiveis:
            match = df_nomes_leves[df_nomes_leves['Leve'] == lmc]
            if not match.empty:
                nome_formatado = match['nome_completo'].values[0]
                mapa_nomes[nome_formatado] = lmc
                mapa_routing[lmc] = match['Routing Code'].values[0]
            else:
                mapa_nomes[lmc] = lmc
                mapa_routing[lmc] = "-"
                
        lista_nomes_exibicao = list(mapa_nomes.keys())
        default_leves = st.session_state.get('leves_selecionados_formatados_saved', [])
        default_leves = [l for l in default_leves if l in lista_nomes_exibicao]
        leves_selecionados_formatados = st.multiselect("Selecione os Leves envolvidos na negociação:", lista_nomes_exibicao, default=default_leves, key="leves_input")
        st.session_state['leves_selecionados_formatados_saved'] = leves_selecionados_formatados
        leves_selecionados = [mapa_nomes[nome] for nome in leves_selecionados_formatados]

    if leves_selecionados:
        with st.expander("3. Definição das Cidades Base", expanded=True):
            cidades_base_dict = {}
            cols = st.columns(len(leves_selecionados))
            for idx, leve in enumerate(leves_selecionados):
                nome_exibicao = leves_selecionados_formatados[idx]
                estado_do_leve = extrair_estado(leve)
                if estado_do_leve: df_cidades_estado = df_slos_clean[df_slos_clean['State'] == estado_do_leve]
                else: df_cidades_estado = df_slos_clean 
                
                opcoes_cidades = {}
                for _, row in df_cidades_estado.iterrows():
                    cid = str(row['Cidade'])
                    est = str(row['State'])
                    chave = str(row['Chave'])
                    opcoes_cidades[f"{cid} - {est}"] = chave
                    
                with cols[idx]:
                    op_list = sorted(list(opcoes_cidades.keys()))
                    saved_val = st.session_state.get(f'cidade_{leve}_saved', None)
                    def_idx = op_list.index(saved_val) if saved_val in op_list else 0
                    cidade_escolhida_display = st.selectbox(f"Cidade Base para:\n{nome_exibicao}", op_list, index=def_idx, key=f"cidade_{leve}_input")
                    st.session_state[f'cidade_{leve}_saved'] = cidade_escolhida_display
                    cidades_base_dict[leve] = opcoes_cidades[cidade_escolhida_display]

        with st.expander("4. Dados Atuais dos Leves Selecionados", expanded=False):
            st.markdown("### ℹ️ Contexto Atual das Bases Selecionadas")
            
            incluir_orfaos = st.checkbox(
                "Incluir pacotes entregues que não constam na abrangência da base", 
                value=False,
                help="Se marcado, a tabela abaixo mostrará todo o faturamento faturado pelo parceiro. Pacotes sem cidade registrada na base oficial de abrangência não serão considerados nos cenários simulados."
            )
            
            if incluir_orfaos:
                st.info("💡 **Atenção:** Esta opção afeta apenas a exibição desta Seção 4. As demais seções (Cenários e Relatórios) continuarão considerando estritamente o que está cadastrado na tabela de Abrangência Oficial do Leve para manter a consistência matemática.")
            
            gambiarras_gerais = set()

            for leve in leves_selecionados:
                vols_leve_raw = df_volume_grouped_raw[df_volume_grouped_raw['Leve'] == leve]
                
                # Rastreador de Órfãos
                cidades_oficiais_leve = df_abrangencia_raw[df_abrangencia_raw['LMC Name'] == leve]['Cidade_Normalizada'].unique()
                vols_orfaos = vols_leve_raw[~vols_leve_raw['Cidade_Normalizada'].isin(cidades_oficiais_leve)]
                
                if not vols_orfaos.empty:
                    qtd_orfaos = vols_orfaos['# Total Packages'].sum()
                    if incluir_orfaos:
                        st.warning(f"⚠️ **Aviso de Conciliação:** A base **{leve}** faturou **{int(qtd_orfaos)} pacotes** em cidades que não existem na sua base de abrangência oficial. Mostrando abaixo os dados com esses pacotes incluídos.")
                        with st.expander(f"Ver lista de cidades não mapeadas ({leve})"):
                            st.dataframe(vols_orfaos[['Cidade', 'Região de preço', '# Total Packages']].sort_values(by='# Total Packages', ascending=False), hide_index=True)
                    else:
                        st.warning(f"⚠️ **Aviso de Conciliação:** A base **{leve}** faturou **{int(qtd_orfaos)} pacotes** em cidades não mapeadas na abrangência. Eles foram excluídos da tabela abaixo. Marque a caixinha acima para incluí-los.")
                        # Remove os orfãos para mostrar a tabela matemática exata
                        vols_leve_raw = vols_leve_raw[vols_leve_raw['Cidade_Normalizada'].isin(cidades_oficiais_leve)]

                regioes_raw = vols_leve_raw['Região de preço'].dropna().unique()
                
                gambiarras_leve = [r for r in regioes_raw if str(r).endswith(" B") or str(r) in EQUIVALENCIAS_REGIOES]
                if gambiarras_leve:
                    st.warning(f"⚠️ **Atenção:** A base **{leve}** possui nomenclaturas desatualizadas ou fora do padrão: **{', '.join(gambiarras_leve)}**.")
                    gambiarras_gerais.update(gambiarras_leve)
                
                if len(regioes_raw) > 0:
                    context_leve_display = []
                    tot_vol_leve = 0
                    tot_fat_leve = 0
                    
                    for reg in regioes_raw:
                        vols_reg = vols_leve_raw[vols_leve_raw['Região de preço'] == reg]
                        vol_reg = vols_reg['# Total Packages'].sum()
                        fat_reg = 0
                        for _, r in vols_reg.iterrows():
                            fx = r['Faixa de peso cubado (g)']
                            qtd = r['# Total Packages']
                            tb = df_frete_raw_clean[(df_frete_raw_clean['LMC name'] == leve) & (df_frete_raw_clean['label'] == reg) & (df_frete_raw_clean['Faixa de peso cubado (g)'] == fx)]
                            p_on = tb['on time amount'].values[0] if not tb.empty else 0
                            fat_reg += qtd * p_on
                        
                        tb_fx1 = df_frete_raw_clean[(df_frete_raw_clean['LMC name'] == leve) & (df_frete_raw_clean['label'] == reg) & (df_frete_raw_clean['Faixa de peso cubado (g)'].astype(str).str.startswith('01'))]
                        v_fx1 = tb_fx1['on time amount'].values[0] if not tb_fx1.empty else 0
                        tk_reg = fat_reg / vol_reg if vol_reg > 0 else 0
                        
                        if vol_reg > 0:
                            context_leve_display.append({
                                "Região de Preço": reg,
                                "Volumetria (30d)": f"{int(vol_reg): ,}".replace(',', '.'),
                                "Faturamento Atual": formatar_moeda(fat_reg),
                                "Valor 1ª Faixa Atual": formatar_moeda(v_fx1),
                                "Ticket Médio Atual": formatar_moeda(tk_reg)
                            })
                            tot_vol_leve += vol_reg
                            tot_fat_leve += fat_reg
                    
                    if tot_vol_leve > 0:
                        context_leve_display.append({
                            "Região de Preço": "Total",
                            "Volumetria (30d)": f"{int(tot_vol_leve): ,}".replace(',', '.'),
                            "Faturamento Atual": formatar_moeda(tot_fat_leve),
                            "Valor 1ª Faixa Atual": "-",
                            "Ticket Médio Atual": formatar_moeda(tot_fat_leve / tot_vol_leve)
                        })
                        st.markdown(f"**Base (LMC):** {leve}")
                        st.dataframe(pd.DataFrame(context_leve_display), hide_index=True, use_container_width=True)
                else:
                    st.info(f"Nenhuma volumetria recente encontrada para a base: {leve}")
            
            st.session_state['gambiarras_encontradas'] = list(gambiarras_gerais)
            st.divider()
            
            tab1, tab2 = st.tabs(["Tabela Frete Peso Atual", "Abrangência e Prazos"])
            with tab1:
                for idx, leve in enumerate(leves_selecionados):
                    st.subheader(f"Tabela Frete Peso: {leves_selecionados_formatados[idx]}")
                    df_frete_leve = df_frete_raw_clean[df_frete_raw_clean['LMC name'] == leve].copy()
                    df_frete_leve['Routing Code'] = mapa_routing.get(leve, "-")
                    df_frete_leve.rename(columns={'label': 'Região de preço', 'on time amount': 'Valor do pacote dentro do prazo', 'out of time amount': 'Valor do pacote fora do prazo'}, inplace=True)
                    df_frete_leve['Valor do pacote dentro do prazo'] = df_frete_leve['Valor do pacote dentro do prazo'].apply(formatar_moeda)
                    df_frete_leve['Valor do pacote fora do prazo'] = df_frete_leve['Valor do pacote fora do prazo'].apply(formatar_moeda)
                    col_exib = ['LMC name', 'Routing Code', 'Região de preço', 'Faixa de peso cubado (g)', 'Valor do pacote dentro do prazo', 'Valor do pacote fora do prazo', 'table name']
                    st.dataframe(df_frete_leve[[c for c in col_exib if c in df_frete_leve.columns]], use_container_width=True, hide_index=True)
            with tab2:
                for idx, leve in enumerate(leves_selecionados):
                    st.subheader(f"Abrangência e Prazos: {leves_selecionados_formatados[idx]}")
                    df_abrangencia_leve = df_abrangencia_raw[df_abrangencia_raw['LMC Name'] == leve].copy()
                    df_abrangencia_leve['Routing Code'] = mapa_routing.get(leve, "-")
                    df_abrangencia_leve['SLO Local (Arquivo)'] = df_abrangencia_leve['Prazo adicional']
                    st.dataframe(df_abrangencia_leve[['LMC Name', 'Routing Code', 'Região de preço', 'Cidade', 'State', 'SLO Local (Arquivo)']], use_container_width=True, hide_index=True)
                    
        with st.expander("5. Definição do Leve/Lead de Destino", expanded=True):
            tipo_destino_options = ["Um Leve Existente (já selecionado)", "Um Novo Lead"]
            saved_td = st.session_state.get('tipo_destino_saved', "Um Leve Existente (já selecionado)")
            td_idx = tipo_destino_options.index(saved_td) if saved_td in tipo_destino_options else 0
            
            tipo_destino = st.radio("O destino da movimentação será para:", tipo_destino_options, index=td_idx, key="tipo_destino_input")
            st.session_state['tipo_destino_saved'] = tipo_destino
            
            nome_destino_final = ""
            cidade_base_destino = ""
            
            if tipo_destino == "Um Leve Existente (já selecionado)":
                saved_nd = st.session_state.get('nome_destino_display_saved', None)
                nd_idx = leves_selecionados_formatados.index(saved_nd) if saved_nd in leves_selecionados_formatados else 0
                
                nome_destino_display = st.selectbox("Selecione o Leve de Destino:", leves_selecionados_formatados, index=nd_idx, key="nome_destino_display_input")
                st.session_state['nome_destino_display_saved'] = nome_destino_display
                
                nome_destino_final = mapa_nomes.get(nome_destino_display)
                cidade_base_destino = cidades_base_dict.get(nome_destino_final)
                st.info(f"**Cidade Base do Destino (Chave):** {cidade_base_destino}")
            else:
                col_n1, col_n2, col_n3 = st.columns(3)
                with col_n1: 
                    nome_destino_final = st.text_input("Nome do Novo Lead:", value=st.session_state.get('nome_destino_final_saved', ''), placeholder="Ex: Lead - SP Sorocaba...", key="nome_destino_final_input")
                    st.session_state['nome_destino_final_saved'] = nome_destino_final
                with col_n2: 
                    estado_opts = sorted(df_slos_clean['State'].dropna().unique().tolist())
                    saved_est = st.session_state.get('estado_lead_saved', None)
                    est_idx = estado_opts.index(saved_est) if saved_est in estado_opts else 0
                    
                    estado_lead = st.selectbox("Estado do Novo Lead:", estado_opts, index=est_idx, key="estado_lead_input")
                    st.session_state['estado_lead_saved'] = estado_lead
                with col_n3: 
                    cidades_novo_lead = df_slos_clean[df_slos_clean['State'] == estado_lead]
                    opcoes_novos = {f"{r['Cidade']} - {r['State']}": r['Chave'] for _, r in cidades_novo_lead.iterrows()}
                    if opcoes_novos:
                        op_list = sorted(list(opcoes_novos.keys()))
                        saved_cid = st.session_state.get('cidade_base_destino_display_saved', None)
                        cid_idx = op_list.index(saved_cid) if saved_cid in op_list else 0
                        
                        cidade_base_destino_display = st.selectbox("Cidade Base do Novo Lead:", op_list, index=cid_idx, key="cidade_base_destino_display_input")
                        st.session_state['cidade_base_destino_display_saved'] = cidade_base_destino_display
                        
                        cidade_base_destino = opcoes_novos[cidade_base_destino_display]
                    else:
                        st.warning("Nenhuma cidade encontrada para este estado.")
                        cidade_base_destino = ""

        if nome_destino_final and cidade_base_destino:
            with st.expander("6. Manipulação de Abrangência", expanded=True):
                config_key = f"{','.join(leves_selecionados)}_{nome_destino_final}"
                
                if "mov_config_key" not in st.session_state or st.session_state.mov_config_key != config_key:
                    df_abrangencia_alvo = df_abrangencia_clean_std[df_abrangencia_clean_std['LMC Name'].isin(leves_selecionados)].copy()
                    df_mov = df_abrangencia_alvo[['LMC Name', 'Região de preço 2023', 'Cidade', 'State']].rename(columns={'Região de preço 2023': 'Região de preço'}).drop_duplicates(subset=['LMC Name', 'Cidade']).copy()
                    df_mov['Destino'] = "Manter no Leve Atual"
                    st.session_state.df_movimentacao = df_mov
                    st.session_state.mov_config_key = config_key

                opcoes_destino = ["Manter no Leve Atual", nome_destino_final]
                st.markdown("⚡ **Ações em Massa**")
                
                col_m1, col_m2, col_m3 = st.columns([1, 1, 1.5])
                with col_m1: bulk_lmc = st.selectbox("1. Filtrar Origem:", ["Selecione...", "Todos os Leves"] + leves_selecionados, key="bulk_lmc")
                with col_m2:
                    opcoes_reg = ["Todas as Regiões"]
                    if bulk_lmc != "Selecione...":
                        if bulk_lmc == "Todos os Leves": opcoes_reg += sorted(list(st.session_state.df_movimentacao['Região de preço'].unique()))
                        else: opcoes_reg += sorted(list(st.session_state.df_movimentacao[st.session_state.df_movimentacao['LMC Name'] == bulk_lmc]['Região de preço'].unique()))
                    bulk_reg = st.selectbox("2. Filtrar Região:", opcoes_reg, key="bulk_reg")
                
                cidades_disponiveis = []
                if bulk_lmc != "Selecione...":
                    mask_cid = pd.Series([True]*len(st.session_state.df_movimentacao), index=st.session_state.df_movimentacao.index)
                    if bulk_lmc != "Todos os Leves": mask_cid &= st.session_state.df_movimentacao['LMC Name'] == bulk_lmc
                    if bulk_reg != "Todas as Regiões": mask_cid &= st.session_state.df_movimentacao['Região de preço'] == bulk_reg
                    cidades_disponiveis = sorted(list(st.session_state.df_movimentacao.loc[mask_cid, 'Cidade'].unique()))
                
                with col_m3: bulk_cidades = st.multiselect("3. Filtrar Cidades (Opcional):", cidades_disponiveis, max_selections=200, placeholder="Selecione até 200 cidades...", key="bulk_cidades")

                col_m4, col_m5, col_m6 = st.columns([1.5, 1, 1])
                with col_m4: bulk_dest = st.selectbox("4. Escolher Destino:", opcoes_destino, key="bulk_dest")
                with col_m5:
                    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                    def aplicar_em_massa():
                        l, r, d, c = st.session_state.bulk_lmc, st.session_state.bulk_reg, st.session_state.bulk_dest, st.session_state.bulk_cidades
                        if l != "Selecione...":
                            if l == "Todos os Leves": mask = pd.Series([True]*len(st.session_state.df_movimentacao), index=st.session_state.df_movimentacao.index)
                            else: mask = st.session_state.df_movimentacao['LMC Name'] == l
                            if r != "Todas as Regiões": mask &= st.session_state.df_movimentacao['Região de preço'] == r
                            if c: mask &= st.session_state.df_movimentacao['Cidade'].isin(c)
                            st.session_state.df_movimentacao.loc[mask, 'Destino'] = d
                    st.button("Aplicar Ação", on_click=aplicar_em_massa, use_container_width=True)
                with col_m6:
                    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                    st.button("🔄 Resetar", on_click=lambda: st.session_state.df_movimentacao.assign(Destino="Manter no Leve Atual"), use_container_width=True)

                st.session_state.df_movimentacao = st.data_editor(
                    st.session_state.df_movimentacao,
                    column_config={"Destino": st.column_config.SelectboxColumn("Destino", options=opcoes_destino, required=True)},
                    disabled=["LMC Name", "Região de preço", "Cidade", "State"], hide_index=True, use_container_width=True, height=400
                )
                df_movidos = st.session_state.df_movimentacao[st.session_state.df_movimentacao['Destino'] == nome_destino_final].copy()

            # --- PROCESSAMENTO BASE NEUTRA E ESTRATÉGIA ---
            df_abrangencia_existente = pd.DataFrame(columns=['LMC Name', 'Região de preço', 'Cidade', 'State', 'Observação'])
            if tipo_destino == "Um Leve Existente (já selecionado)":
                df_abrangencia_existente = df_abrangencia_clean_std[df_abrangencia_clean_std['LMC Name'] == nome_destino_final].copy()
                df_abrangencia_existente = df_abrangencia_existente.rename(columns={'Região de preço 2023': 'Região de preço'})
                df_abrangencia_existente['Observação'] = f"Abrangência atual de {nome_destino_final}"
            
            df_movidos_fmt = df_movidos.copy()
            df_movidos_fmt['Observação'] = "Migrado de: " + df_movidos_fmt['LMC Name']
            df_escopo_final = pd.concat([df_abrangencia_existente[['Cidade', 'State', 'Região de preço', 'Observação']], df_movidos_fmt[['Cidade', 'State', 'Região de preço', 'Observação']]], ignore_index=True).drop_duplicates(subset=['Cidade'])
            
            slo_base_dest = df_slos_clean[df_slos_clean['Chave'] == cidade_base_destino]['SLO'].values
            slo_base_dest_val = slo_base_dest[0] if len(slo_base_dest) > 0 else 0
            
            df_escopo_final['Chave_Alvo'] = df_escopo_final.apply(lambda r: normalize_string(r['Cidade']) + (str(r['State']).strip()[:2].lower() if pd.notna(r['State']) else ''), axis=1)
            df_escopo_final = df_escopo_final.merge(df_slos_clean[['Chave', 'SLO']], left_on='Chave_Alvo', right_on='Chave', how='left')
            df_escopo_final['Novo SLO Local'] = df_escopo_final['SLO'] - slo_base_dest_val
            df_escopo_final['Novo SLO Local'] = df_escopo_final['Novo SLO Local'].apply(lambda x: x if pd.notnull(x) and x > 0 else 0).astype(int)
            colunas_finais_abrangencia = ['Cidade', 'State', 'Região de preço', 'Novo SLO Local', 'Observação']
            
            regioes_finais_destino = sorted(df_escopo_final['Região de preço'].unique().tolist())
            
            # --- NOVA LÓGICA DE ESCOPO DE VOLUME (Apenas Municípios Movimentados) ---
            df_abrangencia_existente_pairs = df_abrangencia_existente[['LMC Name', 'Cidade']].rename(columns={'LMC Name': 'Leve'})
            df_movidos_pairs = df_movidos[['LMC Name', 'Cidade']].rename(columns={'LMC Name': 'Leve'})
            
            df_valid_scope = pd.concat([df_abrangencia_existente_pairs, df_movidos_pairs], ignore_index=True)
            df_valid_scope['Cidade_Normalizada'] = df_valid_scope['Cidade'].apply(normalize_string)
            df_valid_scope = df_valid_scope[['Leve', 'Cidade_Normalizada']].drop_duplicates()
            
            # Cruza a volumetria bruta para conter APENAS os municípios exatos do escopo destino
            df_volume_alvo_raw = df_volume_grouped_raw.merge(df_valid_scope, on=['Leve', 'Cidade_Normalizada'], how='inner')

            df_tabelas_base_list = []
            dict_base_on = {}
            mult_dict_cod = dict(zip(df_price_var_clean['Cod'], df_price_var_clean['Multiplicador']))
            
            for clean_reg in regioes_finais_destino:
                raw_regs = [r for r in df_volume_alvo_raw['Região de preço'].unique() if limpar_regiao(r) == clean_reg]
                
                c_on = 0; s_mult = 0
                for raw_reg in raw_regs:
                    vols_raw_reg = df_volume_alvo_raw[df_volume_alvo_raw['Região de preço'] == raw_reg]
                    for _, r in vols_raw_reg.iterrows():
                        lmc = r['Leve']
                        fx = r['Faixa de peso cubado (g)']
                        qtd = r['# Total Packages']
                        fx_cod = str(fx).strip()[:2]
                        
                        tb = df_frete_raw_clean[(df_frete_raw_clean['LMC name'] == lmc) & (df_frete_raw_clean['label'] == raw_reg) & (df_frete_raw_clean['Faixa de peso cubado (g)'] == fx)]
                        p_on = tb['on time amount'].values[0] if not tb.empty else 0
                        
                        c_on += qtd * p_on
                        s_mult += qtd * mult_dict_cod.get(fx_cod, 1.0)
                        
                if s_mult > 0:
                    base_on = c_on / s_mult
                else:
                    lmc_fallback_list = df_movidos[df_movidos['Região de preço'] == clean_reg]['LMC Name'].tolist()
                    if not lmc_fallback_list and tipo_destino == "Um Leve Existente (já selecionado)":
                        lmc_fallback_list = df_abrangencia_existente[df_abrangencia_existente['Região de preço'] == clean_reg]['LMC Name'].tolist()
                    
                    base_on = 0
                    if lmc_fallback_list:
                        lmc_fallback = lmc_fallback_list[0]
                        tb_fallback = df_frete_clean[(df_frete_clean['LMC name'] == lmc_fallback) & (df_frete_clean['label'] == clean_reg)]
                        fx1 = tb_fallback[tb_fallback['Faixa de peso cubado (g)'].astype(str).str.startswith('01')]
                        if not fx1.empty:
                            base_on = fx1['on time amount'].values[0] / 0.83

                dict_base_on[clean_reg] = base_on
                
                faixas_unicas = sorted(df_frete_clean['Faixa de peso cubado (g)'].dropna().unique())
                df_regiao_base = pd.DataFrame({'Faixa de peso cubado (g)': faixas_unicas})
                df_regiao_base['Cod'] = df_regiao_base['Faixa de peso cubado (g)'].astype(str).str.strip().str[:2]
                df_regiao_base['Multiplicador'] = df_regiao_base['Cod'].map(mult_dict_cod).fillna(1.0)
                
                df_regiao_base['Região de Preço'] = clean_reg
                df_regiao_base['Valor dentro do prazo'] = df_regiao_base['Multiplicador'] * base_on
                df_regiao_base['Valor fora do prazo'] = df_regiao_base['Multiplicador'] * base_on * 0.65
                df_tabelas_base_list.append(df_regiao_base[['Região de Preço', 'Faixa de peso cubado (g)', 'Valor dentro do prazo', 'Valor fora do prazo']])

            df_tabela_base_completa = pd.DataFrame()
            if df_tabelas_base_list:
                df_tabela_base_completa = pd.concat(df_tabelas_base_list, ignore_index=True)

            with st.expander("7. Tabela Base Calculada", expanded=True):
                pode_prosseguir = False
                estrategia_preco = "Tabela Equivalente (Média Ponderada)"
                
                if st.session_state.get('gambiarras_encontradas'):
                    gambiarras = list(st.session_state['gambiarras_encontradas'])
                    st.info("ℹ️ **Nota de Padronização:** O sistema identificou nomenclaturas antigas na aba anterior. Para garantir simulações precisas, os volumes e valores reais foram absorvidos e as regiões foram remapeadas conforme abaixo:")
                    for g in gambiarras:
                        st.markdown(f"- **{g}** ➔ Atualizada para **{limpar_regiao(g)}**")
                
                if df_movidos.empty and tipo_destino == "Um Novo Lead":
                    st.warning("Nenhum município foi movimentado para o Novo Lead ainda.")
                elif df_movidos.empty and tipo_destino == "Um Leve Existente (já selecionado)":
                    st.info("💡 **Modo de Reajuste Comercial:** Nenhuma cidade foi movimentada. O simulador manterá a abrangência e tabela atuais do Leve como base neutra para simulações de ajustes no Passo 8.")
                    pode_prosseguir = True
                else:
                    st.info("💡 **Geração Automática de Tabela Base:** O sistema calculou uma tabela equivalente neutra (média ponderada) consolidando as volumetrias das origens envolvidas. Você poderá aplicar reajustes no próximo passo.")
                    st.markdown("##### 📊 Tabela Base Equivalente (Neutra - 0% de Ajuste)")
                    df_exibicao_base = df_tabela_base_completa.copy()
                    if not df_exibicao_base.empty:
                        df_exibicao_base.rename(columns={'Faixa de peso (g/m³)': 'Faixa de peso cubado (g)'}, inplace=True)
                        df_exibicao_base['Valor dentro do prazo'] = df_exibicao_base['Valor dentro do prazo'].apply(formatar_moeda)
                        df_exibicao_base['Valor fora do prazo'] = df_exibicao_base['Valor fora do prazo'].apply(formatar_moeda)
                        st.dataframe(df_exibicao_base, hide_index=True, use_container_width=True)
                    pode_prosseguir = True

            keys_to_save = {}
            for k, v in st.session_state.items():
                if isinstance(v, (int, float, str, bool, list)):
                    if not str(k).startswith(("btn_", "dl_")) and k != "FormSubmitter":
                        keys_to_save[k] = v
            state_json = json.dumps(keys_to_save)
            
            zip_data = get_backup_zip_cached(
                state_json, 
                st.session_state.df_frete_std.to_csv(index=False), 
                st.session_state.df_abrangencia_std.to_csv(index=False), 
                st.session_state.df_volume_std.to_csv(index=False),
                st.session_state.df_movimentacao.to_csv(index=False) if 'df_movimentacao' in st.session_state else None
            )
            st.sidebar.download_button(
                label="💾 Salvar Estado da Análise (.zip)",
                data=zip_data,
                file_name=f"Backup_Analise_Loggi_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
                key="dl_backup_sidebar"
            )
            
            # --- PROCESSAMENTO DOS CENÁRIOS E RESULTADOS ---
            if pode_prosseguir:
                dict_tabelas_finais = {}
                tabelas_atuais_pdf = {}
                cenario_metrics = {}
                
                context_raw = []
                tot_vol_context = 0
                tot_fat_context = 0
                
                for clean_reg in regioes_finais_destino:
                    raw_regs = [r for r in df_volume_alvo_raw['Região de preço'].unique() if limpar_regiao(r) == clean_reg]
                    vol_total = 0
                    c_on_total = 0
                    
                    for raw_reg in raw_regs:
                        vols_raw_reg = df_volume_alvo_raw[df_volume_alvo_raw['Região de preço'] == raw_reg]
                        for _, r in vols_raw_reg.iterrows():
                            lmc = r['Leve']
                            fx = r['Faixa de peso cubado (g)']
                            qtd = r['# Total Packages']
                            tb = df_frete_raw_clean[(df_frete_raw_clean['LMC name'] == lmc) & (df_frete_raw_clean['label'] == raw_reg) & (df_frete_raw_clean['Faixa de peso cubado (g)'] == fx)]
                            p_on = tb['on time amount'].values[0] if not tb.empty else 0
                            
                            vol_total += qtd
                            c_on_total += qtd * p_on
                            
                    context_raw.append({
                        "Região de Preço": clean_reg,
                        "Volumetria (30d)": int(vol_total),
                        "Faturamento Atual": c_on_total,
                        "Valor 1ª Faixa (Base)": dict_base_on.get(clean_reg, 0) * 0.83
                    })
                    tot_vol_context += vol_total
                    tot_fat_context += c_on_total

                tot_tk_context = tot_fat_context / tot_vol_context if tot_vol_context > 0 else 0
                
                context_display = []
                for item in context_raw:
                    context_display.append({
                        "Região de Preço": item["Região de Preço"],
                        "Volumetria (30d)": f"{item['Volumetria (30d)']: ,}".replace(',', '.'),
                        "Faturamento Atual": formatar_moeda(item['Faturamento Atual']),
                        "Valor 1ª Faixa (Base)": formatar_moeda(item['Valor 1ª Faixa (Base)']),
                        "Ticket Médio Atual": formatar_moeda(item['Faturamento Atual'] / item['Volumetria (30d)'] if item['Volumetria (30d)'] > 0 else 0)
                    })
                    
                context_display.append({
                    "Região de Preço": "Total Geral",
                    "Volumetria (30d)": f"{int(tot_vol_context): ,}".replace(',', '.'),
                    "Faturamento Atual": formatar_moeda(tot_fat_context),
                    "Valor 1ª Faixa (Base)": "-",
                    "Ticket Médio Atual": formatar_moeda(tot_tk_context)
                })

                with st.expander("8. Ajustes Comerciais e Cenários", expanded=True):
                    st.markdown("### ℹ️ Contexto Atual das Regiões Envolvidas")
                    
                    if st.session_state.get('gambiarras_encontradas'):
                        st.info("ℹ️ As regiões de preço fora do padrão foram consolidadas no faturamento exibido abaixo.")
                        
                    if context_display:
                        st.dataframe(pd.DataFrame(context_display), hide_index=True, use_container_width=True)
                    else:
                        st.info("Nenhuma volumetria recente encontrada para as cidades selecionadas.")
                        
                    c_head, c_btn_add, c_btn_del = st.columns([5, 1.5, 1.5])
                    with c_head:
                        st.markdown("### 🎛️ Configuração de Cenários")
                    with c_btn_add:
                        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
                        if st.button("➕ Novo Cenário", use_container_width=True):
                            st.session_state["num_cenarios"] += 1
                            st.rerun()
                    with c_btn_del:
                        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
                        if st.button("➖ Remover Cenário", use_container_width=True):
                            if st.session_state["num_cenarios"] > 1:
                                st.session_state["num_cenarios"] -= 1
                                st.rerun()
                            else:
                                st.toast("⚠️ É preciso ter ao menos 1 cenário.")

                    cenarios_nomes = [f"Cenário {i+1}" for i in range(st.session_state["num_cenarios"])]
                    
                    css_tabs = "<style>"
                    for i in range(st.session_state["num_cenarios"]):
                        bg_cor = CORES_CENARIOS[i % len(CORES_CENARIOS)]
                        css_tabs += f'button[data-baseweb="tab"]:nth-child({i+1}) {{ background-color: {bg_cor} !important; border-radius: 6px 6px 0 0; margin-right: 2px; border: 1px solid #ccc; border-bottom: none; }}\n'
                    css_tabs += "</style>"
                    st.markdown(css_tabs, unsafe_allow_html=True)

                    tabs_cenarios = st.tabs(cenarios_nomes)
                    
                    if "bkp_cenarios" not in st.session_state:
                        st.session_state.bkp_cenarios = {}
                        
                    def backup_input(key):
                        st.session_state.bkp_cenarios[key] = st.session_state[key]

                    for c_idx, tab in enumerate(tabs_cenarios):
                        cen_id = f"c{c_idx+1}"
                        with tab:
                            st.markdown(f"**Ajuste em Massa - {cenarios_nomes[c_idx]}**")
                            cg1, cg2, cg3, cg4 = st.columns([1.5, 1.5, 2, 3])
                            with cg1: st.selectbox("Tipo de Ajuste", ["%", "R$ (1ª Faixa)", "R$ (Ticket Médio)"], key=f"global_tipo_{cen_id}")
                            with cg2: st.number_input("Valor", step=0.5, format="%.2f", key=f"global_val_{cen_id}")
                            with cg3:
                                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                                def aplicar_global(c_id=cen_id):
                                    t = st.session_state[f"global_tipo_{c_id}"]
                                    v = st.session_state[f"global_val_{c_id}"]
                                    for r in regioes_finais_destino:
