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
    df_nomes_leves = processar_nomes_leves
