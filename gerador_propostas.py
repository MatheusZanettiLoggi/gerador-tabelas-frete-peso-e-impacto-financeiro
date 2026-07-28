def generate_html_pdf(nome_destino, estrategia, cidades_movimentadas_str,
                      fat_antigo, vol_fat_antigo, tk_fat_antigo,
                      fat_novo, vol_fat_novo, tk_fat_novo, cresc_fat, perc_cresc,
                      loggi_antigo, vol_loggi, tk_loggi_antigo,
                      loggi_novo, tk_loggi_novo, imp_loggi, perc_imp_loggi,
                      detalhes_reg, df_abrangencia_out, df_tabela_out, tabelas_atuais_pdf):

    # Formatting functions
    def format_money(val):
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    def format_perc(val):
        return f"{val:+.2f}%"

    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Relatório de Simulação</title>
        <style>
            @page {{
                size: A4;
                margin: 20mm 15mm;
                background-color: #fdfbf7;
                @bottom-center {{
                    content: "Simulador de Movimentação de Leves - Desenvolvido por Matheus Zanetti | Página " counter(page);
                    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                    font-size: 8pt;
                    color: #888888;
                    font-style: italic;
                }}
            }}
            body {{
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                margin: 0;
                padding: 0;
                color: #333333;
                background-color: #fdfbf7;
                font-size: 10pt;
                line-height: 1.5;
            }}
            *, *::before, *::after {{
                box-sizing: border-box;
            }}
            
            /* Typography & Colors */
            h1 {{
                color: #002766; /* Loggi Dark Blue */
                font-size: 18pt;
                text-align: center;
                margin-top: 0;
                margin-bottom: 5px;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
            h2 {{
                color: #006aff; /* Loggi Blue */
                font-size: 13pt;
                border-bottom: 2px solid #00baff; /* Loggi Light Blue */
                padding-bottom: 4px;
                margin-top: 25px;
                margin-bottom: 15px;
            }}
            h3 {{
                color: #002766;
                font-size: 11pt;
                margin-top: 15px;
                margin-bottom: 8px;
            }}
            
            /* Header Section */
            .header-meta {{
                text-align: center;
                color: #666;
                font-size: 9pt;
                margin-bottom: 25px;
            }}
            .context-box {{
                background-color: #ffffff;
                border-left: 4px solid #00baff;
                padding: 12px 15px;
                margin-bottom: 25px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            }}
            .context-item {{
                margin-bottom: 6px;
            }}
            .label {{
                font-weight: bold;
                color: #002766;
            }}

            /* Data Cards */
            .card-container {{
                display: block;
                width: 100%;
                margin-bottom: 20px;
                page-break-inside: avoid;
            }}
            .card {{
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 15px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.04);
            }}
            .card-title {{
                font-weight: bold;
                font-size: 11pt;
                color: #002766;
                margin-bottom: 12px;
                border-bottom: 1px solid #f0f0f0;
                padding-bottom: 8px;
            }}
            .metric-row {{
                display: block;
                width: 100%;
                margin-bottom: 8px;
            }}
            .metric-label {{
                display: inline-block;
                width: 180px;
                color: #555;
            }}
            .metric-value {{
                display: inline-block;
                font-weight: bold;
                color: #222;
            }}
            .metric-sub {{
                color: #888;
                font-size: 8.5pt;
                margin-left: 10px;
                font-weight: normal;
            }}
            
            /* Arrow indicators */
            .arrow-up-green {{ color: #09ab3b; font-weight: bold; }}
            .arrow-down-green {{ color: #09ab3b; font-weight: bold; }}
            .arrow-up-red {{ color: #ff4b4b; font-weight: bold; }}
            .arrow-down-red {{ color: #ff4b4b; font-weight: bold; }}
            .arrow-neutral {{ color: #888888; font-weight: bold; }}

            /* Region Details */
            .region-block {{
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                padding: 12px;
                margin-bottom: 10px;
                page-break-inside: avoid;
            }}
            .region-title {{
                font-weight: bold;
                color: #006aff;
                margin-bottom: 8px;
                font-size: 10.5pt;
            }}
            .region-alert {{
                color: #e67e22;
                font-weight: bold;
                font-size: 9pt;
                margin-bottom: 8px;
            }}
            
            /* Tables */
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                background-color: #ffffff;
                font-size: 8.5pt;
            }}
            th {{
                background-color: #002766;
                color: #ffffff;
                font-weight: bold;
                text-align: center;
                padding: 8px 5px;
                border: 1px solid #002766;
            }}
            td {{
                padding: 6px 5px;
                border: 1px solid #e0e0e0;
                text-align: center;
            }}
            tr:nth-child(even) {{
                background-color: #f9fbfd;
            }}
            
            .page-break {{
                page-break-before: always;
            }}
        </style>
    </head>
    <body>
    """

    # --- Header ---
    fuso_brasilia = datetime.now(timezone(timedelta(hours=-3)))
    data_extracao = fuso_brasilia.strftime("%d/%m/%Y às %H:%M")
    
    html_content += f"""
        <h1>Relatório do Simulador de Movimentação de Leves</h1>
        <div class="header-meta">
            Gerado em: {data_extracao}<br>
            Base de Volumetria: Análise dos últimos 30 dias
        </div>
        
        <div class="context-box">
            <div class="context-item"><span class="label">Destino / Lead:</span> {nome_destino}</div>
            <div class="context-item"><span class="label">Estratégia Escolhida:</span> {estrategia}</div>
            <div class="context-item"><span class="label">Municípios Absorvidos:</span> {cidades_movimentadas_str}</div>
        </div>
    """

    # --- Helper for indicators ---
    def get_indicator(val, is_cost=False):
        if val > 0:
            return f'<span class="{"arrow-up-red" if is_cost else "arrow-up-green"}">▲ +{format_money(val)}</span>'
        elif val < 0:
            return f'<span class="{"arrow-down-green" if is_cost else "arrow-down-red"}">▼ {format_money(val)}</span>'
        return f'<span class="arrow-neutral">■ {format_money(val)}</span>'

    def get_perc_indicator(val, is_cost=False):
        if val > 0:
            return f'<span class="{"arrow-up-red" if is_cost else "arrow-up-green"}">▲ +{val:.2f}%</span>'
        elif val < 0:
            return f'<span class="{"arrow-down-green" if is_cost else "arrow-down-red"}">▼ {val:.2f}%</span>'
        return f'<span class="arrow-neutral">■ 0.00%</span>'


    # --- 1. Parceiro ---
    html_content += f"""
        <h2>1. VISÃO DO PARCEIRO (Faturamento do Leve)</h2>
        <div class="card-container">
            <div class="card">
                <div class="metric-row">
                    <span class="metric-label">Faturamento Atual:</span>
                    <span class="metric-value">{format_money(fat_antigo)}</span>
                    <span class="metric-sub">(Vol: {int(vol_fat_antigo):,} | TK: {format_money(tk_fat_antigo)})</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Novo Faturamento:</span>
                    <span class="metric-value">{format_money(fat_novo)}</span>
                    <span class="metric-sub">(Vol: {int(vol_fat_novo):,} | TK: {format_money(tk_fat_novo)})</span>
                </div>
                <div class="metric-row" style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #eee;">
                    <span class="metric-label">Crescimento:</span>
                    <span class="metric-value">{get_indicator(cresc_fat, False)} {get_perc_indicator(perc_cresc, False)}</span>
                </div>
            </div>
        </div>
    """

    # --- 2. Loggi ---
    html_content += f"""
        <h2>2. VISÃO LOGGI (Impacto Financeiro Real)</h2>
        <div class="card-container">
            <div class="card">
                <div class="metric-row">
                    <span class="metric-label">Custo Antigo Global:</span>
                    <span class="metric-value">{format_money(loggi_antigo)}</span>
                    <span class="metric-sub">(Vol: {int(vol_loggi):,} | TK: {format_money(tk_loggi_antigo)})</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Novo Custo Projetado:</span>
                    <span class="metric-value">{format_money(loggi_novo)}</span>
                    <span class="metric-sub">(Vol: {int(vol_loggi):,} | TK: {format_money(tk_loggi_novo)})</span>
                </div>
                <div class="metric-row" style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #eee;">
                    <span class="metric-label">Impacto Financeiro:</span>
                    <span class="metric-value">{get_indicator(imp_loggi, True)} {get_perc_indicator(perc_imp_loggi, True)}</span>
                </div>
            </div>
        </div>
    """

    # --- 3. Detalhamento ---
    html_content += """<h2>3. DETALHAMENTO POR REGIÃO</h2>"""
    
    if detalhes_reg:
        for reg, dados in detalhes_reg.items():
            tk_ant = dados['custo_antigo'] / dados['vol'] if dados['vol'] > 0 else 0
            tk_nov = dados['custo_novo'] / dados['vol'] if dados['vol'] > 0 else 0
            imp_r = dados['custo_novo'] - dados['custo_antigo']
            perc_r = (imp_r / dados['custo_antigo']) * 100 if dados['custo_antigo'] > 0 else 0
            
            ajuste = dados.get('ajuste', 0.0)
            ajuste_html = f'<div class="region-alert">⚠️ Ajuste Comercial Aplicado: {ajuste:+.2f}%</div>' if ajuste != 0.0 else ''
            
            html_content += f"""
            <div class="region-block">
                <div class="region-title">📍 Região: {reg}</div>
                {ajuste_html}
                <div class="metric-row">
                    <span class="metric-label">Custo Antigo:</span>
                    <span class="metric-value">{format_money(dados['custo_antigo'])}</span>
                    <span class="metric-sub">(TK: {format_money(tk_ant)})</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Novo Custo:</span>
                    <span class="metric-value">{format_money(dados['custo_novo'])}</span>
                    <span class="metric-sub">(TK: {format_money(tk_nov)})</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Variação:</span>
                    <span class="metric-value">{get_indicator(imp_r, True)} {get_perc_indicator(perc_r, True)}</span>
                </div>
            </div>
            """
    else:
        html_content += "<p>Nenhum detalhamento de região disponível.</p>"


    # --- Tables HTML Generator ---
    def generate_html_table(df):
        if df.empty: return "<p>Sem dados.</p>"
        html = "<table><thead><tr>"
        for col in df.columns:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>"
        for _, row in df.iterrows():
            html += "<tr>"
            for val in row:
                html += f"<td>{val}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        return html

    # --- 4. Abrangencia ---
    html_content += f"""
        <div class="page-break"></div>
        <h2>4. ABRANGÊNCIA COMPLETA PROJETADA: {nome_destino}</h2>
    """
    colunas_abr_pdf = [c for c in df_abrangencia_out.columns if c != 'State']
    html_content += generate_html_table(df_abrangencia_out[colunas_abr_pdf])

    # --- 5. Tabela Frete ---
    html_content += f"""
        <div class="page-break"></div>
        <h2>5. TABELA FRETE PESO PROJETADA: {nome_destino}</h2>
    """
    html_content += generate_html_table(df_tabela_out)

    # --- 6. Tabelas Atuais ---
    if tabelas_atuais_pdf:
        html_content += f"""
            <div class="page-break"></div>
            <h2>6. TABELAS FRETE PESO ATUAIS (ORIGENS ENVOLVIDAS)</h2>
        """
        for leve_nome, df_tab in tabelas_atuais_pdf.items():
            html_content += f"<h3>Leve Atual: {leve_nome}</h3>"
            html_content += generate_html_table(df_tab)

    html_content += """
    </body>
    </html>
    """
    
    with open("report.html", "w", encoding="utf-8") as f:
        f.write(html_content)
        
    return html_content

# We only need to write the function definition out so we can inspect it or execute it if needed.
print("Function defined.")