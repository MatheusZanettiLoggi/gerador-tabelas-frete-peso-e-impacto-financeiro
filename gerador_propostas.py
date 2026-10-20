st.subheader(f"🤝 Visão do Parceiro ({cen})")
                            cp1, cp2, cp3 = st.columns(3)
                            with cp1:
                                st.metric("Faturamento Atual (Sem Novas Cidades)", formatar_moeda(m['fat_parceiro_antigo']))
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria: {int(m['vol_parceiro_antigo']):,} pacotes</span>", unsafe_allow_html=True)
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio: {formatar_moeda(m['tk_parceiro_antigo'])}</span>", unsafe_allow_html=True)
                            with cp2:
                                st.metric("Novo Faturamento Projetado", formatar_moeda(m['fat_novo']))
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria: {int(m['vol_fat_novo']):,} pacotes</span>", unsafe_allow_html=True)
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio: {formatar_moeda(m['tk_fat_novo'])}</span>", unsafe_allow_html=True)
                            with cp3:
                                val_cresc = m['cresc_fat_parceiro']
                                str_cresc = f"+{formatar_moeda(val_cresc)}" if val_cresc > 0 else (f"-{formatar_moeda(abs(val_cresc))}" if val_cresc < 0 else "R$ 0,00")
                                
                                st.markdown(f"<div style='font-size: 14px; color: gray;'>Crescimento da Operação</div>", unsafe_allow_html=True)
                                st.markdown(f"<div style='font-size: 1.8rem; font-weight: normal; margin-bottom: 5px;'>{str_cresc}</div>", unsafe_allow_html=True)
                                
                                if m['perc_cresc_parceiro'] > 0:
                                    st.markdown(f"<span style='font-size: 0.9em; color: #09ab3b; font-weight: bold;'>▲ +{m['perc_cresc_parceiro']:.2f}% de aumento no faturamento</span>", unsafe_allow_html=True)
                                elif m['perc_cresc_parceiro'] < 0:
                                    st.markdown(f"<span style='font-size: 0.9em; color: #ff4b4b; font-weight: bold;'>▼ {m['perc_cresc_parceiro']:.2f}% de perda de faturamento</span>", unsafe_allow_html=True)
                                else:
                                    st.markdown(f"<span style='font-size: 0.9em; color: gray; font-weight: bold;'>■ 0.00% (Sem alteração)</span>", unsafe_allow_html=True)
                            
                            if m['municipios_migrados_nomes']:
                                mun_str = ", ".join(m['municipios_migrados_nomes'])
                                vol_mig = int(m['vol_migrado'])
                                fat_mig = formatar_moeda(m['fat_migrado_projetado'])
                                st.info(f"💡 **Cidades Adicionadas:** Estamos adicionando os municípios de **{mun_str}**, que representam um volume de **{vol_mig:,} pacotes** a mais e um faturamento projetado adicional de **{fat_mig}**, compondo o Novo Faturamento Projetado.")

                            st.divider()
                            st.subheader(f"🐰 Visão Loggi ({cen})")
                            cl1, cl2, cl3 = st.columns(3)
                            with cl1:
                                st.metric("Custo Antigo Global", formatar_moeda(m['loggi_antigo']))
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria Total: {int(m['vol_loggi']):,} pacotes</span>", unsafe_allow_html=True)
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio Antigo: {formatar_moeda(m['tk_loggi_antigo'])}</span>", unsafe_allow_html=True)
                            with cl2:
                                st.metric("Novo Custo Global Projetado", formatar_moeda(m['loggi_novo']))
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Volumetria Total: {int(m['vol_loggi']):,} pacotes</span>", unsafe_allow_html=True)
                                st.markdown(f"<span style='font-size: 0.9em; color: gray;'>Ticket Médio Novo: {formatar_moeda(m['tk_loggi_novo'])}</span>", unsafe_allow_html=True)
                            with cl3:
                                val_imp = m['imp_loggi']
                                str_imp = f"+{formatar_moeda(val_imp)}" if val_imp > 0 else (f"-{formatar_moeda(abs(val_imp))}" if val_imp < 0 else "R$ 0,00")
                                
                                st.markdown(f"<div style='font-size: 14px; color: gray;'>Impacto Financeiro Loggi</div>", unsafe_allow_html=True)
                                st.markdown(f"<div style='font-size: 1.8rem; font-weight: normal; margin-bottom: 5px;'>{str_imp}</div>", unsafe_allow_html=True)
                                
                                if val_imp > 0:
                                    st.markdown(f"<div style='color: #ff4b4b; background-color: #ffcccc; display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.9em; font-weight: bold; margin-bottom: 10px;'>▲ +{formatar_moeda(val_imp)} (Aumento de Custo)</div>", unsafe_allow_html=True)
                                    st.markdown(f"**% Aumento:** <span style='color:#ff4b4b'>▲ +{m['perc_imp_loggi']:.2f}% de impacto no budget</span>", unsafe_allow_html=True)
                                elif val_imp < 0:
                                    st.markdown(f"<div style='color: #09ab3b; background-color: #ccffcc; display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.9em; font-weight: bold; margin-bottom: 10px;'>▼ -{formatar_moeda(abs(val_imp))} (Economia)</div>", unsafe_allow_html=True)
                                    st.markdown(f"**% Aumento:** <span style='color:#09ab3b'>▼ {m['perc_imp_loggi']:.2f}% de economia no budget</span>", unsafe_allow_html=True)
                                else:
                                    st.markdown(f"<div style='color: gray; background-color: #f0f0f0; display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.9em; font-weight: bold; margin-bottom: 10px;'>■ R$ 0,00 (Neutro)</div>", unsafe_allow_html=True)
                                    st.markdown(f"**% Aumento:** <span style='color:gray'>■ 0.00%</span>", unsafe_allow_html=True)
