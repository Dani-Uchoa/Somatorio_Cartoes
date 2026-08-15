import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

COLUNAS_PADRAO = ['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']


def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def extrair_info_empresa(arquivo_getnet):
    """
    Lê o cabeçalho comercial da planilha da Getnet (Razão Social e Período)
    varrendo as primeiras linhas de qualquer uma das abas disponíveis.
    """
    empresa = None
    periodo = None
    try:
        xls = pd.ExcelFile(arquivo_getnet)
        for sheet in xls.sheet_names:
            raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=6)
            for r in range(raw.shape[0]):
                for c in range(raw.shape[1]):
                    val = raw.iat[r, c]
                    if not isinstance(val, str):
                        continue
                    if empresa is None and val.strip().startswith('Razão Social:'):
                        empresa = val.split(':', 1)[1].strip()
                    if periodo is None and val.strip().startswith('Periodo:'):
                        periodo = val.split(':', 1)[1].strip()
            if empresa and periodo:
                break
    except Exception:
        pass
    return empresa, periodo


def competencia_de_datas(serie_datas):
    """
    Retorna a competência no formato 'MM/YYYY' com base no mês/ano mais frequente
    de uma série de datas. Retorna None se não houver datas válidas.
    """
    s = pd.to_datetime(serie_datas, errors='coerce').dropna()
    if s.empty:
        return None
    moda = s.dt.to_period('M').mode()
    if moda.empty:
        return None
    return moda.iloc[0].strftime('%m/%Y')


def periodo_texto_para_competencia(periodo_texto):
    """
    Converte 'Periodo: 01/07/2026 a 31/07/2026' (já sem o prefixo) para 'MM/YYYY',
    usando a data inicial do intervalo.
    """
    if not periodo_texto:
        return None
    primeira_data = periodo_texto.split(' a ')[0].strip()
    dt = pd.to_datetime(primeira_data, dayfirst=True, errors='coerce')
    if pd.isna(dt):
        return None
    return dt.strftime('%m/%Y')


def processar_ifood(arquivo_ifood):
    """
    Detecta automaticamente qual dos dois modelos de relatório do iFood foi anexado
    e devolve (DataFrame padronizado, nome_da_loja_ou_None).

    Modelo 1 - "Extrato Financeiro" (fato_gerador / descricao_lancamento / data_repasse_esperada)
    Modelo 2 - "Relatório de Pedidos" (STATUS FINAL DO PEDIDO / VALOR DOS ITENS (R$))
               -> este modelo NÃO possui data de repasse, então DATA_PAGAMENTO fica em branco.
    """
    df_ifood = pd.read_excel(arquivo_ifood)
    colunas = set(df_ifood.columns)

    # ------------------------------------------------------------
    # MODELO 1: Extrato Financeiro (formato original já suportado)
    # ------------------------------------------------------------
    colunas_modelo_1 = {'fato_gerador', 'data_criacao_pedido_associado', 'descricao_lancamento', 'valor'}
    if colunas_modelo_1.issubset(colunas):
        vendas_ifood = df_ifood[df_ifood['fato_gerador'] == 'Venda'].copy()
        vendas_ifood['DATA_VENDA'] = pd.to_datetime(vendas_ifood['data_criacao_pedido_associado'], errors='coerce').dt.date

        if 'data_repasse_esperada' in colunas:
            vendas_ifood['DATA_PAGAMENTO'] = pd.to_datetime(vendas_ifood['data_repasse_esperada'], errors='coerce').dt.date
        else:
            vendas_ifood['DATA_PAGAMENTO'] = pd.NaT

        vendas_ifood['valor'] = pd.to_numeric(vendas_ifood['valor'], errors='coerce').fillna(0)

        bruto_ifood = (
            vendas_ifood[vendas_ifood['descricao_lancamento'] == 'Entrada Financeira']
            .groupby(['DATA_VENDA', 'DATA_PAGAMENTO'], dropna=False)['valor']
            .sum()
            .reset_index()
        )
        bruto_ifood.rename(columns={'valor': 'VALOR BRUTO'}, inplace=True)
        bruto_ifood['VALOR TAXA'] = 0.00
        bruto_ifood['TIPO_VENDA'] = 'iFood'
        return bruto_ifood[COLUNAS_PADRAO], None

    # ------------------------------------------------------------
    # MODELO 2: Relatório de Pedidos (novo modelo, sem data de repasse)
    # ------------------------------------------------------------
    colunas_modelo_2 = {'STATUS FINAL DO PEDIDO', 'DATA', 'VALOR DOS ITENS (R$)'}
    if colunas_modelo_2.issubset(colunas):
        # Considera CONCLUIDO e CANCELAMENTO PARCIAL (o valor que efetivamente ficou);
        # exclui apenas o CANCELADO total.
        vendas_ifood = df_ifood[df_ifood['STATUS FINAL DO PEDIDO'] != 'CANCELADO'].copy()

        vendas_ifood['DATA_VENDA'] = pd.to_datetime(vendas_ifood['DATA'], errors='coerce').dt.date
        # Este modelo de relatório não informa data de repasse/liquidação -> fica em branco
        vendas_ifood['DATA_PAGAMENTO'] = pd.NaT

        # VALOR DOS ITENS (R$) = valor de venda real (exclui taxa de entrega repassada ao entregador)
        vendas_ifood['VALOR BRUTO'] = pd.to_numeric(vendas_ifood['VALOR DOS ITENS (R$)'], errors='coerce').fillna(0)

        bruto_ifood = (
            vendas_ifood
            .groupby(['DATA_VENDA', 'DATA_PAGAMENTO'], dropna=False)['VALOR BRUTO']
            .sum()
            .reset_index()
        )
        bruto_ifood['VALOR TAXA'] = 0.00
        bruto_ifood['TIPO_VENDA'] = 'iFood'

        nome_loja = None
        if 'NOME DA LOJA' in colunas:
            modas = df_ifood['NOME DA LOJA'].dropna()
            if not modas.empty:
                nome_loja = modas.mode().iloc[0]

        return bruto_ifood[COLUNAS_PADRAO], nome_loja

    # ------------------------------------------------------------
    # Nenhum modelo reconhecido
    # ------------------------------------------------------------
    raise ValueError(
        "Modelo de planilha do iFood não reconhecido. "
        "Colunas encontradas não correspondem ao 'Extrato Financeiro' nem ao 'Relatório de Pedidos'."
    )


st.set_page_config(page_title="Automação Contábil", layout="centered")
st.title("📊 Consolidador Getnet e iFood")
st.write("Anexe a planilha da GETNET, do IFOOD, ou as duas (do mesmo período) para gerar a consolidação contábil diária.")

arquivo_getnet = st.file_uploader("📂 Anexe a planilha da GETNET (opcional)", type=['xlsx'])
arquivo_ifood = st.file_uploader("📂 Anexe a planilha do IFOOD (opcional)", type=['xlsx'])

if st.button("🚀 Processar Dados"):
    if arquivo_getnet is None and arquivo_ifood is None:
        st.warning("⚠️ Anexe ao menos uma planilha (Getnet e/ou iFood) para processar.")
    else:
        try:
            with st.spinner('Engrenagens girando... Processando matriz de dados! ⚙️'):
                df_vazio = pd.DataFrame(columns=COLUNAS_PADRAO)

                # ==========================================
                # 1. PROCESSAMENTO GETNET (se anexado)
                # ==========================================
                empresa_nome_getnet = None
                periodo_getnet_texto = None
                competencia_getnet = None

                if arquivo_getnet is not None:
                    empresa_nome_getnet, periodo_getnet_texto = extrair_info_empresa(arquivo_getnet)

                    # Cartões (Modalidade e Data de Pagamento)
                    df_cartoes = pd.read_excel(arquivo_getnet, sheet_name='CARTÕES', header=7)
                    df_cartoes = df_cartoes[df_cartoes['STATUS DA TRANSAÇÃO'] == 'Aprovada'].copy()
                    df_cartoes['DATA_VENDA'] = pd.to_datetime(df_cartoes['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                    df_cartoes['DATA_PAGAMENTO'] = pd.to_datetime(df_cartoes['DATA PREVISTA DO 1º PAGAMENTO'], dayfirst=True, errors='coerce').dt.date
                    df_cartoes['VALOR BRUTO'] = pd.to_numeric(df_cartoes['VALOR BRUTO'], errors='coerce').fillna(0)
                    df_cartoes['VALOR TAXA'] = pd.to_numeric(df_cartoes['VALOR TAXA'], errors='coerce').fillna(0)
                    df_cartoes['TIPO_VENDA'] = 'Getnet ' + df_cartoes['BANDEIRA'].astype(str) + ' ' + df_cartoes['MODALIDADE'].astype(str)
                    df_cartoes_limpo = df_cartoes[COLUNAS_PADRAO]

                    # PIX
                    df_pix = pd.read_excel(arquivo_getnet, sheet_name='PIX', header=7)
                    df_pix = df_pix[df_pix['STATUS'] == 'Paga'].copy()
                    df_pix['DATA_VENDA'] = pd.to_datetime(df_pix['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                    df_pix['DATA_PAGAMENTO'] = df_pix['DATA_VENDA']  # Liquidação imediata
                    df_pix['VALOR BRUTO'] = pd.to_numeric(df_pix['VALOR DA VENDA'], errors='coerce').fillna(0)
                    df_pix['VALOR TAXA'] = pd.to_numeric(df_pix['VALOR TAXA'], errors='coerce').fillna(0)
                    df_pix['TIPO_VENDA'] = 'Getnet PIX'
                    df_pix_limpo = df_pix[COLUNAS_PADRAO]

                    # Vouchers
                    df_voucher = pd.read_excel(arquivo_getnet, sheet_name='VOUCHER', header=7)
                    df_voucher = df_voucher[df_voucher['STATUS'] == 'Aprovada'].copy()
                    df_voucher['DATA_VENDA'] = pd.to_datetime(df_voucher['DATA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                    df_voucher['DATA_PAGAMENTO'] = df_voucher['DATA_VENDA']
                    df_voucher['VALOR BRUTO'] = pd.to_numeric(df_voucher['VALOR DA VENDA'], errors='coerce').fillna(0)
                    df_voucher['VALOR TAXA'] = 0.00
                    df_voucher['TIPO_VENDA'] = df_voucher['BANDEIRA'].replace({'Sodexo': 'Pluxee'})
                    df_voucher_limpo = df_voucher[COLUNAS_PADRAO]

                    # Agrupamento Getnet exigindo a chave DATA_PAGAMENTO
                    df_getnet_consolidado = pd.concat([df_cartoes_limpo, df_pix_limpo, df_voucher_limpo], ignore_index=True)
                    getnet_agrupado = df_getnet_consolidado.groupby(['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA'], dropna=False)[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()

                    # Competência: prioriza o texto declarado no cabeçalho da Getnet; usa as datas como reforço/fallback
                    competencia_getnet = periodo_texto_para_competencia(periodo_getnet_texto) or competencia_de_datas(getnet_agrupado['DATA_VENDA'])
                else:
                    df_cartoes_limpo = df_pix_limpo = df_voucher_limpo = df_vazio.copy()
                    getnet_agrupado = df_vazio.copy()

                # ==========================================
                # 2. PROCESSAMENTO IFOOD (se anexado, com detecção automática de modelo)
                # ==========================================
                nome_loja_ifood = None
                competencia_ifood = None

                if arquivo_ifood is not None:
                    ifood_agrupado, nome_loja_ifood = processar_ifood(arquivo_ifood)
                    competencia_ifood = competencia_de_datas(ifood_agrupado['DATA_VENDA'])
                else:
                    ifood_agrupado = df_vazio.copy()

                # ==========================================
                # 3. VALIDAÇÃO DE COMPETÊNCIA (só quando os 2 arquivos são anexados)
                # ==========================================
                if arquivo_getnet is not None and arquivo_ifood is not None:
                    if competencia_getnet and competencia_ifood and competencia_getnet != competencia_ifood:
                        st.error(
                            f"🚫 As planilhas são de competências diferentes e não podem ser processadas juntas.\n\n"
                            f"- Getnet: **{competencia_getnet}**\n"
                            f"- iFood: **{competencia_ifood}**\n\n"
                            f"Anexe relatórios do mesmo mês/ano, ou processe um de cada vez."
                        )
                        st.stop()

                empresa_nome = empresa_nome_getnet or nome_loja_ifood
                competencia = competencia_getnet or competencia_ifood

                # ==========================================
                # 4. CONSOLIDAÇÃO MATRICIAL GERAL
                # ==========================================
                diario_geral = pd.concat([getnet_agrupado, ifood_agrupado], ignore_index=True)
                diario_geral = diario_geral.sort_values(by=['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA']).reset_index(drop=True)

                resumo_geral = diario_geral.groupby('TIPO_VENDA')[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()
                linha_total = pd.DataFrame([{'TIPO_VENDA': 'TOTAL GERAL', 'VALOR BRUTO': resumo_geral['VALOR BRUTO'].sum(), 'VALOR TAXA': resumo_geral['VALOR TAXA'].sum()}])
                resumo_geral = pd.concat([resumo_geral, linha_total], ignore_index=True)

                # ==========================================
                # 5. ESCRITA NO BUFFER (Excel)
                # ==========================================
                buffer = io.BytesIO()
                # Reserva 2 linhas no topo da aba Resumo_Totais para Empresa/Competência
                linha_inicio_resumo = 3 if (empresa_nome or competencia) else 0

                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    diario_geral.to_excel(writer, sheet_name='Movimento_Diario', index=False)
                    resumo_geral.to_excel(writer, sheet_name='Resumo_Totais', index=False, startrow=linha_inicio_resumo)

                    workbook = writer.book
                    header_font = Font(bold=True, color="FFFFFF")
                    header_fill = PatternFill("solid", fgColor="2F4F4F")
                    info_font = Font(bold=True, color="2F4F4F")
                    border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                                    top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

                    # Cabeçalho de identificação (Empresa / Competência) na aba Resumo_Totais
                    if linha_inicio_resumo:
                        ws_resumo = workbook['Resumo_Totais']
                        ws_resumo.cell(row=1, column=1, value=f"Empresa: {empresa_nome or 'Não identificada'}").font = info_font
                        ws_resumo.cell(row=2, column=1, value=f"Competência: {competencia or 'Não identificada'}").font = info_font

                    # Linha do cabeçalho da tabela em cada aba (1-indexado)
                    linha_cabecalho = {'Movimento_Diario': 1, 'Resumo_Totais': linha_inicio_resumo + 1}

                    for sheet_name in workbook.sheetnames:
                        worksheet = workbook[sheet_name]
                        cab = linha_cabecalho.get(sheet_name, 1)
                        for cell in worksheet[cab]:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        for row in worksheet.iter_rows(min_row=cab + 1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                            for cell in row:
                                cell.border = border
                                if isinstance(cell.value, float):
                                    cell.number_format = '#,##0.00'
                                elif isinstance(cell.value, pd.Timestamp) or type(cell.value).__name__ == 'date':
                                    cell.number_format = 'DD/MM/YYYY'  # Força formatação visual para as datas
                        for col in worksheet.columns:
                            max_length = 0
                            column = col[0].column_letter
                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except: pass
                            worksheet.column_dimensions[column].width = (max_length + 2)

                # ==========================================
                # 6. INTERFACE DE USUÁRIO
                # ==========================================
                st.success("✨ Processamento concluído com sucesso!")

                total_getnet = df_cartoes_limpo['VALOR BRUTO'].sum() + df_pix_limpo['VALOR BRUTO'].sum()
                total_ifood = ifood_agrupado['VALOR BRUTO'].sum()
                total_vouchers = df_voucher_limpo['VALOR BRUTO'].sum()
                total_geral = total_getnet + total_ifood + total_vouchers

                st.subheader("📋 Resumo Operacional Bruto")
                st.markdown(f"🏢 **Empresa:** {empresa_nome or 'Não identificada'}  \n📅 **Competência:** {competencia or 'Não identificada'}")
                col1, col2 = st.columns(2)
                with col1:
                    if arquivo_getnet is not None:
                        st.markdown(f"💳 **Getnet (Cartões + PIX):** {formatar_moeda(total_getnet)}")
                    if arquivo_ifood is not None:
                        st.markdown(f"🍔 **iFood:** {formatar_moeda(total_ifood)}")
                    st.markdown(f"💰 **TOTAL GERAL:** {formatar_moeda(total_geral)}")
                with col2:
                    if arquivo_getnet is not None and not df_voucher_limpo.empty:
                        st.markdown("🎟️ **Vouchers (Detalhado):**")
                        vouchers_agrupados = df_voucher_limpo.groupby('TIPO_VENDA')['VALOR BRUTO'].sum()
                        for bandeira, valor in vouchers_agrupados.items():
                            st.markdown(f"- **{bandeira}:** {formatar_moeda(valor)}")

                if arquivo_ifood is not None and ifood_agrupado['DATA_PAGAMENTO'].isna().all():
                    st.caption("ℹ️ O modelo de planilha do iFood anexado não informa data de repasse/liquidação — a coluna DATA_PAGAMENTO ficou em branco para essas linhas.")

                st.markdown("---")

                st.download_button(
                    label="⬇️ Baixar Planilha Consolidada",
                    data=buffer.getvalue(),
                    file_name="Consolidado_Contabilidade.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"🚨 Erro na estrutura dos arquivos: {e}")
