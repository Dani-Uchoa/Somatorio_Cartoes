import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def processar_ifood(arquivo_ifood):
    """
    Detecta automaticamente qual dos dois modelos de relatório do iFood foi anexado
    e devolve um DataFrame padronizado com as colunas:
    DATA_VENDA, DATA_PAGAMENTO, TIPO_VENDA, VALOR BRUTO, VALOR TAXA

    Modelo 1 - "Extrato Financeiro" (fato_gerador / descricao_lancamento / data_repasse_esperada)
    Modelo 2 - "Relatório de Pedidos" (STATUS FINAL DO PEDIDO / TOTAL PAGO PELO CLIENTE (R$))
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
        return bruto_ifood[['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

    # ------------------------------------------------------------
    # MODELO 2: Relatório de Pedidos (novo modelo, sem data de repasse)
    # ------------------------------------------------------------
    colunas_modelo_2 = {'STATUS FINAL DO PEDIDO', 'DATA', 'TOTAL PAGO PELO CLIENTE (R$)'}
    if colunas_modelo_2.issubset(colunas):
        # Considera CONCLUIDO e CANCELAMENTO PARCIAL (o valor que efetivamente ficou);
        # exclui apenas o CANCELADO total.
        vendas_ifood = df_ifood[df_ifood['STATUS FINAL DO PEDIDO'] != 'CANCELADO'].copy()

        vendas_ifood['DATA_VENDA'] = pd.to_datetime(vendas_ifood['DATA'], errors='coerce').dt.date
        # Este modelo de relatório não informa data de repasse/liquidação -> fica em branco
        vendas_ifood['DATA_PAGAMENTO'] = pd.NaT

        vendas_ifood['VALOR BRUTO'] = pd.to_numeric(vendas_ifood['TOTAL PAGO PELO CLIENTE (R$)'], errors='coerce').fillna(0)

        bruto_ifood = (
            vendas_ifood
            .groupby(['DATA_VENDA', 'DATA_PAGAMENTO'], dropna=False)['VALOR BRUTO']
            .sum()
            .reset_index()
        )
        bruto_ifood['VALOR TAXA'] = 0.00
        bruto_ifood['TIPO_VENDA'] = 'iFood'
        return bruto_ifood[['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

    # ------------------------------------------------------------
    # Nenhum modelo reconhecido
    # ------------------------------------------------------------
    raise ValueError(
        "Modelo de planilha do iFood não reconhecido. "
        "Colunas encontradas não correspondem ao 'Extrato Financeiro' nem ao 'Relatório de Pedidos'."
    )


st.set_page_config(page_title="Automação Contábil", layout="centered")
st.title("📊 Consolidador Getnet e iFood")
st.write("Insira os relatórios originais abaixo para gerar a consolidação contábil diária.")

arquivo_getnet = st.file_uploader("📂 Anexe a planilha da GETNET", type=['xlsx'])
arquivo_ifood = st.file_uploader("📂 Anexe a planilha do IFOOD", type=['xlsx'])

if st.button("🚀 Processar Dados"):
    if arquivo_getnet is not None and arquivo_ifood is not None:
        try:
            with st.spinner('Engrenagens girando... Processando matriz de dados! ⚙️'):
                # ==========================================
                # 1. PROCESSAMENTO GETNET
                # ==========================================
                # Cartões (Modalidade e Data de Pagamento)
                df_cartoes = pd.read_excel(arquivo_getnet, sheet_name='CARTÕES', header=7)
                df_cartoes = df_cartoes[df_cartoes['STATUS DA TRANSAÇÃO'] == 'Aprovada'].copy()
                df_cartoes['DATA_VENDA'] = pd.to_datetime(df_cartoes['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_cartoes['DATA_PAGAMENTO'] = pd.to_datetime(df_cartoes['DATA PREVISTA DO 1º PAGAMENTO'], dayfirst=True, errors='coerce').dt.date
                df_cartoes['VALOR BRUTO'] = pd.to_numeric(df_cartoes['VALOR BRUTO'], errors='coerce').fillna(0)
                df_cartoes['VALOR TAXA'] = pd.to_numeric(df_cartoes['VALOR TAXA'], errors='coerce').fillna(0)
                df_cartoes['TIPO_VENDA'] = 'Getnet ' + df_cartoes['BANDEIRA'].astype(str) + ' ' + df_cartoes['MODALIDADE'].astype(str)
                df_cartoes_limpo = df_cartoes[['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                # PIX
                df_pix = pd.read_excel(arquivo_getnet, sheet_name='PIX', header=7)
                df_pix = df_pix[df_pix['STATUS'] == 'Paga'].copy()
                df_pix['DATA_VENDA'] = pd.to_datetime(df_pix['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_pix['DATA_PAGAMENTO'] = df_pix['DATA_VENDA']  # Liquidação imediata
                df_pix['VALOR BRUTO'] = pd.to_numeric(df_pix['VALOR DA VENDA'], errors='coerce').fillna(0)
                df_pix['VALOR TAXA'] = pd.to_numeric(df_pix['VALOR TAXA'], errors='coerce').fillna(0)
                df_pix['TIPO_VENDA'] = 'Getnet PIX'
                df_pix_limpo = df_pix[['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                # Vouchers
                df_voucher = pd.read_excel(arquivo_getnet, sheet_name='VOUCHER', header=7)
                df_voucher = df_voucher[df_voucher['STATUS'] == 'Aprovada'].copy()
                df_voucher['DATA_VENDA'] = pd.to_datetime(df_voucher['DATA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_voucher['DATA_PAGAMENTO'] = df_voucher['DATA_VENDA']
                df_voucher['VALOR BRUTO'] = pd.to_numeric(df_voucher['VALOR DA VENDA'], errors='coerce').fillna(0)
                df_voucher['VALOR TAXA'] = 0.00
                df_voucher['TIPO_VENDA'] = df_voucher['BANDEIRA'].replace({'Sodexo': 'Pluxee'})
                df_voucher_limpo = df_voucher[['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                # Agrupamento Getnet exigindo a chave DATA_PAGAMENTO
                df_getnet_consolidado = pd.concat([df_cartoes_limpo, df_pix_limpo, df_voucher_limpo], ignore_index=True)
                getnet_agrupado = df_getnet_consolidado.groupby(['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA'], dropna=False)[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()

                # ==========================================
                # 2. PROCESSAMENTO IFOOD (detecção automática de modelo)
                # ==========================================
                ifood_agrupado = processar_ifood(arquivo_ifood)

                # ==========================================
                # 3. CONSOLIDAÇÃO MATRICIAL GERAL
                # ==========================================
                diario_geral = pd.concat([getnet_agrupado, ifood_agrupado], ignore_index=True)
                diario_geral = diario_geral.sort_values(by=['DATA_VENDA', 'DATA_PAGAMENTO', 'TIPO_VENDA']).reset_index(drop=True)

                resumo_geral = diario_geral.groupby('TIPO_VENDA')[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()
                linha_total = pd.DataFrame([{'TIPO_VENDA': 'TOTAL GERAL', 'VALOR BRUTO': resumo_geral['VALOR BRUTO'].sum(), 'VALOR TAXA': resumo_geral['VALOR TAXA'].sum()}])
                resumo_geral = pd.concat([resumo_geral, linha_total], ignore_index=True)

                # ==========================================
                # 4. ESCRITA NO BUFFER (Excel)
                # ==========================================
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    diario_geral.to_excel(writer, sheet_name='Movimento_Diario', index=False)
                    resumo_geral.to_excel(writer, sheet_name='Resumo_Totais', index=False)

                    workbook = writer.book
                    header_font = Font(bold=True, color="FFFFFF")
                    header_fill = PatternFill("solid", fgColor="2F4F4F")
                    border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                                    top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

                    for sheet_name in workbook.sheetnames:
                        worksheet = workbook[sheet_name]
                        for cell in worksheet[1]:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
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
                # 5. INTERFACE DE USUÁRIO
                # ==========================================
                st.success("✨ Processamento concluído com sucesso!")

                total_getnet = df_cartoes_limpo['VALOR BRUTO'].sum() + df_pix_limpo['VALOR BRUTO'].sum()
                total_ifood = ifood_agrupado['VALOR BRUTO'].sum()
                total_vouchers = df_voucher_limpo['VALOR BRUTO'].sum()
                total_geral = total_getnet + total_ifood + total_vouchers

                st.subheader("📋 Resumo Operacional Bruto")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"💳 **Getnet (Cartões + PIX):** {formatar_moeda(total_getnet)}")
                    st.markdown(f"🍔 **iFood:** {formatar_moeda(total_ifood)}")
                    st.markdown(f"💰 **TOTAL GERAL:** {formatar_moeda(total_geral)}")
                with col2:
                    st.markdown("🎟️ **Vouchers (Detalhado):**")
                    vouchers_agrupados = df_voucher_limpo.groupby('TIPO_VENDA')['VALOR BRUTO'].sum()
                    for bandeira, valor in vouchers_agrupados.items():
                        st.markdown(f"- **{bandeira}:** {formatar_moeda(valor)}")

                if ifood_agrupado['DATA_PAGAMENTO'].isna().all():
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
    else:
        st.warning("⚠️ Anexe os dois arquivos antes de processar.")
