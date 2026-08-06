import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Automação Contábil", layout="centered")
st.title("Consolidador Getnet e iFood")

# Componentes de Upload da Interface
arquivo_getnet = st.file_uploader("Anexe a planilha da GETNET", type=['xlsx'])
arquivo_ifood = st.file_uploader("Anexe a planilha do IFOOD", type=['xlsx'])

if st.button("Processar Dados"):
    if arquivo_getnet is not None and arquivo_ifood is not None:
        try:
            with st.spinner('Processando matriz de dados...'):
                # O pandas lê diretamente o arquivo injetado pelo navegador
                df_cartoes = pd.read_excel(arquivo_getnet, sheet_name='CARTÕES', header=7)
                df_cartoes = df_cartoes[df_cartoes['STATUS DA TRANSAÇÃO'] == 'Aprovada'].copy()
                df_cartoes['DATA_VENDA'] = pd.to_datetime(df_cartoes['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_cartoes['VALOR BRUTO'] = pd.to_numeric(df_cartoes['VALOR BRUTO'], errors='coerce').fillna(0)
                df_cartoes['VALOR TAXA'] = pd.to_numeric(df_cartoes['VALOR TAXA'], errors='coerce').fillna(0)
                df_cartoes['TIPO_VENDA'] = 'Getnet ' + df_cartoes['BANDEIRA'].astype(str)
                df_cartoes_limpo = df_cartoes[['DATA_VENDA', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                df_pix = pd.read_excel(arquivo_getnet, sheet_name='PIX', header=7)
                df_pix = df_pix[df_pix['STATUS'] == 'Paga'].copy()
                df_pix['DATA_VENDA'] = pd.to_datetime(df_pix['DATA/HORA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_pix['VALOR BRUTO'] = pd.to_numeric(df_pix['VALOR DA VENDA'], errors='coerce').fillna(0)
                df_pix['VALOR TAXA'] = pd.to_numeric(df_pix['VALOR TAXA'], errors='coerce').fillna(0)
                df_pix['TIPO_VENDA'] = 'Getnet PIX'
                df_pix_limpo = df_pix[['DATA_VENDA', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                df_voucher = pd.read_excel(arquivo_getnet, sheet_name='VOUCHER', header=7)
                df_voucher = df_voucher[df_voucher['STATUS'] == 'Aprovada'].copy()
                df_voucher['DATA_VENDA'] = pd.to_datetime(df_voucher['DATA DA VENDA'], dayfirst=True, errors='coerce').dt.date
                df_voucher['VALOR BRUTO'] = pd.to_numeric(df_voucher['VALOR DA VENDA'], errors='coerce').fillna(0)
                df_voucher['VALOR TAXA'] = 0.00 
                df_voucher['TIPO_VENDA'] = df_voucher['BANDEIRA'].replace({'Sodexo': 'Pluxee'})
                df_voucher_limpo = df_voucher[['DATA_VENDA', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                df_getnet_consolidado = pd.concat([df_cartoes_limpo, df_pix_limpo, df_voucher_limpo], ignore_index=True)
                getnet_agrupado = df_getnet_consolidado.groupby(['DATA_VENDA', 'TIPO_VENDA'])[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()

                df_ifood = pd.read_excel(arquivo_ifood)
                vendas_ifood = df_ifood[df_ifood['fato_gerador'] == 'Venda'].copy()
                vendas_ifood['DATA_VENDA'] = pd.to_datetime(vendas_ifood['data_criacao_pedido_associado']).dt.date
                vendas_ifood['valor'] = pd.to_numeric(vendas_ifood['valor'], errors='coerce').fillna(0)

                bruto_ifood = vendas_ifood[vendas_ifood['descricao_lancamento'] == 'Entrada Financeira'].groupby('DATA_VENDA')['valor'].sum().reset_index()
                bruto_ifood.rename(columns={'valor': 'VALOR BRUTO'}, inplace=True)
                bruto_ifood['VALOR TAXA'] = 0.00
                bruto_ifood['TIPO_VENDA'] = 'iFood'
                ifood_agrupado = bruto_ifood[['DATA_VENDA', 'TIPO_VENDA', 'VALOR BRUTO', 'VALOR TAXA']]

                diario_geral = pd.concat([getnet_agrupado, ifood_agrupado], ignore_index=True)
                diario_geral = diario_geral.sort_values(by=['DATA_VENDA', 'TIPO_VENDA']).reset_index(drop=True)

                resumo_geral = diario_geral.groupby('TIPO_VENDA')[['VALOR BRUTO', 'VALOR TAXA']].sum().reset_index()
                linha_total = pd.DataFrame([{'TIPO_VENDA': 'TOTAL GERAL', 'VALOR BRUTO': resumo_geral['VALOR BRUTO'].sum(), 'VALOR TAXA': resumo_geral['VALOR TAXA'].sum()}])
                resumo_geral = pd.concat([resumo_geral, linha_total], ignore_index=True)

                # Escrita em Buffer de Memória (Exigência para Nuvem)
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
                        for col in worksheet.columns:
                            max_length = 0
                            column = col[0].column_letter
                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except: pass
                            worksheet.column_dimensions[column].width = (max_length + 2)

                # Injeção do Botão de Download na Interface
                st.success("Planilha processada com sucesso!")
                st.download_button(
                    label="⬇️ Baixar Planilha Consolidada",
                    data=buffer.getvalue(),
                    file_name="Consolidado_Contabilidade.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro na estrutura dos arquivos: {e}")
    else:
        st.warning("Anexe os dois arquivos antes de processar.")