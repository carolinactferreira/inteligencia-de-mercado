import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import pdfplumber
import re

st.set_page_config(page_title="Relatório de Estoque", layout="wide")

st.title("📊 Gerador de Relatório de Estoque Imobiliário")
st.write("Faça o upload da planilha (.xlsx) ou arquivo PDF (.pdf) de estoque para gerar o relatório formatado.")

def converter_numero_br(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).strip().replace('R$', '').replace('$', '').replace(' ', '')
    if not s or s.lower() in ['nan', 'none', 'null', '<na>', '-']:
        return 0.0
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(re.sub(r'[^0-9.]', '', s))
    except:
        return 0.0

def limpar_texto(val, padrao='-'):
    if pd.isna(val) or val is None:
        return padrao
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', '<na>']:
        return padrao
    return s

def mapear_tipologia(t):
    t_str = str(t).strip().upper()
    if any(term in t_str for term in ['SUITE', 'SUÍTE', '2QS', '2 Q S', '2Q S', '2 QTS S', '2 DORM S', '2D S']):
        return '2 Q S'
    elif any(term in t_str for term in ['2Q', '2 Q', '2QTS', '2 QTS', '2 DORM', '2D', '2 QUARTOS']):
        return '2 Q'
    return 'OUTRA TIPOLOGIA'

def categorizar_vaga(row, col_moto='', col_carro='', col_vaga=''):
    moto_val = str(row.get(col_moto, '')).strip().upper() if col_moto else ''
    carro_val = str(row.get(col_carro, '')).strip().upper() if col_carro else ''
    vaga_gen = str(row.get(col_vaga, '')).strip().upper() if col_vaga else ''
    
    if moto_val in ['1', 'SIM', 'S', 'MOTO', 'TRUE', '1.0'] or 'MOTO' in vaga_gen:
        return 'Vaga Moto'
    elif carro_val in ['1', 'SIM', 'S', 'CARRO', 'TRUE', '1.0'] or 'CARRO' in vaga_gen or any(c in vaga_gen for c in ['1', '2', '3', 'SIM', 'S']):
        return 'Vaga Carro'
    
    if converter_numero_br(moto_val) > 0: return 'Vaga Moto'
    if converter_numero_br(carro_val) > 0: return 'Vaga Carro'

    return 'Sem Vaga'

def extrair_tabelas_pdf(file_bytes):
    todas_linhas = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                if tabela:
                    for r in tabela:
                        if any(r):
                            todas_linhas.append([str(c).strip().replace('\n', ' ') if c is not None else '' for c in r])
            if not tabelas:
                texto = page.extract_text()
                if texto:
                    for line in texto.split('\n'):
                        parts = line.split()
                        if len(parts) >= 3:
                            todas_linhas.append(parts)
                            
    if not todas_linhas:
        return pd.DataFrame()

    header_idx = 0
    for idx, row in enumerate(todas_linhas[:10]):
        row_str = " ".join(row).lower()
        if any(k in row_str for k in ['bloco', 'torre', 'apto', 'unidade', 'tipologia', 'valor', 'área', 'area', 'vaga']):
            header_idx = idx
            break
            
    header = todas_linhas[header_idx]
    seen = {}
    header_unico = []
    for h in header:
        clean_h = h if h else "col"
        if clean_h in seen:
            seen[clean_h] += 1
            header_unico.append(f"{clean_h}_{seen[clean_h]}")
        else:
            seen[clean_h] = 0
            header_unico.append(clean_h)
            
    dados = todas_linhas[header_idx + 1:]
    dados_limpos = []
    for r in dados:
        if len(r) == len(header_unico):
            if " ".join(r).lower() != " ".join(header).lower():
                dados_limpos.append(r)
        elif len(r) < len(header_unico):
            dados_limpos.append(r + [''] * (len(header_unico) - len(r)))
        else:
            dados_limpos.append(r[:len(header_unico)])
            
    return pd.DataFrame(dados_limpos, columns=header_unico)

uploaded_file = st.file_uploader("Arraste sua planilha (.xlsx) ou arquivo (.pdf) aqui", type=["xlsx", "pdf"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    if uploaded_file.name.endswith('.pdf'):
        df = extrair_tabelas_pdf(file_bytes)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
    if df.empty:
        st.error("Não foi possível extrair dados do arquivo enviado.")
    else:
        st.subheader("📋 Dados Brutos Identificados no Arquivo")
        st.write(f"Total de registros encontrados: **{len(df)} linhas**")
        st.dataframe(df.head(15), use_container_width=True)

        st.subheader("⚙️ Mapeamento de Colunas (Ajuste se necessário)")
        colunas_disponiveis = ["(Nenhuma / Não existe)"] + list(df.columns)

        def encontrar_melhor_coluna(opcoes, palavras_chave):
            for col in opcoes[1:]:
                col_clean = str(col).lower()
                if any(p in col_clean for p in palavras_chave):
                    return col
            return opcoes[0]

        col1, col2, col3 = st.columns(3)
        with col1:
            sel_bloco = st.selectbox("Coluna do Bloco/Torre:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['bloco', 'torre', 'tr'])))
            sel_apto = st.selectbox("Coluna do Apto/Unidade:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['apto', 'unidade', 'ap', 'up'])))
        with col2:
            sel_tipo = st.selectbox("Coluna da Tipologia/Planta:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['tipologia', 'tipo', 'planta', 'descricao', 'descrição', 'dorm'])))
            sel_area = st.selectbox("Coluna da Área Total/Privativa:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['área', 'area', 'm2', 'privativa'])))
        with col3:
            sel_valor = st.selectbox("Coluna do Valor do Imóvel:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['valor', 'preço', 'preco', 'venda', 'tabela'])))
            sel_avaliacao = st.selectbox("Coluna do Valor da Avaliação:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['avaliação', 'avaliacao', 'aval'])))

        st.markdown("---")
        
        col_vaga1, col_vaga2 = st.columns(2)
        with col_vaga1:
            sel_carro = st.selectbox("Coluna Vaga Carro:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['carro'])))
        with col_vaga2:
            sel_moto = st.selectbox("Coluna Vaga Moto:", colunas_disponiveis, index=colunas_disponiveis.index(encontrar_melhor_coluna(colunas_disponiveis, ['moto'])))

        if st.button("🚀 Gerar Relatório Formatado em Excel", type="primary"):
            with st.spinner("Processando dados e formatando Excel..."):
                
                s_bloco = df[sel_bloco] if sel_bloco in df.columns else pd.Series(['-']*len(df))
                s_apto = df[sel_apto] if sel_apto in df.columns else pd.Series(['-']*len(df))
                s_tipo = df[sel_tipo] if sel_tipo in df.columns else pd.Series(['-']*len(df))
                
                s_area = df[sel_area] if sel_area in df.columns else pd.Series([0]*len(df))
                s_valor = df[sel_valor] if sel_valor in df.columns else pd.Series([0]*len(df))
                s_aval = df[sel_avaliacao] if sel_avaliacao in df.columns else pd.Series([0]*len(df))

                df_proc = pd.DataFrame()
                df_proc['Bloco'] = s_bloco.apply(lambda x: limpar_texto(x, '-'))
                df_proc['Apto'] = s_apto.apply(lambda x: limpar_texto(x, '-'))
                df_proc['Tipologia_Orig'] = s_tipo.apply(lambda x: limpar_texto(x, '-'))
                df_proc['Tipologia_Clean'] = df_proc['Tipologia_Orig'].apply(mapear_tipologia)
                
                df_proc['Área'] = s_area.apply(converter_numero_br)
                df_proc['Valor_Imovel'] = s_valor.apply(converter_numero_br)
                df_proc['Valor_Avaliacao'] = s_aval.apply(converter_numero_br)

                col_c_name = sel_carro if sel_carro in df.columns else ''
                col_m_name = sel_moto if sel_moto in df.columns else ''
                
                df_proc['Categoria_Vaga'] = df.apply(lambda r: categorizar_vaga(r, col_moto=col_m_name, col_carro=col_c_name), axis=1)

                output = io.BytesIO()
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

                fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
                font_title = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
                headers = ['Bloco/Torre', 'Apto/Unidade', 'Tipologia Original', 'Área Total (m²)', 'Valor Imóvel (R$)', 'Valor Avaliação (R$)']

                # ABA 1: ESTOQUE GERAL
                ws_geral = wb.create_sheet(title="Estoque Geral")
                ws_geral.merge_cells('A1:F1')
                ws_geral['A1'] = "RELATÓRIO DE ESTOQUE — TODAS AS UNIDADES"
                ws_geral['A1'].font = font_title
                ws_geral['A1'].fill = fill_title
                ws_geral['A1'].alignment = Alignment(horizontal='center', vertical='center')
                
                ws_geral.append([])
                ws_geral.append(headers)

                for _, row in df_proc.iterrows():
                    ws_geral.append([
                        row['Bloco'],
                        row['Apto'],
                        row['Tipologia_Orig'],
                        row['Área'],
                        row['Valor_Imovel'],
                        row['Valor_Avaliacao']
                    ])

                # ABAS CATEGORIZADAS
                configuracoes = [
                    ('Moto 2QS', 'Vaga Moto', '2 Q S'),
                    ('Carro 2QS', 'Vaga Carro', '2 Q S'),
                    ('Sem Vaga 2QS', 'Sem Vaga', '2 Q S'),
                    ('Moto 2Q', 'Vaga Moto', '2 Q'),
                    ('Carro 2Q', 'Vaga Carro', '2 Q'),
                    ('Sem Vaga 2Q', 'Sem Vaga', '2 Q'),
                    ('Outras Unidades', 'TODAS', 'OUTRA TIPOLOGIA')
                ]

                for sheet_name, cat_key, tipo_key in configuracoes:
                    ws = wb.create_sheet(title=sheet_name)
                    
                    if cat_key == 'TODAS':
                        sub_df = df_proc[df_proc['Tipologia_Clean'] == tipo_key]
                    else:
                        sub_df = df_proc[(df_proc['Categoria_Vaga'] == cat_key) & (df_proc['Tipologia_Clean'] == tipo_key)]
                    
                    ws.merge_cells('A1:F1')
                    ws['A1'] = f"ESTOQUE — {cat_key.upper()} ({tipo_key})"
                    ws['A1'].font = font_title
                    ws['A1'].fill = fill_title
                    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    ws.append([])
                    ws.append(headers)
                    
                    for _, row in sub_df.iterrows():
                        ws.append([
                            row['Bloco'],
                            row['Apto'],
                            row['Tipologia_Orig'],
                            row['Área'],
                            row['Valor_Imovel'],
                            row['Valor_Avaliacao']
                        ])

                # Formatação de moedas e tamanhos no Excel
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(min_row=4, max_col=6):
                        if len(row) >= 6:
                            row[3].number_format = '#,##0.00'
                            row[4].number_format = 'R$ #,##0.00'
                            row[5].number_format = 'R$ #,##0.00'
                    for col in sheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = get_column_letter(col[0].column)
                        sheet.column_dimensions[col_letter].width = max(max_len + 3, 14)

                wb.save(output)
                processed_data = output.getvalue()
                
                st.success("Relatório gerado com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Relatório Excel Formatado",
                    data=processed_data,
                    file_name="Relatorio_Estoque_Formatado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )     