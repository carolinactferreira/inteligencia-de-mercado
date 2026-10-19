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
    if not s or s.lower() in ['nan', 'none', 'null', '<na>', '(nenhuma / não existe)']:
        return padrao
    return s

def parse_concatenated_string(text_str):
    s = str(text_str).strip()
    
    apto_match = re.search(r'(?:APTO|APT|UNIDADE|UP|AP)\s*[:\.-]?\s*(\w+)', s, re.IGNORECASE)
    apto = apto_match.group(1) if apto_match else '-'
    
    tipo_match = re.search(r'\b(\d\s*Q\s*S|\d\s*QTS\s*S|\d\s*DORM\s*S|\d\s*Q|\d\s*QTS|\d\s*DORM|\d\s*D)\b', s, re.IGNORECASE)
    tipologia = tipo_match.group(1) if tipo_match else '-'
    
    clean_torre = re.split(r'\s*-\s*\d\s*Q|\s*-\s*\d\s*D|\s*-\s*APTO|\s*-\s*UNIDADE', s, flags=re.IGNORECASE)[0].strip()
    bloco = clean_torre if clean_torre else s
    
    return bloco, apto, tipologia

def mapear_tipologia(t):
    t_str = str(t).strip().upper()
    if not t_str or t_str in ['-', 'NAN', 'NONE']:
        return 'OUTRA TIPOLOGIA'
    
    has_suite = any(term in t_str for term in ['SUITE', 'SUÍTE', ' S', 'S ']) or t_str.endswith('S')
    
    if '3' in t_str:
        return '3 Q S' if has_suite else '3 Q'
    elif '2' in t_str:
        return '2 Q S' if has_suite else '2 Q'
    elif '1' in t_str:
        return '1 Q S' if has_suite else '1 Q'
        
    return 'OUTRA TIPOLOGIA'

def categorizar_vaga(row, col_moto='', col_carro=''):
    moto_val = str(row.get(col_moto, '')).strip().upper() if col_moto in row else ''
    carro_val = str(row.get(col_carro, '')).strip().upper() if col_carro in row else ''
    
    if moto_val in ['1', 'SIM', 'S', 'MOTO', 'TRUE', '1.0']:
        return 'Vaga Moto'
    elif carro_val in ['1', 'SIM', 'S', 'CARRO', 'TRUE', '1.0']:
        return 'Vaga Carro'
    
    if converter_numero_br(moto_val) > 0: return 'Vaga Moto'
    if converter_numero_br(carro_val) > 0: return 'Vaga Carro'

    return 'Sem Vaga'

def auto_detect_columns(df):
    mapping = {'empreendimento': None, 'bloco': None, 'apto': None, 'tipologia': None, 'area': None, 'valor_imovel': None, 'valor_avaliacao': None, 'carro': None, 'moto': None}
    
    for col in df.columns:
        c_lower = str(col).lower()
        if any(k in c_lower for k in ['empreendimento', 'empreend', 'projeto']) and not mapping['empreendimento']:
            mapping['empreendimento'] = col
        elif any(k in c_lower for k in ['bloco', 'torre']) and not mapping['bloco']:
            mapping['bloco'] = col
        elif any(k in c_lower for k in ['apto', 'unidade', 'apartamento']) and not mapping['apto']:
            mapping['apto'] = col
        elif any(k in c_lower for k in ['tipologia', 'tipo', 'planta', 'dorm']) and not mapping['tipologia']:
            mapping['tipologia'] = col
        elif any(k in c_lower for k in ['area', 'área', 'm2', 'privativa']) and not mapping['area']:
            mapping['area'] = col
        elif any(k in c_lower for k in ['venda', 'imovel', 'imóvel', 'preço', 'preco']) and 'aval' not in c_lower and not mapping['valor_imovel']:
            mapping['valor_imovel'] = col
        elif any(k in c_lower for k in ['aval']) and not mapping['valor_avaliacao']:
            mapping['valor_avaliacao'] = col
        elif 'carro' in c_lower and not mapping['carro']:
            mapping['carro'] = col
        elif 'moto' in c_lower and not mapping['moto']:
            mapping['moto'] = col

    sample_rows = df.head(10)
    for col in df.columns:
        vals = [str(v).strip() for v in sample_rows[col].dropna()]
        if not mapping['bloco'] and any('TORRE' in v.upper() or 'BLOCO' in v.upper() for v in vals) and any('APTO' in v.upper() for v in vals):
            mapping['bloco'] = col
        if not mapping['area']:
            numeric_vals = []
            for v in vals:
                try:
                    vf = float(v.replace(',', '.'))
                    if 10.0 <= vf <= 400.0: numeric_vals.append(vf)
                except: pass
            if len(numeric_vals) >= len(vals) * 0.4 and len(vals) > 0:
                mapping['area'] = col
    return mapping

def get_area_series(df, sel_area):
    if sel_area in df.columns and sel_area != "(Nenhuma / Não existe)":
        return df[sel_area]
    for col in df.columns:
        vals = df[col].dropna()
        parsed_vals = []
        for v in vals:
            try:
                vf = float(str(v).replace(',', '.'))
                if 15.0 <= vf <= 400.0: parsed_vals.append(vf)
            except: pass
        if len(parsed_vals) >= len(vals) * 0.4 and len(vals) > 0:
            return df[col]
    return pd.Series([0] * len(df))

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

        st.subheader("⚙️ Mapeamento de Colunas (Ajustado Automaticamente)")
        colunas_disponiveis = ["(Nenhuma / Não existe)"] + list(df.columns)
        mapping_auto = auto_detect_columns(df)

        def idx_of(val):
            return colunas_disponiveis.index(val) if val in colunas_disponiveis else 0

        # Primeira Linha de mapeamento com o EMPREENDIMENTO
        col_emp_ui, col1, col2 = st.columns(3)
        with col_emp_ui:
            sel_emp_col = st.selectbox("Coluna do Empreendimento:", colunas_disponiveis, index=idx_of(mapping_auto['empreendimento']))
        with col1:
            sel_bloco = st.selectbox("Coluna do Bloco/Torre:", colunas_disponiveis, index=idx_of(mapping_auto['bloco']))
        with col2:
            sel_apto = st.selectbox("Coluna do Apto/Unidade:", colunas_disponiveis, index=idx_of(mapping_auto['apto']))

        col3, col4, col5 = st.columns(3)
        with col3:
            sel_tipo = st.selectbox("Coluna da Tipologia/Planta:", colunas_disponiveis, index=idx_of(mapping_auto['tipologia']))
        with col4:
            sel_area = st.selectbox("Coluna da Área Total/Privativa:", colunas_disponiveis, index=idx_of(mapping_auto['area']))
        with col5:
            sel_valor = st.selectbox("Coluna do Valor do Imóvel:", colunas_disponiveis, index=idx_of(mapping_auto['valor_imovel']))

        col6, col7, col8 = st.columns(3)
        with col6:
            sel_avaliacao = st.selectbox("Coluna Valor da Avaliação:", colunas_disponiveis, index=idx_of(mapping_auto['valor_avaliacao']))
        with col7:
            sel_carro = st.selectbox("Coluna Vaga Carro:", colunas_disponiveis, index=idx_of(mapping_auto['carro']))
        with col8:
            sel_moto = st.selectbox("Coluna Vaga Moto:", colunas_disponiveis, index=idx_of(mapping_auto['moto']))

        st.markdown("---")

        # OPÇÃO DE ESCOLHER O EMPREENDIMENTO
        if sel_emp_col != "(Nenhuma / Não existe)":
            lista_empreendimentos = [str(e).strip() for e in df[sel_emp_col].dropna().unique() if str(e).strip() != '']
            if len(lista_empreendimentos) > 1:
                st.subheader("🏢 Seleção de Empreendimento")
                st.info("Foram encontrados vários empreendimentos neste arquivo. Selecione qual deseja exportar ou processe todos juntos.")
                emp_escolhido = st.selectbox("Escolha como gerar o relatório:", ["Processar Todos (Juntos)"] + lista_empreendimentos)
            else:
                emp_escolhido = lista_empreendimentos[0] if lista_empreendimentos else "Processar Todos (Juntos)"
        else:
            emp_escolhido = "Processar Todos (Juntos)"

        if st.button("🚀 Gerar Relatório Formatado em Excel", type="primary"):
            with st.spinner("Processando dados e formatando Excel..."):
                
                # FILTRA OS DADOS SE UM EMPREENDIMENTO ESPECÍFICO FOI ESCOLHIDO
                if emp_escolhido != "Processar Todos (Juntos)":
                    df_filtrado = df[df[sel_emp_col].astype(str).str.strip() == emp_escolhido].copy()
                    nome_aba_principal = emp_escolhido.upper()
                else:
                    df_filtrado = df.copy()
                    nome_aba_principal = "ESTOQUE GERAL"

                # Séries dos dados
                s_emp = df_filtrado[sel_emp_col] if sel_emp_col in df_filtrado.columns else pd.Series(['-']*len(df_filtrado))
                s_bloco = df_filtrado[sel_bloco] if sel_bloco in df_filtrado.columns else pd.Series(['-']*len(df_filtrado))
                s_apto = df_filtrado[sel_apto] if sel_apto in df_filtrado.columns else pd.Series(['-']*len(df_filtrado))
                s_tipo = df_filtrado[sel_tipo] if sel_tipo in df_filtrado.columns else pd.Series(['-']*len(df_filtrado))
                s_area = get_area_series(df_filtrado, sel_area)
                s_valor = df_filtrado[sel_valor] if sel_valor in df_filtrado.columns else pd.Series([0]*len(df_filtrado))
                s_aval = df_filtrado[sel_avaliacao] if sel_avaliacao in df_filtrado.columns else pd.Series([0]*len(df_filtrado))

                res_emp, res_bloco, res_apto, res_tipo_orig = [], [], [], []

                for idx in range(len(df_filtrado)):
                    raw_e = str(s_emp.iloc[idx]).strip()
                    raw_b = str(s_bloco.iloc[idx]).strip()
                    raw_a = str(s_apto.iloc[idx]).strip()
                    raw_t = str(s_tipo.iloc[idx]).strip()

                    if ('TORRE' in raw_b.upper() or 'BLOCO' in raw_b.upper()) and ('APTO' in raw_b.upper() or 'Q' in raw_b.upper()):
                        ext_b, ext_a, ext_t = parse_concatenated_string(raw_b)
                        final_b = ext_b if ext_b else raw_b
                        final_a = raw_a if raw_a not in ['-', 'nan', 'None', '', '(Nenhuma / Não existe)'] else ext_a
                        final_t = raw_t if raw_t not in ['-', 'nan', 'None', '', '(Nenhuma / Não existe)'] else ext_t
                    else:
                        final_b = raw_b if raw_b not in ['nan', 'None', ''] else '-'
                        final_a = raw_a if raw_a not in ['nan', 'None', ''] else '-'
                        final_t = raw_t if raw_t not in ['nan', 'None', ''] else '-'

                    res_emp.append(limpar_texto(raw_e))
                    res_bloco.append(limpar_texto(final_b))
                    res_apto.append(limpar_texto(final_a))
                    res_tipo_orig.append(limpar_texto(final_t))

                df_proc = pd.DataFrame({
                    'Empreendimento': res_emp,
                    'Bloco': res_bloco,
                    'Apto': res_apto,
                    'Tipologia_Orig': res_tipo_orig,
                    'Área': [converter_numero_br(v) for v in s_area],
                    'Valor_Imovel': [converter_numero_br(v) for v in s_valor],
                    'Valor_Avaliacao': [converter_numero_br(v) for v in s_aval],
                    'Categoria_Vaga': [categorizar_vaga(row, col_moto=sel_moto, col_carro=sel_carro) for _, row in df_filtrado.iterrows()]
                })
                df_proc['Tipologia_Clean'] = df_proc['Tipologia_Orig'].apply(mapear_tipologia)

                output = io.BytesIO()
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

                fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
                font_title = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
                headers = ['Empreendimento', 'Bloco/Torre', 'Apto/Unidade', 'Tipologia Original', 'Área Total (m²)', 'Valor Imóvel (R$)', 'Valor Avaliação (R$)']

                # ABA 1: ESTOQUE GERAL
                ws_geral = wb.create_sheet(title="Estoque Geral")
                ws_geral.merge_cells('A1:G1')
                ws_geral['A1'] = f"{nome_aba_principal} — TODAS AS UNIDADES"
                ws_geral['A1'].font = font_title
                ws_geral['A1'].fill = fill_title
                ws_geral['A1'].alignment = Alignment(horizontal='center', vertical='center')
                
                ws_geral.append([])
                ws_geral.append(headers)

                for _, row in df_proc.iterrows():
                    ws_geral.append([
                        row['Empreendimento'], row['Bloco'], row['Apto'], row['Tipologia_Orig'],
                        row['Área'], row['Valor_Imovel'], row['Valor_Avaliacao']
                    ])

                # ABAS CATEGORIZADAS
                configuracoes = [
                    ('Moto 1Q', 'Vaga Moto', '1 Q'),
                    ('Carro 1Q', 'Vaga Carro', '1 Q'),
                    ('Sem Vaga 1Q', 'Sem Vaga', '1 Q'),
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
                    
                    ws.merge_cells('A1:G1')
                    ws['A1'] = f"{nome_aba_principal} — {cat_key.upper()} ({tipo_key})"
                    ws['A1'].font = font_title
                    ws['A1'].fill = fill_title
                    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    ws.append([])
                    ws.append(headers)
                    
                    for _, row in sub_df.iterrows():
                        ws.append([
                            row['Empreendimento'], row['Bloco'], row['Apto'], row['Tipologia_Orig'],
                            row['Área'], row['Valor_Imovel'], row['Valor_Avaliacao']
                        ])

                # Formatação no Excel
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(min_row=4, max_col=7):
                        if len(row) >= 7:
                            row[4].number_format = '#,##0.00'
                            row[5].number_format = 'R$ #,##0.00'
                            row[6].number_format = 'R$ #,##0.00'
                    for col in sheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = get_column_letter(col[0].column)
                        sheet.column_dimensions[col_letter].width = max(max_len + 3, 14)

                wb.save(output)
                processed_data = output.getvalue()
                
                st.success("Relatório gerado com sucesso!")
                
                nome_f = f"Relatorio_{emp_escolhido.replace(' ', '_')}.xlsx" if emp_escolhido != "Processar Todos (Juntos)" else "Relatorio_Estoque_Geral_Completo.xlsx"
                
                st.download_button(
                    label="📥 Baixar Relatório Excel Formatado",
                    data=processed_data,
                    file_name=nome_f,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                