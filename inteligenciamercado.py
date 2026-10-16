import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import pdfplumber

st.set_page_config(page_title="Relatório de Estoque", layout="wide")

st.title("📊 Gerador de Relatório de Estoque Imobiliário")
st.write("Faça o upload da planilha (.xlsx) ou arquivo PDF (.pdf) de estoque para gerar o relatório formatado.")

def categorizar_vaga(row):
    if row.get('Moto', 0) > 0: return 'Vaga Moto'
    elif row.get('Carro', 0) > 0: return 'Vaga Carro'
    else: return 'Sem Vaga'

def mapear_tipologia(t):
    t_str = str(t).strip().upper()
    if '2 Q S' in t_str or '2QS' in t_str: return '2 Q S'
    elif '2 Q' in t_str or '2Q' in t_str: return '2 Q'
    return 'OUTRO'

def extrair_tabelas_pdf(file_bytes):
    todas_tabelas = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                if tabela and len(tabela) > 1:
                    df_tabela = pd.DataFrame(tabela[1:], columns=tabela[0])
                    todas_tabelas.append(df_tabela)
    if todas_tabelas:
        return pd.concat(todas_tabelas, ignore_index=True)
    return pd.DataFrame()

# Aceita tanto .xlsx quanto .pdf
uploaded_file = st.file_uploader("Arraste sua planilha (.xlsx) ou arquivo (.pdf) aqui", type=["xlsx", "pdf"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    if uploaded_file.name.endswith('.pdf'):
        df = extrair_tabelas_pdf(file_bytes)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))
        
    if df.empty:
        st.error("Não foi possível extrair dados do arquivo enviado.")
    else:
        # Identificar empreendimentos disponíveis
        if 'Empreendimento' in df.columns:
            empreendimentos = df['Empreendimento'].dropna().unique().tolist()
            emp_selecionado = st.selectbox("Selecione o Empreendimento:", empreendimentos)
        else:
            emp_selecionado = "Empreendimento"
            
        if st.button("Gerar Relatório"):
            with st.spinner("Processando dados e formatando Excel..."):
                df_filtrado = df[df['Empreendimento'] == emp_selecionado].copy() if 'Empreendimento' in df.columns else df.copy()
                
                colunas_num = ['Carro', 'Moto', 'Área interna', 'Quintal', 'Varanda', 'Valor do Imóvel', 'Valor da Avaliação']
                for col in colunas_num:
                    if col in df_filtrado.columns:
                        df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors='coerce').fillna(0)
                        
                df_filtrado['Área Total'] = df_filtrado.get('Área interna', 0) + df_filtrado.get('Quintal', 0) + df_filtrado.get('Varanda', 0)
                df_filtrado['Tipologia_clean'] = df_filtrado.get('Tipologia', '').apply(mapear_tipologia)
                df_filtrado['Categoria Vaga'] = df_filtrado.apply(categorizar_vaga, axis=1)

                output = io.BytesIO()
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

                configuracoes = [
                    ('Moto 2QS', 'Vaga Moto', '2 Q S'),
                    ('Carro 2QS', 'Vaga Carro', '2 Q S'),
                    ('Sem Vaga 2QS', 'Sem Vaga', '2 Q S'),
                    ('Moto 2Q', 'Vaga Moto', '2 Q'),
                    ('Carro 2Q', 'Vaga Carro', '2 Q'),
                    ('Sem Vaga 2Q', 'Sem Vaga', '2 Q')
                ]

                fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
                font_title = Font(name='Calibri', size=14, bold=True, color='FFFFFF')

                for sheet_name, cat_key, tipo_key in configuracoes:
                    ws = wb.create_sheet(title=sheet_name)
                    sub_df = df_filtrado[(df_filtrado['Categoria Vaga'] == cat_key) & (df_filtrado['Tipologia_clean'] == tipo_key)]
                    
                    ws.merge_cells('A1:F1')
                    ws['A1'] = f"{emp_selecionado} — {cat_key.upper()} ({tipo_key})"
                    ws['A1'].font = font_title
                    ws['A1'].fill = fill_title
                    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    headers = ['Bloco/Torre', 'Apto', 'Tipologia', 'Área Total', 'Valor Imóvel', 'Valor Avaliação']
                    ws.append([])
                    ws.append(headers)
                    
                    for idx, row in sub_df.iterrows():
                        ws.append([
                            str(row.get('Bloco', row.get('Torre', '-'))),
                            str(row.get('Apto', '-')),
                            str(row.get('Tipologia', '-')),
                            round(row.get('Área Total', 0), 2),
                            round(row.get('Valor do Imóvel', 0), 2),
                            round(row.get('Valor da Avaliação', 0), 2)
                        ])

                wb.save(output)
                processed_data = output.getvalue()
                
                st.success("Relatório gerado com sucesso!")
                
                st.download_button(
                    label="📥 Baixar Relatório Excel",
                    data=processed_data,
                    file_name=f"Relatorio_Estoque_{emp_selecionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )