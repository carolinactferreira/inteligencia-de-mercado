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
    moto = row.get('Moto', 0) if 'Moto' in row else 0
    carro = row.get('Carro', 0) if 'Carro' in row else 0
    try:
        moto = float(moto)
    except:
        moto = 0
    try:
        carro = float(carro)
    except:
        carro = 0

    if moto > 0:
        return 'Vaga Moto'
    elif carro > 0:
        return 'Vaga Carro'
    else:
        return 'Sem Vaga'

def mapear_tipologia(t):
    t_str = str(t).strip().upper()
    if '2 Q S' in t_str or '2QS' in t_str: return '2 Q S'
    elif '2 Q' in t_str or '2Q' in t_str: return '2 Q'
    return 'OUTRO'

def obter_coluna_flexivel(df, nomes_possiveis, valor_padrao=''):
    """Procura uma coluna no DataFrame ignorando maiúsculas/minúsculas e espaços."""
    colunas_df = {str(c).strip().lower(): c for c in df.columns}
    for nome in nomes_possiveis:
        nome_clean = nome.strip().lower()
        if nome_clean in colunas_df:
            return df[colunas_df[nome_clean]]
    return pd.Series([valor_padrao] * len(df), index=df.index)

def extrair_tabelas_pdf(file_bytes):
    todas_tabelas = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                if tabela and len(tabela) > 1:
                    header = [str(c).strip().replace('\n', ' ') if c is not None else '' for c in tabela[0]]
                    
                    seen = {}
                    header_unico = []
                    for h in header:
                        if h in seen:
                            seen[h] += 1
                            header_unico.append(f"{h}_{seen[h]}" if h else f"col_{seen[h]}")
                        else:
                            seen[h] = 0
                            header_unico.append(h if h else "col_0")
                            
                    df_tabela = pd.DataFrame(tabela[1:], columns=header_unico)
                    todas_tabelas.append(df_tabela)
                    
    if todas_tabelas:
        df_concat = pd.concat(todas_tabelas, ignore_index=True)
        df_concat.columns = [str(c).strip().replace('\n', ' ') for c in df_concat.columns]
        return df_concat
    return pd.DataFrame()

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
        col_emp = obter_coluna_flexivel(df, ['Empreendimento', 'Empreend', 'Nome do Empreendimento'], None)
        if col_emp.dropna().any():
            empreendimentos = col_emp.dropna().unique().tolist()
            emp_selecionado = st.selectbox("Selecione o Empreendimento:", empreendimentos)
            df_filtrado = df[col_emp == emp_selecionado].copy()
        else:
            emp_selecionado = "Empreendimento"
            df_filtrado = df.copy()
            
        if st.button("Gerar Relatório"):
            with st.spinner("Processando dados e formatando Excel..."):
                tipologia_series = obter_coluna_flexivel(df_filtrado, ['Tipologia', 'Tipo', 'Descrição', 'Descricao'], '')
                df_filtrado['Tipologia_clean'] = tipologia_series.astype(str).apply(mapear_tipologia)

                carro_series = obter_coluna_flexivel(df_filtrado, ['Carro', 'Vaga Carro', 'Vagas Carro'], 0)
                moto_series = obter_coluna_flexivel(df_filtrado, ['Moto', 'Vaga Moto', 'Vagas Moto'], 0)
                
                df_filtrado['Carro'] = pd.to_numeric(carro_series, errors='coerce').fillna(0)
                df_filtrado['Moto'] = pd.to_numeric(moto_series, errors='coerce').fillna(0)
                
                df_filtrado['Categoria Vaga'] = df_filtrado.apply(categorizar_vaga, axis=1)

                area_int = pd.to_numeric(obter_coluna_flexivel(df_filtrado, ['Área interna', 'Area interna', 'Área Privativa', 'Area Privativa'], 0), errors='coerce').fillna(0)
                quintal = pd.to_numeric(obter_coluna_flexivel(df_filtrado, ['Quintal'], 0), errors='coerce').fillna(0)
                varanda = pd.to_numeric(obter_coluna_flexivel(df_filtrado, ['Varanda'], 0), errors='coerce').fillna(0)
                
                df_filtrado['Área Total'] = area_int + quintal + varanda

                val_imovel = pd.to_numeric(obter_coluna_flexivel(df_filtrado, ['Valor do Imóvel', 'Valor Imovel', 'Valor da Unidade', 'Valor'], 0), errors='coerce').fillna(0)
                val_avaliacao = pd.to_numeric(obter_coluna_flexivel(df_filtrado, ['Valor da Avaliação', 'Valor Avaliacao', 'Avaliação', 'Avaliacao'], 0), errors='coerce').fillna(0)
                
                df_filtrado['Valor do Imóvel'] = val_imovel
                df_filtrado['Valor da Avaliação'] = val_avaliacao

                bloco_series = obter_coluna_flexivel(df_filtrado, ['Bloco', 'Torre'], '-')
                apto_series = obter_coluna_flexivel(df_filtrado, ['Apto', 'Unidade', 'Apartamento'], '-')

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
                            str(bloco_series.loc[idx] if idx in bloco_series.index else '-'),
                            str(apto_series.loc[idx] if idx in apto_series.index else '-'),
                            str(tipologia_series.loc[idx] if idx in tipologia_series.index else '-'),
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
                    file_name=f"Relatorio_Estoque_{str(emp_selecionado).replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                