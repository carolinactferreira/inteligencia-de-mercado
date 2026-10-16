import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Relatório de Estoque", layout="wide")

st.title("📊 Gerador de Relatório de Estoque Imobiliário")
st.write("Faça o upload da planilha bruta de estoque para gerar o arquivo formatado automaticamente.")

# Função para categorizar vagas
def categorizar_vaga(row):
    if row.get('Moto', 0) > 0: return 'Vaga Moto'
    elif row.get('Carro', 0) > 0: return 'Vaga Carro'
    else: return 'Sem Vaga'

def mapear_tipologia(t):
    t_str = str(t).strip().upper()
    if '2 Q S' in t_str or '2QS' in t_str: return '2 Q S'
    elif '2 Q' in t_str or '2Q' in t_str: return '2 Q'
    return 'OUTRO'

uploaded_file = st.file_uploader("Arraste sua planilha (.xlsx) aqui", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Identificar empreendimentos disponíveis
    if 'Empreendimento' in df.columns:
        empreendimentos = df['Empreendimento'].dropna().unique().tolist()
        emp_selecionado = st.selectbox("Selecione o Empreendimento:", empreendimentos)
        
        if st.button("Gerar Relatório"):
            with st.spinner("Processando dados e formatando Excel..."):
                # Filtrar dados
                df_filtrado = df[df['Empreendimento'] == emp_selecionado].copy()
                
                # Limpeza
                colunas_num = ['Carro', 'Moto', 'Área interna', 'Quintal', 'Varanda', 'Valor do Imóvel', 'Valor da Avaliação']
                for col in colunas_num:
                    if col in df_filtrado.columns:
                        df_filtrado[col] = df_filtrado[col].fillna(0)
                        
                df_filtrado['Área Total'] = df_filtrado.get('Área interna', 0) + df_filtrado.get('Quintal', 0) + df_filtrado.get('Varanda', 0)
                df_filtrado['Tipologia_clean'] = df_filtrado.get('Tipologia', '').apply(mapear_tipologia)
                df_filtrado['Categoria Vaga'] = df_filtrado.apply(categorizar_vaga, axis=1)

                # Criar Excel na Memória
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

                # Estilos base
                fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
                font_title = Font(name='Calibri', size=14, bold=True, color='FFFFFF')

                for sheet_name, cat_key, tipo_key in configuracoes:
                    ws = wb.create_sheet(title=sheet_name)
                    sub_df = df_filtrado[(df_filtrado['Categoria Vaga'] == cat_key) & (df_filtrado['Tipologia_clean'] == tipo_key)]
                    
                    # Cabeçalho
                    ws.merge_cells('A1:I1')
                    ws['A1'] = f"{emp_selecionado} — {cat_key.upper()} ({tipo_key})"
                    ws['A1'].font = font_title
                    ws['A1'].fill = fill_title
                    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Tabela de Dados Simplificada para o código
                    headers = ['Bloco/Torre', 'Apto', 'Tipologia', 'Área Total', 'Valor Imóvel', 'Valor Avaliação']
                    ws.append([]) # Linha em branco
                    ws.append(headers)
                    
                    qtd_unidades = len(sub_df)
                    area_media = sub_df['Área Total'].mean() if qtd_unidades > 0 else 0
                    val_imovel_med = sub_df['Valor do Imóvel'].mean() if qtd_unidades > 0 else 0
                    
                    for idx, row in sub_df.iterrows():
                        ws.append([
                            str(row.get('Bloco', row.get('Torre', '-'))),
                            str(row.get('Apto', '-')),
                            str(row['Tipologia']),
                            round(row['Área Total'], 2),
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