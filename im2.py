# -*- coding: utf-8 -*-
"""
===============================================================================
AUTOMAÇÃO EM PYTHON: ATUALIZAÇÃO DE DESEMPENHO E ESTOQUE DA CONCORRÊNCIA
===============================================================================
Este script funciona de forma 100% independente, sem depender do Google Drive.

Bibliotecas necessárias:
    pip install pandas openpyxl
"""

import os
import re
import pandas as pd
import openpyxl

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def padronizar_tipologia(tipologia):
    t = normalizar_texto(str(tipologia))
    if 'suíte' in t or 'suite' in t or '2qs' in t or 'qs' in t:
        if '2 quartos + 1' in t:
            return '2 quartos + 1 suíte'
        elif '2 suítes' in t or '2 suites' in t:
            return '2 suítes'
        return '2 quartos (suíte)'
    elif '2q' in t or '2 quartos' in t or '2 dorm' in t:
        return '2 quartos'
    elif '1q' in t or '1 quarto' in t or '1 dorm' in t:
        return '1 quarto'
    elif '3q' in t or '3 quartos' in t or '3 dorm' in t:
        return '3 quartos'
    return str(tipologia).strip()

def padronizar_vaga(vaga):
    v = normalizar_texto(str(vaga))
    if 'moto' in v:
        return 'vaga moto'
    elif 'carro' in v or 'coberta' in v or 'descoberta' in v or '1 vaga' in v or '2 vagas' in v:
        return 'vaga carro'
    elif 'sem' in v or 'nenhum' in v or '0' in v or v == '':
        return 'sem vaga'
    return str(vaga).strip().lower()

def carregar_relatorio_streamlit(caminho_streamlit):
    """
    Carrega o relatório exportado pelo Streamlit (em CSV ou XLSX)
    e consolida os KPIs agrupados por (Empreendimento, Tipologia, Tipologia Vaga).
    """
    if caminho_streamlit.lower().endswith('.csv'):
        df = pd.read_csv(caminho_streamlit)
    else:
        df = pd.read_excel(caminho_streamlit)
    
    col_map = {}
    for col in df.columns:
        c_norm = normalizar_texto(col)
        if 'emp' in c_norm or 'nome' in c_norm or 'projeto' in c_norm:
            col_map['empreendimento'] = col
        elif 'tipo' in c_norm and 'vaga' not in c_norm:
            col_map['tipologia'] = col
        elif 'vaga' in c_norm or 'garagem' in c_norm:
            col_map['vaga'] = col
        elif 'área' in c_norm or 'area' in c_norm:
            col_map['area'] = col
        elif 'valor' in c_norm or 'preço' in c_norm or 'preco' in c_norm:
            if 'avalia' in c_norm:
                col_map['avaliacao'] = col
            else:
                col_map['valor_imovel'] = col
        elif 'estoque' in c_norm or 'unidade' in c_norm or 'qtd' in c_norm:
            col_map['estoque'] = col

    if 'estoque' not in col_map:
        df['_qtd_estoque'] = 1
        col_map['estoque'] = '_qtd_estoque'

    emp_col = col_map.get('empreendimento', df.columns[0])
    tipo_col = col_map.get('tipologia', df.columns[1] if len(df.columns) > 1 else df.columns[0])
    vaga_col = col_map.get('vaga', df.columns[2] if len(df.columns) > 2 else df.columns[0])
    area_col = col_map.get('area', None)
    valor_col = col_map.get('valor_imovel', None)
    aval_col = col_map.get('avaliacao', None)
    est_col = col_map['estoque']

    df['emp_padrao'] = df[emp_col].astype(str).str.strip()
    df['tipo_padrao'] = df[tipo_col].apply(padronizar_tipologia)
    df['vaga_padrao'] = df[vaga_col].apply(padronizar_vaga)

    agg_dict = {est_col: 'sum'}
    if area_col: agg_dict[area_col] = 'mean'
    if valor_col: agg_dict[valor_col] = 'mean'
    if aval_col: agg_dict[aval_col] = 'mean'

    resumo = df.groupby(['emp_padrao', 'tipo_padrao', 'vaga_padrao'], as_index=False).agg(agg_dict)
    
    rename_map = {
        'emp_padrao': 'Empreendimento',
        'tipo_padrao': 'Tipologia',
        'vaga_padrao': 'Tipologia vaga',
        est_col: 'Unidades em estoque'
    }
    if area_col: rename_map[area_col] = 'Área total'
    if valor_col: rename_map[valor_col] = 'Valor do imóvel'
    if aval_col: rename_map[aval_col] = 'Valor da avaliação'

    resumo = resumo.rename(columns=rename_map)
    return resumo

def atualizar_planilha_performance(caminho_performance, resumo_streamlit, mes_alvo):
    """
    Atualiza a aba 'mes_alvo' no arquivo Performance_concorrencia.xlsx com as métricas extraídas.
    """
    if not os.path.exists(caminho_performance):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_performance}")

    wb = openpyxl.load_workbook(caminho_performance)
    
    if mes_alvo not in wb.sheetnames:
        raise ValueError(f"A aba '{mes_alvo}' não existe. Abas disponíveis: {wb.sheetnames}")

    ws = wb[mes_alvo]
    
    header_row_idx = None
    headers = {}
    
    for row_idx in range(1, 10):
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, 15)]
        row_str = [normalizar_texto(str(v)) for v in row_vals if v is not None]
        if 'nome do empreendimento' in row_str or 'empreendimento' in row_str:
            header_row_idx = row_idx
            for col_idx, val in enumerate(row_vals, start=1):
                if val:
                    headers[normalizar_texto(str(val))] = col_idx
            break

    if not header_row_idx:
        raise ValueError("Cabeçalho da tabela não encontrado na aba informada.")

    col_emp = headers.get('nome do empreendimento') or headers.get('empreendimento')
    col_tipo = headers.get('tipologia')
    col_vaga = headers.get('tipologia vaga') or headers.get('vaga')
    col_est = headers.get('unidades em estoque') or headers.get('estoque')
    col_area = headers.get('área total') or headers.get('area total')
    col_valor = headers.get('valor do imóvel') or headers.get('valor do imovel')
    col_aval = headers.get('valor da avaliação') or headers.get('valor da avaliacao')

    atualizados = 0
    max_row = ws.max_row

    for row in range(header_row_idx + 1, max_row + 1):
        emp_val = ws.cell(row=row, column=col_emp).value
        tipo_val = ws.cell(row=row, column=col_tipo).value
        vaga_val = ws.cell(row=row, column=col_vaga).value if col_vaga else None

        if not emp_val or not tipo_val:
            continue

        emp_norm = str(emp_val).strip()
        tipo_norm = padronizar_tipologia(tipo_val)
        vaga_norm = padronizar_vaga(vaga_val) if vaga_val else 'sem vaga'

        match = resumo_streamlit[
            (resumo_streamlit['Empreendimento'].str.lower() == emp_norm.lower()) &
            (resumo_streamlit['Tipologia'].apply(padronizar_tipologia) == tipo_norm) &
            (resumo_streamlit['Tipologia vaga'].apply(padronizar_vaga) == vaga_norm)
        ]

        if not match.empty:
            dados_unidade = match.iloc[0]
            
            if col_est and 'Unidades em estoque' in dados_unidade:
                ws.cell(row=row, column=col_est, value=float(dados_unidade['Unidades em estoque']))
            if col_area and 'Área total' in dados_unidade and pd.notna(dados_unidade['Área total']):
                ws.cell(row=row, column=col_area, value=float(dados_unidade['Área total']))
            if col_valor and 'Valor do imóvel' in dados_unidade and pd.notna(dados_unidade['Valor do imóvel']):
                ws.cell(row=row, column=col_valor, value=float(dados_unidade['Valor do imóvel']))
            if col_aval and 'Valor da avaliação' in dados_unidade and pd.notna(dados_unidade['Valor da avaliação']):
                ws.cell(row=row, column=col_aval, value=float(dados_unidade['Valor da avaliação']))
            
            atualizados += 1

    caminho_saida = caminho_performance.replace('.xlsx', '_atualizado.xlsx')
    wb.save(caminho_saida)
    print(f"\n[SUCESSO] {atualizados} linhas da concorrência foram atualizadas na aba '{mes_alvo}'.")
    print(f"[ARQUIVO GERADO] {caminho_saida}\n")
    return caminho_saida

if __name__ == '__main__':
    import sys
    print("="*60)
    print("   AUTOMAÇÃO DE DESEMPENHO E ESTOQUE DA CONCORRÊNCIA")
    print("="*60)
    
    if len(sys.argv) >= 4:
        rel_streamlit = sys.argv[1]
        perf_excel = sys.argv[2]
        mes = sys.argv[3]
    else:
        rel_streamlit = input("1. Digite o caminho do arquivo do Streamlit (ex: estoque_outubro.xlsx): ").strip()
        perf_excel = input("2. Digite o caminho da planilha de performance (ex: Performance_concorrencia.xlsx): ").strip()
        mes = input("3. Digite o nome da aba do mês (ex: Outubro): ").strip()

    if os.path.exists(rel_streamlit) and os.path.exists(perf_excel):
        resumo = carregar_relatorio_streamlit(rel_streamlit)
        atualizar_planilha_performance(perf_excel, resumo, mes)
    else:
        print("\n[ERRO] Verifique se os arquivos informados existem no diretório especificado.")
        