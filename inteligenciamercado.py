import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import streamlit as st

st.set_page_config(page_title="Relatório de Estoque", layout="wide")

st.title("📊 Gerador de Relatório de Estoque Imobiliário")
st.write(
    "Faça o upload da planilha (.xlsx) ou arquivo PDF (.pdf) de estoque para gerar o relatório formatado."
)


def converter_numero_br(val):
  if pd.isna(val) or val is None:
    return 0.0
  s = str(val).strip().replace("R$", "").replace("$", "").replace(" ", "")
  if not s or s.lower() in ["nan", "none", "null", "<na>", "-"]:
    return 0.0
  if "," in s and "." in s:
    s = s.replace(".", "").replace(",", ".")
  elif "," in s:
    s = s.replace(",", ".")
  try:
    return float(re.sub(r"[^0-9.]", "", s))
  except:
    return 0.0


def limpar_texto(val, padrao="-"):
  if pd.isna(val) or val is None:
    return padrao
  s = str(val).strip()
  if not s or s.lower() in [
      "nan",
      "none",
      "null",
      "<na>",
      "(nenhuma / não existe)",
  ]:
    return padrao
  return s


def mapear_tipologia(t):
  t_str = str(t).strip().upper()
  if not t_str or t_str in ["-", "NAN", "NONE"]:
    return "OUTRA TIPOLOGIA"
  has_suite = (
      any(term in t_str for term in ["SUITE", "SUÍTE", " S", "S "])
      or t_str.endswith("S")
  )
  if "3" in t_str:
    return "3 Q S" if has_suite else "3 Q"
  elif "2" in t_str:
    return "2 Q S" if has_suite else "2 Q"
  elif "1" in t_str:
    return "1 Q S" if has_suite else "1 Q"
  return "OUTRA TIPOLOGIA"


def categorizar_vaga(row, col_moto="", col_carro=""):
  moto_val = str(row.get(col_moto, "")).strip().upper() if col_moto in row else ""
  carro_val = (
      str(row.get(col_carro, "")).strip().upper() if col_carro in row else ""
  )
  if moto_val in ["1", "SIM", "S", "MOTO", "TRUE", "1.0"]:
    return "Vaga Moto"
  elif carro_val in ["1", "SIM", "S", "CARRO", "TRUE", "1.0"]:
    return "Vaga Carro"
  if converter_numero_br(moto_val) > 0:
    return "Vaga Moto"
  if converter_numero_br(carro_val) > 0:
    return "Vaga Carro"
  return "Sem Vaga"


def auto_detect_columns(df):
  mapping = {
      "empreendimento": None,
      "bloco": None,
      "apto": None,
      "tipologia": None,
      "area": None,
      "valor_imovel": None,
      "valor_avaliacao": None,
      "carro": None,
      "moto": None,
  }
  for col in df.columns:
    c_lower = str(col).lower()
    if (
        any(
            k in c_lower
            for k in ["empreend", "projeto", "condominio", "condomínio", "obra"]
        )
        and not mapping["empreendimento"]
    ):
      mapping["empreendimento"] = col
    elif any(k in c_lower for k in ["bloco", "torre"]) and not mapping["bloco"]:
      mapping["bloco"] = col
    elif (
        any(k in c_lower for k in ["apto", "unidade", "apartamento"])
        and not mapping["apto"]
    ):
      mapping["apto"] = col
    elif (
        any(k in c_lower for k in ["tipologia", "tipo", "planta", "dorm"])
        and not mapping["tipologia"]
    ):
      mapping["tipologia"] = col
    elif (
        any(k in c_lower for k in ["area", "área", "m2", "privativa"])
        and not mapping["area"]
    ):
      mapping["area"] = col
    elif (
        any(
            k in c_lower
            for k in ["venda", "imovel", "imóvel", "preço", "preco"]
        )
        and "aval" not in c_lower
        and not mapping["valor_imovel"]
    ):
      mapping["valor_imovel"] = col
    elif any(k in c_lower for k in ["aval"]) and not mapping["valor_avaliacao"]:
      mapping["valor_avaliacao"] = col
    elif "carro" in c_lower and not mapping["carro"]:
      mapping["carro"] = col
    elif "moto" in c_lower and not mapping["moto"]:
      mapping["moto"] = col
  return mapping


def get_area_series(df, sel_area):
  if sel_area in df.columns and sel_area != "(Nenhuma / Não existe)":
    return df[sel_area]
  return pd.Series([0] * len(df))


def extrair_tabelas_pdf(file_bytes):
  dados_estruturados = []

  try:
    images = convert_from_bytes(file_bytes, dpi=250)

    for img in images:
      width, _ = img.size
      ocr_data = pytesseract.image_to_data(
          img, lang="por", config="--psm 6", output_type=pytesseract.Output.DICT
      )

      # Agrupa palavras pela coordenada Y (altura da linha)
      words_list = []
      for i in range(len(ocr_data["text"])):
        txt = ocr_data["text"][i].strip()
        if txt:
          words_list.append({
              "text": txt,
              "x0": ocr_data["left"][i],
              "top": ocr_data["top"][i],
          })

      words_list.sort(key=lambda w: w["top"])
      linhas_agrupadas = []
      linha_atual = []
      top_atual = None

      for w in words_list:
        if top_atual is None:
          top_atual = w["top"]
          linha_atual.append(w)
        elif abs(w["top"] - top_atual) <= 14:
          linha_atual.append(w)
        else:
          linhas_agrupadas.append(linha_atual)
          linha_atual = [w]
          top_atual = w["top"]
      if linha_atual:
        linhas_agrupadas.append(linha_atual)

      for line_words in linhas_agrupadas:
        line_words.sort(key=lambda w: w["x0"])
        texto = " ".join([w["text"] for w in line_words])

        # Normaliza caracteres que o OCR confunde (ex: 'Ο' grego por 'O', 'ΑΡΤΟ' por 'APTO')
        texto_norm = (
            texto.replace("Ο", "O")
            .replace("ο", "o")
            .replace("Α", "A")
            .replace("Ρ", "P")
            .replace("Τ", "T")
        )

        if "Regional" in texto_norm and "Tipologia" in texto_norm:
          continue

        # 1. Bloco / Torre
        torre_match = re.search(
            r"(TORRE\s*[A-Z0-9]+|BLOCO\s*[A-Z0-9]+)", texto_norm, re.IGNORECASE
        )
        if not torre_match:
          continue
        torre_val = torre_match.group(1).strip()

        # 2. Empreendimento (limpa "Campinas Campinas" antes do nome em maiúsculas)
        emp_bruto = texto_norm[: torre_match.start()].strip(" -")
        maiusculas = re.search(
            r"([A-ZÀ-Ú]{2,}(?:\s+[A-ZÀ-Ú0-9]{2,})+)", emp_bruto
        )
        emp_val = maiusculas.group(1).strip() if maiusculas else emp_bruto

        # 3. Apto / Unidade
        apto_match = re.search(
            r"(?:APTO|UNIDADE|APT)\s*[:\.-]?\s*([A-Z0-9]+)",
            texto_norm,
            re.IGNORECASE,
        )
        apto_val = apto_match.group(1).strip() if apto_match else "-"

        # 4. Tipologia (corrige "20" lido pelo OCR no lugar de "2Q" e trata "2QR" e "2QS")
        tipo_val = "2 Q"
        tipo_match = re.search(
            r"-\s*(\d\s*(?:Q\s*S|Q\s*R|Q|0|DORM))\s*-",
            texto_norm,
            re.IGNORECASE,
        )
        if tipo_match:
          t_raw = tipo_match.group(1).upper().replace(" ", "")
          if t_raw in ["20", "2Q", "2QR"]:
            tipo_val = "2 Q"
          elif t_raw == "2QS":
            tipo_val = "2 Q S"
          elif t_raw in ["10", "1Q"]:
            tipo_val = "1 Q"
          elif t_raw in ["30", "3Q"]:
            tipo_val = "3 Q"
          else:
            tipo_val = tipo_match.group(1).strip()
        elif "2QS" in texto_norm.replace(" ", "").upper():
          tipo_val = "2 Q S"

        # 5. Área Interna (39,78 / 36,66 / 40,6 / 41,34 / 43,76 / 44,51)
        area_val = "0"
        for a in re.findall(r"\b(\d{2,3}[\.,]\d{1,2})\b", texto_norm):
          val_f = float(a.replace(",", "."))
          if 15.0 <= val_f <= 400.0:
            area_val = a.replace(".", ",")
            break

        # 6. Valores Imóvel e Avaliação (números inteiros > 30.000)
        texto_sem_versao = re.sub(r"\b\d+\.\d+\.\d+\.\d+\b", "", texto_norm)
        valores_grandes = []
        for num_str in re.findall(r"\b\d{5,7}\b", texto_sem_versao):
          if int(num_str) > 30000:
            valores_grandes.append(num_str)

        valor_imovel = valores_grandes[0] if len(valores_grandes) > 0 else "0"
        valor_aval = (
            valores_grandes[1] if len(valores_grandes) > 1 else valor_imovel
        )

        # 7. Vagas Carro e Moto: analisa tudo que aparece fisicamente à direita da coluna "Versão da tabela"
        v_carro, v_moto = "0", "0"
        x_versao = 0
        for w in line_words:
          if re.search(r"\d+\.\d+\.\d+\.\d+", w["text"]):
            x_versao = w["x0"]
            break

        if x_versao > 0:
          digitos_vagas = []
          for w in line_words:
            txt_limpo = (
                w["text"]
                .replace("。", "0")
                .replace("O", "0")
                .replace("o", "0")
                .replace("I", "1")
                .replace("l", "1")
            )
            if txt_limpo in ["0", "1", "2"] and w["x0"] > (x_versao + 70):
              digitos_vagas.append((w["x0"], txt_limpo))

          if len(digitos_vagas) >= 2:
            v_carro = digitos_vagas[-2][1]
            v_moto = digitos_vagas[-1][1]
          elif len(digitos_vagas) == 1:
            if digitos_vagas[0][0] >= (width * 0.94):
              v_moto = digitos_vagas[0][1]
            else:
              v_carro = digitos_vagas[0][1]

        dados_estruturados.append([
            emp_val,
            torre_val,
            apto_val,
            tipo_val,
            area_val,
            valor_imovel,
            valor_aval,
            v_carro,
            v_moto,
        ])
  except Exception as e:
    st.error(f"Erro ao processar PDF: {e}")

  return pd.DataFrame(
      dados_estruturados,
      columns=[
          "Empreendimento",
          "Bloco/Torre",
          "Apto/Unidade",
          "Tipologia",
          "Área",
          "Valor Imóvel",
          "Valor Avaliação",
          "Carro",
          "Moto",
      ],
  )


uploaded_file = st.file_uploader(
    "Arraste sua planilha (.xlsx) ou arquivo (.pdf) aqui", type=["xlsx", "pdf"]
)

if uploaded_file is not None:
  file_bytes = uploaded_file.read()

  if uploaded_file.name.endswith(".pdf"):
    with st.spinner("Lendo tabelas do PDF..."):
      df = extrair_tabelas_pdf(file_bytes)
  else:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

  if df.empty:
    st.error("Não foi possível extrair dados do arquivo enviado.")
  else:
    st.subheader("📋 Dados Brutos Identificados no Arquivo")
    st.write(f"Total de registros encontrados: **{len(df)} linhas**")
    st.dataframe(df.head(15), use_container_width=True)

    st.subheader("⚙️ Mapeamento de Colunas")
    colunas_disponiveis = ["(Nenhuma / Não existe)"] + list(df.columns)
    mapping_auto = auto_detect_columns(df)

    def idx_of(val):
      return colunas_disponiveis.index(val) if val in colunas_disponiveis else 0

    col01, col02 = st.columns(2)
    with col01:
      sel_emp = st.selectbox(
          "Coluna do Empreendimento:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["empreendimento"]),
      )
    with col02:
      sel_bloco = st.selectbox(
          "Coluna do Bloco/Torre:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["bloco"]),
      )

    col1, col2, col3 = st.columns(3)
    with col1:
      sel_apto = st.selectbox(
          "Coluna do Apto/Unidade:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["apto"]),
      )
      sel_tipo = st.selectbox(
          "Coluna da Tipologia/Planta:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["tipologia"]),
      )
    with col2:
      sel_area = st.selectbox(
          "Coluna da Área Total/Privativa:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["area"]),
      )
      sel_valor = st.selectbox(
          "Coluna do Valor do Imóvel:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["valor_imovel"]),
      )
    with col3:
      sel_avaliacao = st.selectbox(
          "Coluna do Valor da Avaliação:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["valor_avaliacao"]),
      )
      sel_carro = st.selectbox(
          "Coluna Vaga Carro:",
          colunas_disponiveis,
          index=idx_of(mapping_auto["carro"]),
      )

    sel_moto = st.selectbox(
        "Coluna Vaga Moto:",
        colunas_disponiveis,
        index=idx_of(mapping_auto["moto"]),
    )

    st.markdown("---")

    lista_emps = []
    if sel_emp != "(Nenhuma / Não existe)":
      lista_emps = [
          str(x).strip()
          for x in df[sel_emp].dropna().unique()
          if str(x).strip() and str(x).lower() != "nan"
      ]

    if lista_emps:
      st.subheader("🏢 Seleção de Empreendimento")
      filtro_emp = st.multiselect(
          "Selecione os empreendimentos para incluir no relatório:",
          options=lista_emps,
          default=lista_emps,
      )
    else:
      filtro_emp = []

    if st.button("🚀 Gerar Relatório Formatado em Excel", type="primary"):
      with st.spinner("Processando dados e formatando Excel..."):

        if lista_emps and filtro_emp:
          df_proc_source = df[
              df[sel_emp].astype(str).str.strip().isin(filtro_emp)
          ].copy()
          nome_relatorio = (
              filtro_emp[0]
              if len(filtro_emp) == 1
              else "Varios_Empreendimentos"
          )
        else:
          df_proc_source = df.copy()
          nome_relatorio = "Relatorio_Geral"

        s_emp = (
            df_proc_source[sel_emp]
            if sel_emp in df_proc_source.columns
            else pd.Series(["-"] * len(df_proc_source))
        )
        s_bloco = (
            df_proc_source[sel_bloco]
            if sel_bloco in df_proc_source.columns
            else pd.Series(["-"] * len(df_proc_source))
        )
        s_apto = (
            df_proc_source[sel_apto]
            if sel_apto in df_proc_source.columns
            else pd.Series(["-"] * len(df_proc_source))
        )
        s_tipo = (
            df_proc_source[sel_tipo]
            if sel_tipo in df_proc_source.columns
            else pd.Series(["-"] * len(df_proc_source))
        )
        s_area = get_area_series(df_proc_source, sel_area)
        s_valor = (
            df_proc_source[sel_valor]
            if sel_valor in df_proc_source.columns
            else pd.Series([0] * len(df_proc_source))
        )
        s_aval = (
            df_proc_source[sel_avaliacao]
            if sel_avaliacao in df_proc_source.columns
            else pd.Series([0] * len(df_proc_source))
        )

        df_proc = pd.DataFrame({
            "Empreendimento": [limpar_texto(v) for v in s_emp],
            "Bloco": [limpar_texto(v) for v in s_bloco],
            "Apto": [limpar_texto(v) for v in s_apto],
            "Tipologia_Orig": [limpar_texto(v) for v in s_tipo],
            "Área": [converter_numero_br(v) for v in s_area],
            "Valor_Imovel": [converter_numero_br(v) for v in s_valor],
            "Valor_Avaliacao": [converter_numero_br(v) for v in s_aval],
            "Categoria_Vaga": [
                categorizar_vaga(row, col_moto=sel_moto, col_carro=sel_carro)
                for _, row in df_proc_source.iterrows()
            ],
        })
        df_proc["Tipologia_Clean"] = df_proc["Tipologia_Orig"].apply(
            mapear_tipologia
        )

        def contem_numero(texto):
          return any(char.isdigit() for char in str(texto))

        df_proc = df_proc[
            (df_proc["Valor_Imovel"] > 0)
            | (df_proc["Área"] > 0)
            | (df_proc["Apto"].apply(contem_numero))
            | (df_proc["Bloco"].apply(contem_numero))
        ]

        output = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        data_hoje = datetime.datetime.now().strftime("%d/%m")
        fill_title = PatternFill(
            start_color="1B365D", end_color="1B365D", fill_type="solid"
        )
        font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        font_subtitle = Font(
            name="Calibri", size=11, italic=True, color="FFFFFF"
        )
        fill_kpi_header = PatternFill(
            start_color="F2F4F8", end_color="F2F4F8", fill_type="solid"
        )
        font_kpi_header = Font(
            name="Calibri", size=10, bold=True, color="595959"
        )
        fill_kpi_val = PatternFill(
            start_color="F8F9FA", end_color="F8F9FA", fill_type="solid"
        )
        font_kpi_val = Font(name="Calibri", size=16, bold=True, color="1B365D")
        font_kpi_sub = Font(name="Calibri", size=9, italic=True, color="808080")
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        headers = [
            "Empreendimento",
            "Bloco/Torre",
            "Apto/Unidade",
            "Tipologia Original",
            "Área Total (m²)",
            "Valor Imóvel (R$)",
            "Valor Avaliação (R$)",
        ]

        def aplicar_kpis(ws, titulo, df_tab):
          total_unid = len(df_tab)
          area_media = df_tab["Área"].mean() if total_unid > 0 else 0
          val_imovel = df_tab["Valor_Imovel"].mean() if total_unid > 0 else 0
          val_aval = df_tab["Valor_Avaliacao"].mean() if total_unid > 0 else 0

          ws.merge_cells("A1:G1")
          ws["A1"] = titulo
          ws["A1"].font = font_title
          ws["A1"].fill = fill_title
          ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

          ws.merge_cells("A2:G2")
          ws["A2"] = f"Resumo Automático de Estoque e Vagas (Base: {data_hoje})"
          ws["A2"].font = font_subtitle
          ws["A2"].fill = fill_title
          ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

          ws.merge_cells("A4:B4")
          ws["A4"] = "TOTAL DE UNIDADES"
          ws.merge_cells("C4:D4")
          ws["C4"] = "ÁREA MÉDIA (TOTAL)"
          ws["E4"] = "VALOR MÉDIO DO IMÓVEL"
          ws.merge_cells("F4:G4")
          ws["F4"] = "VALOR MÉDIO DE AVALIAÇÃO"

          for c in ["A4", "B4", "C4", "D4", "E4", "F4", "G4"]:
            ws[c].fill = fill_kpi_header
            ws[c].font = font_kpi_header
            ws[c].alignment = Alignment(horizontal="center", vertical="center")
            ws[c].border = border_thin

          ws.merge_cells("A5:B5")
          ws["A5"] = total_unid
          ws.merge_cells("C5:D5")
          ws["C5"] = (
              f"{area_media:,.2f} m²".replace(",", "X")
              .replace(".", ",")
              .replace("X", ".")
          )
          ws["E5"] = (
              f"R$ {val_imovel:,.2f}".replace(",", "X")
              .replace(".", ",")
              .replace("X", ".")
          )
          ws.merge_cells("F5:G5")
          ws["F5"] = (
              f"R$ {val_aval:,.2f}".replace(",", "X")
              .replace(".", ",")
              .replace("X", ".")
          )

          for c in ["A5", "B5", "C5", "D5", "E5", "F5", "G5"]:
            ws[c].fill = fill_kpi_val
            ws[c].font = font_kpi_val
            ws[c].alignment = Alignment(horizontal="center", vertical="center")
            ws[c].border = border_thin

          ws.merge_cells("A6:B6")
          ws["A6"] = "unidades no estoque"
          ws.merge_cells("C6:D6")
          ws["C6"] = "Metragem média calculada"
          ws["E6"] = "Valor médio de tabela"
          ws.merge_cells("F6:G6")
          ws["F6"] = "Valor médio de avaliação"

          for c in ["A6", "B6", "C6", "D6", "E6", "F6", "G6"]:
            ws[c].fill = fill_kpi_val
            ws[c].font = font_kpi_sub
            ws[c].alignment = Alignment(horizontal="center", vertical="center")
            ws[c].border = border_thin

          ws.append([])
          ws.append(headers)

          for _, row in df_tab.iterrows():
            ws.append([
                row["Empreendimento"],
                row["Bloco"],
                row["Apto"],
                row["Tipologia_Orig"],
                row["Área"],
                row["Valor_Imovel"],
                row["Valor_Avaliacao"],
            ])

        ws_geral = wb.create_sheet(title="Estoque Geral")
        aplicar_kpis(
            ws_geral, f"RELATÓRIO DE ESTOQUE — {nome_relatorio.upper()}", df_proc
        )

        configuracoes = [
            ("Moto 1Q", "Vaga Moto", "1 Q"),
            ("Carro 1Q", "Vaga Carro", "1 Q"),
            ("Sem Vaga 1Q", "Sem Vaga", "1 Q"),
            ("Moto 2QS", "Vaga Moto", "2 Q S"),
            ("Carro 2QS", "Vaga Carro", "2 Q S"),
            ("Sem Vaga 2QS", "Sem Vaga", "2 Q S"),
            ("Moto 2Q", "Vaga Moto", "2 Q"),
            ("Carro 2Q", "Vaga Carro", "2 Q"),
            ("Sem Vaga 2Q", "Sem Vaga", "2 Q"),
            ("Outras Unidades", "TODAS", "OUTRA TIPOLOGIA"),
        ]

        for sheet_name, cat_key, tipo_key in configuracoes:
          if cat_key == "TODAS":
            sub_df = df_proc[df_proc["Tipologia_Clean"] == tipo_key]
          else:
            sub_df = df_proc[
                (df_proc["Categoria_Vaga"] == cat_key)
                & (df_proc["Tipologia_Clean"] == tipo_key)
            ]

          if sub_df.empty:
            continue

          ws = wb.create_sheet(title=sheet_name)
          aplicar_kpis(
              ws,
              f"{nome_relatorio.upper()} — {cat_key.upper()} ({tipo_key})",
              sub_df,
          )

        for sheet in wb.worksheets:
          for row in sheet.iter_rows(min_row=9, max_col=7):
            if len(row) >= 7:
              row[4].number_format = "#,##0.00"
              row[5].number_format = "R$ #,##0.00"
              row[6].number_format = "R$ #,##0.00"
          for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 15)

        wb.save(output)
        processed_data = output.getvalue()

        st.success("Relatório gerado com sucesso!")

        st.download_button(
            label="📥 Baixar Relatório Excel Formatado",
            data=processed_data,
            file_name=f"Relatorio_{nome_relatorio.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )