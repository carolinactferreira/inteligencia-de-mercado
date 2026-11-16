import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import pdfplumber
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


def parse_concatenated_string(text_str):
  s = str(text_str).strip()
  apto_match = re.search(
      r"(?:APTO|APT|UNIDADE|UP|AP)\s*[:\.-]?\s*([A-Za-z0-9]+)", s, re.IGNORECASE
  )
  apto = apto_match.group(1) if apto_match else "-"

  tipo_match = re.search(
      r"\b(\d\s*Q\s*S|\d\s*QTS\s*S|\d\s*DORM\s*S|\d\s*Q|\d\s*QTS|\d\s*DORM|\d\s*D)\b",
      s,
      re.IGNORECASE,
  )
  tipologia = tipo_match.group(1) if tipo_match else "-"

  clean_torre = re.split(
      r"\s*-\s*\d\s*Q|\s*-\s*\d\s*D|\s*-\s*APTO|\s*-\s*UNIDADE",
      s,
      flags=re.IGNORECASE,
  )[0].strip()
  bloco = clean_torre if clean_torre else s
  return bloco, apto, tipologia


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
  for col in df.columns:
    vals = df[col].dropna()
    parsed_vals = []
    for v in vals:
      try:
        vf = float(str(v).replace(",", "."))
        if 15.0 <= vf <= 400.0:
          parsed_vals.append(vf)
      except:
        pass
    if len(parsed_vals) >= len(vals) * 0.4 and len(vals) > 0:
      return df[col]
  return pd.Series([0] * len(df))


def extrair_tabelas_pdf(file_bytes):
  dados_estruturados = []
  debug_linhas = []

  try:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
      for page in pdf.pages:
        words = page.extract_words()
        if not words:
          continue

        # Descobre a posição X (horizontal) exata dos cabeçalhos "Carro" e "Moto" na página
        x_carro = None
        x_moto = None
        for w in words:
          txt_w = w["text"].lower()
          if "carr" in txt_w and w["top"] < 150:
            x_carro = w["x0"]
          elif "moto" in txt_w and w["top"] < 150:
            x_moto = w["x0"]

        # Agrupa palavras por linha
        lines_dict = {}
        for w in words:
          top_rounded = round(w["top"] / 5) * 5
          lines_dict.setdefault(top_rounded, []).append(w)

        for top_key in sorted(lines_dict.keys()):
          line_words = sorted(lines_dict[top_key], key=lambda x: x["x0"])
          texto = " ".join([w["text"] for w in line_words]).strip()

          if len(texto) < 15 or ("Regional" in texto and "Tipologia" in texto):
            continue

          debug_linhas.append(texto)

          # 1. Bloco / Torre
          torre_match = re.search(
              r"(TORRE\s*[A-Z0-9]+|BLOCO\s*[A-Z0-9]+)", texto, re.IGNORECASE
          )
          torre_val = torre_match.group(1).strip() if torre_match else "-"

          # 2. Empreendimento
          if torre_match:
            emp_bruto = texto[: torre_match.start()].strip()
            emp_bruto = re.sub(r"[-\s]+$", "", emp_bruto)
            emp_val = re.sub(
                r"^(?:[A-Za-zÀ-ÿ]+\s+){1,2}(?=[A-Z]{2,}\s+[A-Z]{2,})",
                "",
                emp_bruto,
            ).strip()
            if not emp_val:
              emp_val = emp_bruto
          else:
            emp_val = "-"

          # 3. Apto / Unidade
          apto_match = re.search(
              r"(?:APTO|UNIDADE|APT)\s*[:\.-]?\s*([A-Z0-9]+)",
              texto,
              re.IGNORECASE,
          )
          apto_val = apto_match.group(1).strip() if apto_match else "-"

          # 4. Tipologia
          tipo_match = re.search(
              r"\b(\d\s*Q(?:\s*S)?|\d\s*DORM(?:\s*S)?)\b", texto, re.IGNORECASE
          )
          tipo_val = tipo_match.group(1).strip() if tipo_match else "-"

          if torre_val == "-" and apto_val == "-" and tipo_val == "-":
            continue

          # 5. Área Interna (aceita vírgula, ponto ou espaço caso o PDF quebre ex: "39,78" ou "39.78")
          area_val = "0"
          areas = re.findall(r"\b(\d{2,3}[\.,]\d{1,2})\b", texto)
          for a in areas:
            val_f = float(a.replace(",", "."))
            if 15.0 <= val_f <= 400.0:
              area_val = a
              break

          # 6. Valores Imóvel e Avaliação
          texto_sem_versao = re.sub(r"\b\d+\.\d+\.\d+\.\d+\b", "", texto)
          valores_grandes = []
          for num_str in re.findall(r"\b\d[\d\.,]{4,}\b", texto_sem_versao):
            num_limpo = converter_numero_br(num_str)
            if num_limpo > 30000:
              valores_grandes.append(num_str)

          valor_imovel = valores_grandes[0] if len(valores_grandes) > 0 else "0"
          valor_aval = (
              valores_grandes[1] if len(valores_grandes) > 1 else valor_imovel
          )

          # 7. Vagas Carro e Moto por coordenada X exata (ou fallback pelo final da linha)
          v_carro, v_moto = "0", "0"
          if x_carro is not None and x_moto is not None:
            for w in line_words:
              if w["text"].isdigit():
                if abs(w["x0"] - x_carro) < 25:
                  v_carro = w["text"]
                elif abs(w["x0"] - x_moto) < 25:
                  v_moto = w["text"]
          else:
            # Fallback: verifica os últimos números após a versão (ex: 4.0.13.0)
            pos_versao = re.search(r"\b\d+\.\d+\.\d+\.\d+\b", texto)
            if pos_versao:
              cauda = texto[pos_versao.end() :].strip()
              nums_cauda = re.findall(r"\b\d+\b", cauda)
              # Na cauda vêm: [Oportunidades, Vagas Carro, Vagas Moto]
              if len(nums_cauda) >= 3:
                v_carro = nums_cauda[1]
                v_moto = nums_cauda[2]
              elif len(nums_cauda) == 2:
                # Se só vieram 2 números após oportunidades, verifica qual coluna tem valor
                v_carro = nums_cauda[1]

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
    st.error(f"Erro ao ler PDF: {e}")

  df_res = pd.DataFrame(
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
  return df_res, "\n".join(debug_linhas[:15])


uploaded_file = st.file_uploader(
    "Arraste sua planilha (.xlsx) ou arquivo (.pdf) aqui", type=["xlsx", "pdf"]
)

if uploaded_file is not None:
  file_bytes = uploaded_file.read()
  debug_texto = ""

  if uploaded_file.name.endswith(".pdf"):
    df, debug_texto = extrair_tabelas_pdf(file_bytes)
  else:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

  if df.empty:
    st.error("Não foi possível extrair dados do arquivo enviado.")
  else:
    with st.expander("🔍 Raio-X das linhas lidas do PDF (Diagnóstico)"):
      st.text(debug_texto)

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

        res_emp, res_bloco, res_apto, res_tipo_orig = [], [], [], []

        for idx in range(len(df_proc_source)):
          raw_e = str(s_emp.iloc[idx]).strip()
          raw_b = str(s_bloco.iloc[idx]).strip()
          raw_a = str(s_apto.iloc[idx]).strip()
          raw_t = str(s_tipo.iloc[idx]).strip()

          if ("TORRE" in raw_b.upper() or "BLOCO" in raw_b.upper()) and (
              "APTO" in raw_b.upper() or "Q" in raw_b.upper()
          ):
            ext_b, ext_a, ext_t = parse_concatenated_string(raw_b)
            final_b = ext_b if ext_b else raw_b
            final_a = (
                raw_a
                if raw_a
                not in ["-", "nan", "None", "", "(Nenhuma / Não existe)"]
                else ext_a
            )
            final_t = (
                raw_t
                if raw_t
                not in ["-", "nan", "None", "", "(Nenhuma / Não existe)"]
                else ext_t
            )
          else:
            final_b = raw_b if raw_b not in ["nan", "None", ""] else "-"
            final_a = raw_a if raw_a not in ["nan", "None", ""] else "-"
            final_t = raw_t if raw_t not in ["nan", "None", ""] else "-"

          res_emp.append(limpar_texto(raw_e))
          res_bloco.append(limpar_texto(final_b))
          res_apto.append(limpar_texto(final_a))
          res_tipo_orig.append(limpar_texto(final_t))

        df_proc = pd.DataFrame({
            "Empreendimento": res_emp,
            "Bloco": res_bloco,
            "Apto": res_apto,
            "Tipologia_Orig": res_tipo_orig,
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