import csv
import re
import unicodedata
import warnings
import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Auditoría de Faltantes - Popayán", page_icon="🥤", layout="wide"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# --- FUNCIONES DE LIMPIEZA Y FORMATO ---
def remover_tildes_y_raros(texto):
  if pd.isna(texto):
    return ""
  txt = str(texto).upper().strip()
  txt = unicodedata.normalize("NFD", txt)
  txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
  txt = re.sub(r"[^A-Z0-9]", " ", txt)
  return re.sub(r"\s+", " ", txt).strip()


def normalizar_sap(val):
  if pd.isna(val):
    return "1"
  txt = str(val).strip().split(".")[0]
  txt = re.sub(r"\D", "", txt)
  if txt in ["1", ""] or not txt:
    return "1"
  return txt


def limpiar_monto(val):
  if pd.isna(val):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  val_str = str(val).replace("$", "").replace("-", "0").strip()
  if not val_str or val_str.lower() in ["none", "nan"]:
    return 0.0
  val_str = val_str.replace(".", "").replace(",", ".")
  try:
    return float(val_str)
  except:
    return 0.0


def convertir_a_fecha(val):
  if pd.isna(val):
    return pd.NaT
  dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
  return dt


def formatear_fecha_ui(val_dt):
  if pd.isna(val_dt):
    return "Sin Fecha"
  return val_dt.strftime("%d/%m/%Y")


PALABRAS_PROHIBIDAS = [
    "POSTOBON POPAYAN",
    "POSTOBON",
    "POPAYAN",
    "TESORERIA",
    "CARTERA",
    "VENTAS",
    "LOGISTICA",
    "SI",
    "NO",
    "NONE",
    "NAN",
    "RESPONSABLE / ZONA",
    "PASE A CARTERA",
    "PENDIENTE",
    "PAGO POR VERIFICAR",
]


# --- CARGAR HISTÓRICO CSV MAESTRO ---
def cargar_historico_csv(file_obj):
  if file_obj is None:
    return pd.DataFrame()
  try:
    rows = []
    content = file_obj.getvalue().decode("latin1", errors="ignore")
    reader = csv.reader(content.splitlines(), delimiter=";")
    for r in reader:
      if any(str(cell).strip() for cell in r):
        rows.append(r)

    if not rows:
      return pd.DataFrame()

    header_idx = -1
    for idx, r in enumerate(rows):
      row_str = " ".join([str(c).upper() for c in r])
      if (
          "IMPORTE" in row_str
          or "SALDO" in row_str
          or "TITULAR" in row_str
          or "FECHA" in row_str
      ):
        header_idx = idx
        break

    if header_idx == -1:
      header_idx = 0

    header_row = rows[header_idx]
    data_rows = rows[header_idx + 1 :]

    col_fecha_i, col_valor_i, col_sap_i, col_titular_i = 3, 4, 5, 6
    for i, h in enumerate(header_row):
      h_str = str(h).upper()
      if "FECHA" in h_str:
        col_fecha_i = i
      elif "IMPORTE" in h_str or "VALOR" in h_str:
        col_valor_i = i
      elif "SAP" in h_str or "DUEDOR" in h_str:
        col_sap_i = i
      elif "TITULAR" in h_str or "ZONA" in h_str:
        col_titular_i = i

    parsed_rows = []
    for r in data_rows:
      if len(r) <= max(col_fecha_i, col_valor_i, col_sap_i):
        continue
      f_dt = convertir_a_fecha(
          r[col_fecha_i] if col_fecha_i < len(r) else ""
      )
      v_num = limpiar_monto(r[col_valor_i] if col_valor_i < len(r) else "")
      s_code = normalizar_sap(r[col_sap_i] if col_sap_i < len(r) else "")
      t_name = remover_tildes_y_raros(
          r[col_titular_i] if col_titular_i < len(r) else ""
      )
      if not t_name or t_name in PALABRAS_PROHIBIDAS:
        t_name = "RESPONSABLE / ZONA"

      if v_num > 0 and pd.notna(f_dt) and f_dt.year >= 2026:
        parsed_rows.append({
            "Nombre_Deudor_OK": t_name,
            "Deudor_SAP_OK": s_code,
            "Valor_Faltante_Num": v_num,
            "Abonos_Num": 0.0,
            "Saldo_Num": v_num,
            "Fecha_DT": f_dt,
            "Fecha_UI": formatear_fecha_ui(f_dt),
            "Motivo_OK": "Histórico Maestro",
        })
    return pd.DataFrame(parsed_rows)
  except Exception as e:
    return pd.DataFrame()


# --- CARGAR ANEXO EXCEL POR POSICIONES EXACTAS (POPAYÁN) ---
def cargar_anexo_excel(file_obj):
  if file_obj is None:
    return pd.DataFrame()
  try:
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    sheet = wb.active
    df_raw = pd.DataFrame(list(sheet.values))
    df_raw = df_raw.dropna(how="all").reset_index(drop=True)

    # Identificamos la fila de cabecera donde está "DEUDOR SAP" o "VALOR"
    header_idx = 5  # Por defecto fila 6 (índice 5)
    for idx, row in df_raw.iterrows():
      row_str = " ".join(
          [str(c).upper().replace("\n", " ").strip() for c in row if pd.notna(c)]
      )
      if "DEUDOR SAP" in row_str or "VALOR DEL FALTANTE" in row_str:
        header_idx = idx
        break

    # Tomamos los datos a partir de la siguiente fila
    df = df_raw.iloc[header_idx + 1 :].copy()

    # POSICIONES EXACTAS SEGÚN TU ESTRUCTURA DE POPAYÁN:
    col_sap_idx = 2  # Columna C: DEUDOR SAP
    col_fecha_idx = 3  # Columna D: FECHA EN QUE GENERA
    col_valor_idx = 6  # Columna G: VALOR DEL FALTANTE
    col_abono_idx = 7  # Columna H: ABONOS REALIZADOS
    col_saldo_idx = 8  # Columna I: SALDO
    col_responsable_idx = 10  # Columna K: NOMBRE DEL RESPONSABLE
    col_motivo_idx = 11  # Columna L: MOTIVOS

    nombres_finales, saps_finales, mot_finales = [], [], []
    importes_finales, abonos_finales, saldos_finales = [], [], []
    fechas_finales = []

    for _, row in df.iterrows():
      # Extraemos los valores de forma segura por índice de columna
      val_imp = (
          limpiar_monto(row.iloc[col_valor_idx])
          if col_valor_idx < len(row)
          else 0.0
      )
      val_abo = (
          limpiar_monto(row.iloc[col_abono_idx])
          if col_abono_idx < len(row)
          else 0.0
      )
      val_sal = (
          limpiar_monto(row.iloc[col_saldo_idx])
          if col_saldo_idx < len(row)
          else 0.0
      )

      # Si el valor inicial viene vacío pero hay saldo o abono, lo reconstruimos
      if val_imp <= 0:
        if val_sal > 0 or val_abo > 0:
          val_imp = val_sal + val_abo
        else:
          continue  # Fila vacía o sin valor relevante

      c_code = (
          normalizar_sap(row.iloc[col_sap_idx])
          if col_sap_idx < len(row) and pd.notna(row.iloc[col_sap_idx])
          else "1"
      )
      motivo = (
          str(row.iloc[col_motivo_idx]).strip()
          if col_motivo_idx < len(row) and pd.notna(row.iloc[col_motivo_idx])
          else ""
      )

      if col_responsable_idx < len(row) and pd.notna(
          row.iloc[col_responsable_idx]
      ):
        txt = remover_tildes_y_raros(row.iloc[col_responsable_idx])
        nombre_ok = (
            txt
            if txt and txt not in PALABRAS_PROHIBIDAS and not txt.isdigit()
            else "RESPONSABLE / ZONA"
        )
      else:
        nombre_ok = "RESPONSABLE / ZONA"

      f_val = (
          row.iloc[col_fecha_idx] if col_fecha_idx < len(row) else pd.NaT
      )

      # Cuadre contable de saldos
      if val_sal == 0.0 and val_imp > 0 and val_abo > 0:
        val_sal = max(0.0, val_imp - val_abo)
      elif val_abo == 0.0 and val_imp > 0 and val_sal > 0:
        val_abo = max(0.0, val_imp - val_sal)
      elif val_imp == 0.0 and val_sal > 0:
        val_imp = val_sal + val_abo

      nombres_finales.append(nombre_ok)
      saps_finales.append(c_code)
      mot_finales.append(motivo)
      importes_finales.append(val_imp)
      abonos_finales.append(val_abo)
      saldos_finales.append(val_sal)
      fechas_finales.append(f_val)

    res_df = pd.DataFrame()
    res_df["Nombre_Deudor_OK"] = nombres_finales
    res_df["Deudor_SAP_OK"] = saps_finales
    res_df["Valor_Faltante_Num"] = importes_finales
    res_df["Abonos_Num"] = abonos_finales
    res_df["Saldo_Num"] = saldos_finales
    res_df["Fecha_DT"] = [convertir_a_fecha(f) for f in fechas_finales]
    res_df["Fecha_UI"] = [formatear_fecha_ui(d) for d in res_df["Fecha_DT"]]
    res_df["Motivo_OK"] = mot_finales

    return res_df[res_df["Valor_Faltante_Num"] > 0].copy()
  except Exception as e:
    st.error(f"Error leyendo el anexo: {e}")
    return pd.DataFrame()


# --- MOTOR DE AUDITORÍA ROBUSTO ---
def procesar_archivos(file_hist, file_anexo):
  if file_anexo is None:
    return (
        "⚠️ Por favor sube el Anexo del día para realizar la auditoría.",
        None,
        None,
        pd.DataFrame(),
        pd.DataFrame(),
        pd.DataFrame(),
        pd.DataFrame(),
    )

  df_anexo = cargar_anexo_excel(file_anexo)
  df_hist = (
      cargar_historico_csv(file_hist) if file_hist is not None else pd.DataFrame()
  )

  if df_anexo.empty:
    return (
        "⚠️ El anexo cargado no tiene registros válidos.",
        None,
        None,
        pd.DataFrame(),
        pd.DataFrame(),
        pd.DataFrame(),
        pd.DataFrame(),
    )

  fechas_validas = df_anexo["Fecha_DT"].dropna()
  fecha_dia_actual = (
      fechas_validas.max()
      if not fechas_validas.empty
      else pd.Timestamp("2026-10-31")
  )

  # 1. Nuevos Faltantes (Día Actual exacto con saldo > 0)
  es_dia_actual = df_anexo["Fecha_DT"].dt.date == fecha_dia_actual.date()
  nuevos = df_anexo[es_dia_actual & (df_anexo["Saldo_Num"] > 0)].copy()

  # 2. Pendientes Activos (Días anteriores con saldo > 0)
  pendientes = df_anexo[
      (~es_dia_actual) & (df_anexo["Saldo_Num"] > 0)
  ].copy()

  # 3. Faltantes Saldados / Pagados (Saldo <= 0)
  saldados = df_anexo[df_anexo["Saldo_Num"] <= 0].copy()

  # 4. Omitidos en Anexo (Comparación exacta de montos con el histórico)
  omitidos = pd.DataFrame()
  if not df_hist.empty:
    pendientes_hist = df_hist[df_hist["Saldo_Num"] > 0].copy()
    montos_presentes_hoy = set(df_anexo["Valor_Faltante_Num"].round(2))

    omitidos_list = []
    for _, row in pendientes_hist.iterrows():
      monto_hist = round(float(row["Valor_Faltante_Num"]), 2)
      if monto_hist not in montos_presentes_hoy and row["Saldo_Num"] > 0:
        omitidos_list.append(row)

    if omitidos_list:
      omitidos = pd.DataFrame(omitidos_list).drop_duplicates(
          subset=["Deudor_SAP_OK", "Valor_Faltante_Num"]
      )

  def preparar_df_ui(df_in):
    if df_in is None or df_in.empty:
      return pd.DataFrame(
          columns=[
              "Nombre Deudor",
              "SAP",
              "Fecha",
              "Valor Faltante Inicial ($)",
              "Abonos Realizados ($)",
              "Saldo Pendiente Actual ($)",
              "Motivo",
          ]
      )
    out = pd.DataFrame()
    out["Nombre Deudor"] = df_in["Nombre_Deudor_OK"]
    out["SAP"] = df_in["Deudor_SAP_OK"]
    out["Fecha"] = df_in["Fecha_UI"]
    out["Valor Faltante Inicial ($)"] = df_in["Valor_Faltante_Num"].apply(
        lambda x: f"${int(round(limpiar_monto(x))):,}"
    )
    out["Abonos Realizados ($)"] = df_in["Abonos_Num"].apply(
        lambda x: f"${int(round(limpiar_monto(x))):,}"
    )
    out["Saldo Pendiente Actual ($)"] = df_in["Saldo_Num"].apply(
        lambda x: f"${int(round(limpiar_monto(x))):,}"
    )
    out["Motivo"] = df_in["Motivo_OK"] if "Motivo_OK" in df_in.columns else ""
    return out

  df_nuevos_ui = preparar_df_ui(nuevos)
  df_saldados_ui = preparar_df_ui(saldados)
  df_pendientes_ui = preparar_df_ui(pendientes)
  df_omitidos_ui = preparar_df_ui(omitidos)

  tot_nuevos = nuevos["Saldo_Num"].sum() if not nuevos.empty else 0
  tot_saldados = saldados["Abonos_Num"].sum() if not saldados.empty else 0
  tot_pendientes = pendientes["Saldo_Num"].sum() if not pendientes.empty else 0
  tot_omitidos = omitidos["Saldo_Num"].sum() if not omitidos.empty else 0

  m_str = f"## 🥤 **AUDITORÍA DE FALTANTES - POPAYÁN**\n\n"
  m_str += f"- 🔴 **Nuevos Faltantes (Día Actual - {formatear_fecha_ui(fecha_dia_actual)}):** {len(nuevos)} registro(s) | **${int(tot_nuevos):,}**\n"
  m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} registro(s) | **${int(tot_saldados):,}**\n"
  m_str += f"- 🟣 **Pendientes Activos (Días Anteriores):** {len(pendientes)} registro(s) | **${int(tot_pendientes):,}**\n"
  m_str += f"- ⚠️ **Omitidos en Anexo:** {len(omitidos)} registro(s) | **${int(tot_omitidos):,}**\n"

  categorias = [
      "Nuevos Faltantes",
      "Faltantes Saldados",
      "Pendientes Activos",
      "Omitidos en Anexo",
  ]
  cantidades = [len(nuevos), len(saldados), len(pendientes), len(omitidos)]
  montos = [tot_nuevos, tot_saldados, tot_pendientes, tot_omitidos]
  colores = {
      "Nuevos Faltantes": "#EF553B",
      "Faltantes Saldados": "#00CC96",
      "Pendientes Activos": "#AB63FA",
      "Omitidos en Anexo": "#FFA15A",
  }

  df_summary = pd.DataFrame(
      {"Categoría": categorias, "Cantidad": cantidades, "Monto": montos}
  )

  fig_pie = px.pie(
      df_summary,
      values="Cantidad",
      names="Categoría",
      title="<b>Distribución por Cantidad de Casos</b>",
      color="Categoría",
      color_discrete_map=colores,
      hole=0.4,
  )
  fig_bar = px.bar(
      df_summary,
      x="Categoría",
      y="Monto",
      text_auto=".0f",
      title="<b>Impacto Financiero Real ($ COP)</b>",
      color="Categoría",
      color_discrete_map=colores,
  )
  fig_bar.update_layout(showlegend=False, yaxis_title="Monto ($)")

  return (
      m_str,
      fig_pie,
      fig_bar,
      df_nuevos_ui,
      df_saldados_ui,
      df_pendientes_ui,
      df_omitidos_ui,
  )


# --- INTERFAZ EN STREAMLIT ---
st.markdown(
    "# 🥤 **POSTOBÓN POPAYÁN - SISTEMA DE AUDITORÍA DE FALTANTES**"
)

col1, col2 = st.columns([1, 1])

with col1:
  st.subheader("1. Carga de Archivos")
  file_hist = st.file_uploader("Histórico Maestro (.csv)", type=["csv"])
  file_anexo = st.file_uploader(
      "Anexo del Día Actual (.xlsx / .xlsm)", type=["xlsx", "xlsm", "csv"]
  )

with col2:
  st.subheader("Estado del Proceso")
  ejecutar = st.button("🚀 EJECUTAR AUDITORÍA", type="primary")

if ejecutar:
  if file_anexo is None:
    st.warning("⚠️ Por favor sube el Anexo del día para realizar la auditoría.")
  else:
    with st.spinner("Procesando auditoría..."):
      (
          m_str,
          fig_pie,
          fig_bar,
          df_nuevos,
          df_saldados,
          df_pendientes,
          df_omitidos,
      ) = procesar_archivos(file_hist, file_anexo)

    st.markdown("---")
    st.markdown(m_str)

    if fig_pie is not None and fig_bar is not None:
      gcol1, gcol2 = st.columns(2)
      with gcol1:
        st.plotly_chart(fig_pie, use_container_width=True)
      with gcol2:
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("## 📄 **Detalle de Registros (Listos para Copiar y Pegar)**")

    tab1, tab2, tab3, tab4 = st.tabs([
        "🔴 Nuevos Faltantes",
        "🟢 Faltantes Saldados / Pagados",
        "🟣 Pendientes Activos",
        "⚠️ Omitidos en Anexo",
    ])

    with tab1:
      if df_nuevos.empty:
        st.info("No hay faltantes nuevos generados en el corte de hoy.")
      else:
        st.dataframe(df_nuevos, use_container_width=True)
    with tab2:
      if df_saldados.empty:
        st.info("No se detectaron faltantes saldados/pagados en este corte.")
      else:
        st.dataframe(df_saldados, use_container_width=True)
    with tab3:
      if df_pendientes.empty:
        st.info("No hay pendientes activos de fechas anteriores.")
      else:
        st.dataframe(df_pendientes, use_container_width=True)
    with tab4:
      if df_omitidos.empty:
        st.info("¡Excelente! No hay registros omitidos en este anexo.")
      else:
        st.dataframe(df_omitidos, use_container_width=True)
else:
  st.info(
      "💡 Sube el **Histórico Maestro (.csv)** y el **Anexo del día**, luego"
      " presiona **EJECUTAR AUDITORÍA**."
  )
