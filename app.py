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


def extraer_nombre_real_fila(row, col_titular=None, col_cliente=None):
    if col_titular and col_titular in row.index and pd.notna(row[col_titular]):
        txt = remover_tildes_y_raros(row[col_titular])
        if (
            txt
            and not txt.isdigit()
            and len(txt) > 2
            and txt not in PALABRAS_PROHIBIDAS
        ):
            return txt

    if col_cliente and col_cliente in row.index and pd.notna(row[col_cliente]):
        txt = remover_tildes_y_raros(row[col_cliente])
        if (
            txt
            and not txt.isdigit()
            and len(txt) > 2
            and txt not in PALABRAS_PROHIBIDAS
        ):
            txt_clean = re.sub(r"^CLIENTE\s+\d+\s*", "", txt).strip()
            if txt_clean and not txt_clean.isdigit() and len(txt_clean) > 2:
                return txt_clean
            elif txt and not txt.isdigit():
                return txt

    for val in row.values:
        if pd.notna(val):
            txt = remover_tildes_y_raros(val)
            if (
                txt
                and not txt.isdigit()
                and not re.match(r"^\d{2}/\d{2}/\d{4}$", txt)
                and len(txt) > 3
            ):
                if (
                    txt not in PALABRAS_PROHIBIDAS
                    and not txt.startswith("31000")
                ):
                    if (
                        "PAGO" not in txt
                        and "FACTURA" not in txt
                        and "DIFERENCIA" not in txt
                    ):
                        return txt

    return "RESPONSABLE / ZONA"


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
                or "MOTIVO" in row_str
            ):
                header_idx = idx
                break

        if header_idx == -1:
            header_idx = 2

        headers = [
            str(c).upper().replace("\n", " ").strip() for c in rows[header_idx]
        ]
        data_rows = rows[header_idx + 1 :]

        if len(data_rows) > 5000:
            data_rows = data_rows[-5000:]

        df = pd.DataFrame(data_rows)
        df.columns = headers[: len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep="first")].copy()

        col_sap = next(
            (
                c
                for c in df.columns
                if "DUEDOR SAP" in c
                or "DEUDOR SAP" in c
                or "DUEDOR" in c
                or "DEUDOR" in c
            ),
            None,
        )
        if not col_sap:
            col_sap = next(
                (
                    c
                    for c in df.columns
                    if "SAP" in c and "CLE" not in c and "SI" not in c
                ),
                None,
            )

        col_titular = next(
            (
                c
                for c in df.columns
                if "TITULAR" in c
                or "RESPONSABLE" in c
                or "ZONA" in c
                or "TRANSPORTADOR" in c
            ),
            None,
        )
        col_cliente = next(
            (c for c in df.columns if "OBSERVAC" in c or "CLIENTE" in c), None
        )
        col_fecha = next((c for c in df.columns if "FECHA" in c), None)
        col_valor = next(
            (c for c in df.columns if "IMPORTE" in c or "VALOR" in c), None
        )
        col_abono = next((c for c in df.columns if "ABONO" in c), None)
        col_saldo = next((c for c in df.columns if "SALDO" in c), None)
        col_motivo = next((c for c in df.columns if "MOTIVO" in c), None)

        nombres_extraidos = []
        for _, row in df.iterrows():
            nombre_ok = extraer_nombre_real_fila(row, col_titular, col_cliente)
            nombres_extraidos.append(nombre_ok)

        df["Nombre_Deudor_OK"] = nombres_extraidos
        df["Deudor_SAP_OK"] = (
            df[col_sap].apply(normalizar_sap) if col_sap else "1"
        )
        df["Valor_Faltante_Num"] = (
            df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        )
        df["Abonos_Num"] = (
            df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        )
        df["Saldo_Num"] = (
            df[col_saldo].apply(limpiar_monto)
            if col_saldo
            else (df["Valor_Faltante_Num"] - df["Abonos_Num"])
        )
        df["Fecha_DT"] = (
            df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        )
        df["Fecha_UI"] = df["Fecha_DT"].apply(formatear_fecha_ui)
        df["Motivo_OK"] = (
            df[col_motivo].astype(str).str.strip()
            if col_motivo and col_motivo in df.columns
            else ""
        )

        return df[
            (df["Valor_Faltante_Num"] > 0)
            & (df["Fecha_DT"].notna())
            & (df["Fecha_DT"].dt.year >= 2026)
        ].copy()
    except Exception as e:
        return pd.DataFrame()


# --- CARGAR ANEXO EXCEL CON MAPEO BLINDADO ANTI-CONFUSIÓN DE SALDO ---
def cargar_anexo_excel(file_obj):
    if file_obj is None:
        return pd.DataFrame()
    try:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))
        df_raw = df_raw.dropna(how="all").reset_index(drop=True)

        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = [
                str(c).upper().replace("\n", " ").strip()
                for c in row
                if pd.notna(c)
            ]
            if any("IMPORTE" in c or "VALOR" in c for c in row_str) and any(
                "FECHA" in c for c in row_str
            ):
                header_idx = idx
                break

        df = df_raw.iloc[header_idx + 1 :].copy()
        headers = [
            str(c).upper().replace("\n", " ").strip()
            if pd.notna(c)
            else f"COL_{i}"
            for i, c in enumerate(df_raw.iloc[header_idx].values)
        ]
        df.columns = headers[: len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep="first")].copy()

        col_sap = next(
            (
                c
                for c in df.columns
                if "DEUDOR SAP" in c or "SAP" in c or "DUEDOR" in c
            ),
            None,
        )
        col_titular = next(
            (
                c
                for c in df.columns
                if "TITULAR" in c
                or "RESPONSABLE" in c
                or "ZONA" in c
                or "TRANSPORTADOR" in c
            ),
            None,
        )
        col_cliente = next(
            (c for c in df.columns if "CLIENTE" in c or "OBSERVAC" in c), None
        )
        col_fecha = next(
            (c for c in df.columns if "FECHA" in c and "GENERA" in c),
            next((c for c in df.columns if "FECHA" in c), None),
        )

        # BÚSQUEDA 100% BLINDADA: Excluye explícitamente la palabra SALDO para evitar choques con el paréntesis
        col_valor = None
        for c in df.columns:
            c_up = str(c).upper()
            if (
                ("VALOR" in c_up or "IMPORTE" in c_up or "FALTANT" in c_up)
                and "SALDO" not in c_up
                and "ABONO" not in c_up
            ):
                col_valor = c
                break

        col_abono = None
        for c in df.columns:
            c_up = str(c).upper()
            if "ABONO" in c_up and "SALDO" not in c_up:
                col_abono = c
                break

        col_saldo = None
        for c in df.columns:
            c_up = str(c).upper()
            if "SALDO" in c_up:
                col_saldo = c
                break

        col_motivo = next((c for c in df.columns if "MOTIVO" in c), None)

        # Depuración en barra lateral para verificar mapeo exacto
        st.sidebar.info(
            f"🔎 **Mapeo Excel Detectado:**\n- Valor Inicial: `{col_valor}`\n-"
            f" Abonos: `{col_abono}`\n- Saldo: `{col_saldo}`"
        )

        nombres_finales, saps_finales, mot_finales = [], [], []
        importes_finales, abonos_finales, saldos_finales = [], [], []

        for _, row in df.iterrows():
            nombre_ok = extraer_nombre_real_fila(row, col_titular, col_cliente)
            c_code = (
                normalizar_sap(row[col_sap])
                if col_sap and pd.notna(row[col_sap])
                else "1"
            )
            motivo = (
                str(row[col_motivo]).strip()
                if col_motivo and pd.notna(row[col_motivo])
                else ""
            )

            val_imp = limpiar_monto(row[col_valor]) if col_valor else 0.0
            val_abo = limpiar_monto(row[col_abono]) if col_abono else 0.0
            val_sal = (
                limpiar_monto(row[col_saldo])
                if col_saldo
                else (val_imp - val_abo)
            )

            nombres_finales.append(nombre_ok)
            saps_finales.append(c_code)
            mot_finales.append(motivo)
            importes_finales.append(val_imp)
            abonos_finales.append(val_abo)
            saldos_finales.append(val_sal)

        df["Nombre_Deudor_OK"] = nombres_finales
        df["Deudor_SAP_OK"] = saps_finales
        df["Valor_Faltante_Num"] = importes_finales
        df["Abonos_Num"] = abonos_finales
        df["Saldo_Num"] = saldos_finales
        df["Fecha_DT"] = (
            df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        )
        df["Fecha_UI"] = df["Fecha_DT"].apply(formatear_fecha_ui)
        df["Motivo_OK"] = mot_finales

        return df[df["Valor_Faltante_Num"] > 0].copy()
    except Exception as e:
        st.error(f"Error leyendo el anexo: {e}")
        return pd.DataFrame()


# --- MOTOR DE AUDITORÍA ROBUSTO Y EQUILIBRADO ---
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
        cargar_historico_csv(file_hist)
        if file_hist is not None
        else pd.DataFrame()
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
        else pd.Timestamp("2026-10-29")
    )

    fecha_limite_reciente = fecha_dia_actual - pd.Timedelta(days=30)

    keys_exactas_anexo = set()
    keys_identidad_anexo = set()

    for _, row in df_anexo.iterrows():
        sap = str(row["Deudor_SAP_OK"]).strip()
        nom = str(row["Nombre_Deudor_OK"]).strip().upper()
        f_dt = row["Fecha_DT"]
        f_str = f_dt.strftime("%Y-%m-%d") if pd.notna(f_dt) else ""
        monto_anexo = round(float(row["Valor_Faltante_Num"]), 2)

        if sap and sap != "1":
            keys_exactas_anexo.add((sap, f_str, monto_anexo))
            keys_identidad_anexo.add((sap, f_str))
        if nom and nom not in PALABRAS_PROHIBIDAS:
            keys_exactas_anexo.add((nom, f_str, monto_anexo))
            keys_identidad_anexo.add((nom, f_str))

    activos_list = []
    for _, row in df_anexo.iterrows():
        row_f = row.copy()
        activos_list.append(row_f)

    df_activos = (
        pd.DataFrame(activos_list)
        if activos_list
        else pd.DataFrame(columns=df_anexo.columns)
    )

    if not df_activos.empty:
        df_con_saldo = df_activos[df_activos["Saldo_Num"] > 0].copy()
        df_saldados_en_anexo = df_activos[df_activos["Saldo_Num"] <= 0].copy()

        if pd.notna(fecha_dia_actual):
            es_dia_actual = (
                df_con_saldo["Fecha_DT"].dt.date == fecha_dia_actual.date()
            )
            nuevos = df_con_saldo[es_dia_actual].copy()
            pendientes = df_con_saldo[
                df_con_saldo["Fecha_DT"].dt.date != fecha_dia_actual.date()
            ].copy()
        else:
            nuevos = df_con_saldo.copy()
            pendientes = pd.DataFrame(columns=df_anexo.columns)
    else:
        nuevos = pd.DataFrame(columns=df_anexo.columns)
        pendientes = pd.DataFrame(columns=df_anexo.columns)
        df_saldados_en_anexo = pd.DataFrame(columns=df_anexo.columns)

    saldados_hist_list = []

    if not df_saldados_en_anexo.empty:
        for _, r_sal in df_saldados_en_anexo.iterrows():
            row_pagado = r_sal.copy()
            row_pagado["Abonos_Num"] = row_pagado["Valor_Faltante_Num"]
            row_pagado["Saldo_Num"] = 0.0
            saldados_hist_list.append(row_pagado)

    if not df_hist.empty:
        for _, r_hist in df_hist.iterrows():
            sap_h = str(r_hist["Deudor_SAP_OK"]).strip()
            nom_h = str(r_hist["Nombre_Deudor_OK"]).strip().upper()
            f_dt_h = r_hist["Fecha_DT"]
            f_str_h = f_dt_h.strftime("%Y-%m-%d") if pd.notna(f_dt_h) else ""
            monto_h = round(float(r_hist["Valor_Faltante_Num"]), 2)

            if pd.notna(f_dt_h) and f_dt_h.date() == fecha_dia_actual.date():
                continue

            es_reciente = (
                pd.notna(f_dt_h)
                and (f_dt_h >= fecha_limite_reciente)
                and (f_dt_h.date() < fecha_dia_actual.date())
            )
            if not es_reciente:
                continue

            sigue_en_anexo = (
                sap_h != "1"
                and (
                    (sap_h, f_str_h, monto_h) in keys_exactas_anexo
                    or (sap_h, f_str_h) in keys_identidad_anexo
                )
            ) or (
                nom_h not in PALABRAS_PROHIBIDAS
                and (
                    (nom_h, f_str_h, monto_h) in keys_exactas_anexo
                    or (nom_h, f_str_h) in keys_identidad_anexo
                )
            )

            if not sigue_en_anexo and r_hist["Saldo_Num"] > 0:
                row_pagado = r_hist.copy()
                row_pagado["Abonos_Num"] = row_pagado["Valor_Faltante_Num"]
                row_pagado["Saldo_Num"] = 0.0
                saldados_hist_list.append(row_pagado)

    saldados = (
        pd.DataFrame(saldados_hist_list)
        if saldados_hist_list
        else pd.DataFrame(columns=df_anexo.columns)
    )
    if not saldados.empty:
        saldados = saldados.drop_duplicates(
            subset=["Nombre_Deudor_OK", "Fecha_DT", "Valor_Faltante_Num"]
        )

    omitidos = pd.DataFrame(columns=df_anexo.columns)

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
        out["Motivo"] = (
            df_in["Motivo_OK"] if "Motivo_OK" in df_in.columns else ""
        )
        return out

    df_nuevos_ui = preparar_df_ui(nuevos)
    df_saldados_ui = preparar_df_ui(saldados)
    df_pendientes_ui = preparar_df_ui(pendientes)
    df_omitidos_ui = preparar_df_ui(omitidos)

    tot_nuevos = nuevos["Saldo_Num"].sum() if not nuevos.empty else 0
    tot_saldados = saldados["Abonos_Num"].sum() if not saldados.empty else 0
    tot_pendientes = (
        pendientes["Saldo_Num"].sum() if not pendientes.empty else 0
    )
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
st.markdown("# 🥤 **POSTOBÓN POPAYÁN - SISTEMA DE AUDITORÍA DE FALTANTES**")

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

        gcol1, gcol2 = st.columns(2)
        with gcol1:
            st.plotly_chart(fig_pie, use_container_width=True)
        with gcol2:
            st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown(
            "## 📄 **Detalle de Registros (Listos para Copiar y Pegar)**"
        )

        tab1, tab2, tab3, tab4 = st.tabs(
            [
                "🔴 Nuevos Faltantes",
                "🟢 Faltantes Saldados / Pagados",
                "🟣 Pendientes Activos",
                "⚠️ Omitidos en Anexo",
            ]
        )

        dataframes = [df_nuevos, df_saldados, df_pendientes, df_omitidos]
        tabs = [tab1, tab2, tab3, tab4]

        for tab, df_item in zip(tabs, dataframes):
            with tab:
                if df_item.empty:
                    st.info(
                        "No hay registros en esta categoría para este corte."
                    )
                else:
                    st.dataframe(df_item, use_container_width=True)
else:
    st.info(
        "💡 Sube el **Histórico Maestro (.csv)** y el **Anexo del día**, luego presiona **EJECUTAR AUDITORÍA**."
    )
