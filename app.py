import streamlit as st
import pandas as pd
import openpyxl
import re
import warnings
import unicodedata
import plotly.express as px
import csv

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Auditoría de Faltantes - Popayán",
    page_icon="🥤",
    layout="wide"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- FUNCIONES DE LIMPIEZA Y FORMATO ---
def remover_tildes_y_raros(texto):
    if pd.isna(texto): return ""
    txt = str(texto).upper().strip()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
    txt = re.sub(r'[^A-Z0-9]', ' ', txt)
    return re.sub(r'\s+', ' ', txt).strip()

def normalizar_sap(val):
    if pd.isna(val): return "1"
    txt = str(val).strip().split('.')[0]
    txt = re.sub(r'\D', '', txt)
    if txt in ['1', ''] or not txt: return '1'
    return txt

def limpiar_monto(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).replace('$', '').replace('-', '0').strip()
    if not val_str or val_str.lower() in ['none', 'nan']: return 0.0
    val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def convertir_a_fecha(val):
    if pd.isna(val): return pd.NaT
    dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
    return dt

def formatear_fecha_ui(val_dt):
    if pd.isna(val_dt): return 'Sin Fecha'
    return val_dt.strftime('%d/%m/%Y')


# --- CARGAR HISTÓRICO CSV ---
def cargar_historico_csv(file_obj):
    if file_obj is None: return pd.DataFrame()
    try:
        rows = []
        content = file_obj.getvalue().decode('latin1', errors='ignore')
        reader = csv.reader(content.splitlines(), delimiter=';')
        for r in reader:
            if any(str(cell).strip() for cell in r):
                rows.append(r)
            
        if not rows: return pd.DataFrame()
        
        header_idx = -1
        for idx, r in enumerate(rows):
            row_str = " ".join([str(c).upper() for c in r])
            if "VALOR" in row_str or "IMPORTE" in row_str or "SALDO" in row_str:
                header_idx = idx
                break
                
        if header_idx == -1: header_idx = 2
        
        headers = [str(c).upper().replace('\n', ' ').strip() for c in rows[header_idx]]
        data_rows = rows[header_idx + 1:]
        
        if len(data_rows) > 5000:
            data_rows = data_rows[-5000:]
            
        df = pd.DataFrame(data_rows)
        df.columns = headers[:len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
        
        col_sap = next((c for c in df.columns if 'CLIENTE' in c or 'SAP' in c or 'DUEDOR' in c), df.columns[2] if len(df.columns) > 2 else None)
        col_nombre = next((c for c in df.columns if 'NOMBRE' in c or 'TITULAR' in c or 'RESPONSABLE' in c or 'ZONA' in c), df.columns[3] if len(df.columns) > 3 else None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c), df.columns[4] if len(df.columns) > 4 else None)
        col_valor = next((c for c in df.columns if 'VALOR' in c or 'IMPORTE' in c), df.columns[7] if len(df.columns) > 7 else None)
        col_abono = next((c for c in df.columns if 'ABONO' in c), df.columns[8] if len(df.columns) > 8 else None)
        col_saldo = next((c for c in df.columns if 'SALDO' in c), df.columns[9] if len(df.columns) > 9 else None)
        col_motivo = next((c for c in df.columns if 'MOTIVO' in c), df.columns[11] if len(df.columns) > 11 else None)
        
        df['Deudor_SAP_OK'] = df[col_sap].apply(normalizar_sap) if col_sap else '1'
        df['Nombre_Deudor_OK'] = df[col_nombre].apply(remover_tildes_y_raros) if col_nombre else 'RESPONSABLE / ZONA'
        df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        df['Saldo_Num'] = df[col_saldo].apply(limpiar_monto) if col_saldo else (df['Valor_Faltante_Num'] - df['Abonos_Num'])
        df['Fecha_DT'] = df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
        df['Motivo_OK'] = df[col_motivo].astype(str).str.strip() if col_motivo and col_motivo in df.columns else ''
        
        return df[(df['Saldo_Num'] > 0) & (df['Deudor_SAP_OK'] != '') & (df['Fecha_DT'].notna()) & (df['Fecha_DT'].dt.year >= 2026)].copy()
    except Exception as e:
        return pd.DataFrame()


# --- CARGAR ANEXO EXCEL ---
def cargar_anexo_excel(file_obj):
    if file_obj is None: return pd.DataFrame()
    try:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))
        df_raw = df_raw.dropna(how='all').reset_index(drop=True)
        
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = [str(c).upper().replace('\n', ' ').strip() for c in row if pd.notna(c)]
            if any("VALOR" in c or "IMPORTE" in c for c in row_str) and any("FECHA" in c for c in row_str):
                header_idx = idx
                break
                
        df = df_raw.iloc[header_idx + 1:].copy()
        headers = [str(c).upper().replace('\n', ' ').strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
        df.columns = headers[:len(df.columns)]
        df = df.loc[:, df.columns.notna()].copy()
        df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
        
        col_client_name = next((c for c in df.columns if 'NOMBRE' in c and 'CLIENTE' in c), None)
        col_trans_code = next((c for c in df.columns if 'DUEDOR SAP' in c or 'SAP' in c or 'CODIGO' in c), None)
        col_trans_name = next((c for c in df.columns if 'TITULAR' in c or 'TRANSPORTADOR' in c or 'RESPONSABLE' in c or 'ZONA' in c), None)
        col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
        col_valor = next((c for c in df.columns if 'VALOR' in c or 'IMPORTE' in c), None)
        col_abono = next((c for c in df.columns if 'ABONO' in c), None)
        col_saldo = next((c for c in df.columns if 'SALDO' in c), None)
        col_motivo = next((c for c in df.columns if 'MOTIVO' in c), None)
        
        df['Valor_Faltante_Num'] = df[col_valor].apply(limpiar_monto) if col_valor else 0.0
        df['Abonos_Num'] = df[col_abono].apply(limpiar_monto) if col_abono else 0.0
        
        saldo_calc = df['Valor_Faltante_Num'] - df['Abonos_Num']
        if col_saldo:
            s_val = df[col_saldo].apply(limpiar_monto)
            df['Saldo_Num'] = s_val.where(s_val > 0, saldo_calc)
        else:
            df['Saldo_Num'] = saldo_calc
        
        nombres_finales, saps_finales, mot_finales = [], [], []
        for _, row in df.iterrows():
            t_name = remover_tildes_y_raros(row[col_trans_name]) if col_trans_name and pd.notna(row[col_trans_name]) else ''
            c_name = remover_tildes_y_raros(row[col_client_name]) if col_client_name and pd.notna(row[col_client_name]) else ''
            
            if t_name and t_name not in ['POSTOBON POPAYAN', 'POSTOBON', 'TRANSPORTADOR', '']:
                nombre_str = t_name
            elif c_name and c_name not in ['POSTOBON POPAYAN', 'POSTOBON', 'CLIENTE', '']:
                nombre_str = c_name
            else:
                nombre_str = t_name if t_name else (c_name if c_name else 'RESPONSABLE / ZONA')
            
            c_code = normalizar_sap(row[col_trans_code]) if col_trans_code and pd.notna(row[col_trans_code]) else '1'
            motivo = str(row[col_motivo]).strip() if col_motivo and pd.notna(row[col_motivo]) else ''
            
            nombres_finales.append(nombre_str)
            saps_finales.append(c_code)
            mot_finales.append(motivo)
            
        df['Nombre_Deudor_OK'] = nombres_finales
        df['Deudor_SAP_OK'] = saps_finales
        df['Fecha_DT'] = df[col_fecha].apply(convertir_a_fecha) if col_fecha else pd.NaT
        df['Fecha_UI'] = df['Fecha_DT'].apply(formatear_fecha_ui)
        df['Motivo_OK'] = mot_finales
        
        return df[df['Valor_Faltante_Num'] > 0].copy()
    except Exception as e:
        return pd.DataFrame()


# --- MOTOR DE AUDITORÍA LÓGICO Y EXACTO ---
def procesar_archivos(file_hist, file_anexo):
    if file_anexo is None:
        return "⚠️ Por favor sube el Anexo del día para realizar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_anexo = cargar_anexo_excel(file_anexo)
    df_hist = cargar_historico_csv(file_hist) if file_hist is not None else pd.DataFrame()
    
    if df_anexo.empty:
        return "⚠️ El anexo cargado no tiene registros válidos.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    fechas_validas = df_anexo['Fecha_DT'].dropna()
    fecha_dia_actual = fechas_validas.max() if not fechas_validas.empty else pd.Timestamp('2026-10-23')

    # Diccionario del Histórico para cruce exacto por (SAP, Fecha)
    mapa_hist_completo = {}
    if not df_hist.empty:
        for _, r in df_hist.iterrows():
            sap = str(r['Deudor_SAP_OK']).strip()
            f_str = r['Fecha_DT'].strftime('%Y-%m-%d') if pd.notna(r['Fecha_DT']) else ''
            monto = round(r['Valor_Faltante_Num'], 2)
            mapa_hist_completo[(sap, f_str, monto)] = r
            # También guardamos por (SAP, Fecha) para abonos parciales
            mapa_hist_completo[(sap, f_str)] = r

    # 1. PROCESAR REGISTROS DEL ANEXO DE HOY
    saldados_list = []
    activos_list = []
    
    keys_atendidas_hoy = set()

    for _, row in df_anexo.iterrows():
        sap = str(row['Deudor_SAP_OK']).strip()
        f_str = row['Fecha_DT'].strftime('%Y-%m-%d') if pd.notna(row['Fecha_DT']) else ''
        monto_hoy = round(row['Valor_Faltante_Num'], 2)
        key_exact = (sap, f_str, monto_hoy)
        key_generic = (sap, f_str)
        
        keys_atendidas_hoy.add(key_generic)
        keys_atendidas_hoy.add(key_exact)
        
        row_f = row.copy()
        
        # Si el saldo es 0 o está pagado hoy
        if row_f['Saldo_Num'] <= 0.01 or row_f['Abonos_Num'] >= row_f['Valor_Faltante_Num']:
            row_f['Abonos_Num'] = row_f['Valor_Faltante_Num']
            row_f['Saldo_Num'] = 0.0
            saldados_list.append(row_f)
        else:
            # Si viene con saldo activo, cruzamos con el histórico para recuperar el Valor Inicial si hubo abono parcial
            if key_generic in mapa_hist_completo:
                r_hist = mapa_hist_completo[key_generic]
                v_inicial_hist = r_hist['Valor_Faltante_Num']
                
                # Ejemplo: Eduar Montenegro de $1.660.881 con abono de $700.000 -> saldo $960.881
                if v_inicial_hist > row_f['Valor_Faltante_Num']:
                    row_f['Valor_Faltante_Num'] = v_inicial_hist
                    row_f['Abonos_Num'] = v_inicial_hist - row_f['Saldo_Num']
                elif row_f['Abonos_Num'] > 0:
                    row_f['Valor_Faltante_Num'] = row_f['Saldo_Num'] + row_f['Abonos_Num']
                    
                if r_hist['Nombre_Deudor_OK'] and r_hist['Nombre_Deudor_OK'] not in ['POSTOBON POPAYAN', 'RESPONSABLE / ZONA']:
                    row_f['Nombre_Deudor_OK'] = r_hist['Nombre_Deudor_OK']
                    
            activos_list.append(row_f)

    df_activos = pd.DataFrame(activos_list) if activos_list else pd.DataFrame(columns=df_anexo.columns)
    saldados_anexo = pd.DataFrame(saldados_list) if saldados_list else pd.DataFrame(columns=df_anexo.columns)

    # Separar Activos en Nuevos Faltantes (Día Actual) vs Pendientes Activos (Días Anteriores)
    if not df_activos.empty and pd.notna(fecha_dia_actual):
        es_dia_actual = df_activos['Fecha_DT'].dt.date == fecha_dia_actual.date()
        nuevos = df_activos[es_dia_actual].copy()
        pendientes = df_activos[~es_dia_actual].copy()
    else:
        nuevos = df_activos.copy()
        pendientes = pd.DataFrame(columns=df_anexo.columns)

    # 2. IDENTIFICAR EN EL HISTÓRICO LOS REGISTROS QUE YA NO APARECEN HOY (FUERAN PAGADOS)
    saldados_hist_list = []
    if not df_hist.empty:
        for _, r_hist in df_hist.iterrows():
            sap_h = str(r_hist['Deudor_SAP_OK']).strip()
            f_str_h = r_hist['Fecha_DT'].strftime('%Y-%m-%d') if pd.notna(r_hist['Fecha_DT']) else ''
            monto_h = round(r_hist['Valor_Faltante_Num'], 2)
            
            # Si el faltante estaba pendiente en el histórico pero YA NO viene en el anexo de hoy, FUE PAGADO / SALDADO
            if (sap_h, f_str_h) not in keys_atendidas_hoy and (sap_h, f_str_h, monto_h) not in keys_atendidas_hoy:
                row_pagado = r_hist.copy()
                row_pagado['Abonos_Num'] = row_pagado['Valor_Faltante_Num']
                row_pagado['Saldo_Num'] = 0.0
                saldados_hist_list.append(row_pagado)

    saldados_hist = pd.DataFrame(saldados_hist_list) if saldados_hist_list else pd.DataFrame(columns=df_anexo.columns)
    
    # Unir todos los saldados/pagados
    saldados = pd.concat([saldados_anexo, saldados_hist], ignore_index=True) if not saldados_hist.empty else saldados_anexo
    saldados = saldados.drop_duplicates(subset=['Deudor_SAP_OK', 'Fecha_DT', 'Valor_Faltante_Num']) if not saldados.empty else saldados

    # Pestaña Omitidos queda limpia
    omitidos = pd.DataFrame(columns=df_anexo.columns)
            
    def preparar_df_ui(df_in):
        if df_in is None or df_in.empty:
            return pd.DataFrame(columns=['Nombre Deudor', 'SAP', 'Fecha', 'Valor Faltante Inicial ($)', 'Abonos Realizados ($)', 'Saldo Pendiente Actual ($)', 'Motivo'])
        
        out = pd.DataFrame()
        out['Nombre Deudor'] = df_in['Nombre_Deudor_OK']
        out['SAP'] = df_in['Deudor_SAP_OK']
        out['Fecha'] = df_in['Fecha_UI']
        out['Valor Faltante Inicial ($)'] = df_in['Valor_Faltante_Num'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        out['Abonos Realizados ($)'] = df_in['Abonos_Num'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        out['Saldo Pendiente Actual ($)'] = df_in['Saldo_Num'].apply(lambda x: f"${int(round(limpiar_monto(x))):,}")
        out['Motivo'] = df_in['Motivo_OK'] if 'Motivo_OK' in df_in.columns else ''
        return out
        
    df_nuevos_ui = preparar_df_ui(nuevos)
    df_saldados_ui = preparar_df_ui(saldados)
    df_pendientes_ui = preparar_df_ui(pendientes)
    df_omitidos_ui = preparar_df_ui(omitidos)
    
    tot_nuevos = nuevos['Saldo_Num'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Abonos_Num'].sum() if not saldados.empty else 0
    tot_pendientes = pendientes['Saldo_Num'].sum() if not pendientes.empty else 0
    tot_omitidos = omitidos['Saldo_Num'].sum() if not omitidos.empty else 0
    
    m_str = f"## 🥤 **AUDITORÍA DE FALTANTES - POPAYÁN**\n\n"
    m_str += f"- 🔴 **Nuevos Faltantes (Día Actual - {formatear_fecha_ui(fecha_dia_actual)}):** {len(nuevos)} registro(s) | **${int(tot_nuevos):,}**\n"
    m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} registro(s) | **${int(tot_saldados):,}**\n"
    m_str += f"- 🟣 **Pendientes Activos (Días Anteriores):** {len(pendientes)} registro(s) | **${int(tot_pendientes):,}**\n"
    m_str += f"- ⚠️ **Omitidos en Anexo:** {len(omitidos)} registro(s) | **${int(tot_omitidos):,}**\n"
    
    categorias = ['Nuevos Faltantes', 'Faltantes Saldados', 'Pendientes Activos', 'Omitidos en Anexo']
    cantidades = [len(nuevos), len(saldados), len(pendientes), len(omitidos)]
    montos = [tot_nuevos, tot_saldados, tot_pendientes, tot_omitidos]
    colores = {'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Activos': '#AB63FA', 'Omitidos en Anexo': '#FFA15A'}
    
    df_summary = pd.DataFrame({
        'Categoría': categorias,
        'Cantidad': cantidades,
        'Monto': montos
    })
    
    fig_pie = px.pie(
        df_summary, values='Cantidad', names='Categoría',
        title="<b>Distribución por Cantidad de Casos</b>",
        color='Categoría',
        color_discrete_map=colores,
        hole=0.4
    )
    
    fig_bar = px.bar(
        df_summary, x='Categoría', y='Monto', text_auto='.0f',
        title="<b>Impacto Financiero Real ($ COP)</b>",
        color='Categoría',
        color_discrete_map=colores
    )
    fig_bar.update_layout(showlegend=False, yaxis_title="Monto ($)")
    
    return m_str, fig_pie, fig_bar, df_nuevos_ui, df_saldados_ui, df_pendientes_ui, df_omitidos_ui


# --- INTERFAZ EN STREAMLIT ---
st.markdown("# 🥤 **POSTOBÓN POPAYÁN - SISTEMA DE AUDITORÍA DE FALTANTES**")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Carga de Archivos")
    file_hist = st.file_uploader("Histórico Maestro (.csv)", type=['csv'])
    file_anexo = st.file_uploader("Anexo del Día Actual (.xlsx / .xlsm)", type=['xlsx', 'xlsm', 'csv'])

with col2:
    st.subheader("Estado del Proceso")
    ejecutar = st.button("🚀 EJECUTAR AUDITORÍA", type="primary")

if ejecutar:
    if file_anexo is None:
        st.warning("⚠️ Por favor sube el Anexo del día para realizar la auditoría.")
    else:
        with st.spinner("Procesando auditoría..."):
            m_str, fig_pie, fig_bar, df_nuevos, df_saldados, df_pendientes, df_omitidos = procesar_archivos(file_hist, file_anexo)
            
        st.markdown("---")
        st.markdown(m_str)
        
        gcol1, gcol2 = st.columns(2)
        with gcol1:
            st.plotly_chart(fig_pie, use_container_width=True)
        with gcol2:
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("## 📄 **Detalle de Registros (Listos para Copiar y Pegar)**")
        
        tab1, tab2, tab3, tab4 = st.tabs(["🔴 Nuevos Faltantes", "🟢 Faltantes Saldados / Pagados", "🟣 Pendientes Activos", "⚠️ Omitidos en Anexo"])
        
        with tab1:
            st.dataframe(df_nuevos, use_container_width=True)
        with tab2:
            if df_saldados.empty:
                st.info("No se detectaron faltantes saldados/pagados en este corte.")
            else:
                st.dataframe(df_saldados, use_container_width=True)
        with tab3:
            st.dataframe(df_pendientes, use_container_width=True)
        with tab4:
            if df_omitidos.empty:
                st.info("¡Excelente! No hay registros omitidos en este anexo.")
            else:
                st.dataframe(df_omitidos, use_container_width=True)
else:
    st.info("💡 Sube el **Histórico Maestro (.csv)** y el **Anexo del día**, luego presiona **EJECUTAR AUDITORÍA**.")
