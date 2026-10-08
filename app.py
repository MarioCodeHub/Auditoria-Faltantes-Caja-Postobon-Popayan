import streamlit as st
import pandas as pd
import openpyxl
import warnings
import plotly.express as px

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Auditoría de Faltantes - Postobón Popayán",
    page_icon="🥤",
    layout="wide"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- FUNCIONES DE LECTURA Y LIMPIEZA DE DATOS ---
def cargar_y_limpiar(file_obj):
    if file_obj is None: return pd.DataFrame()
    filename = file_obj.name
    if filename.endswith('.csv'):
        encodings = ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']
        for enc in encodings:
            try:
                df_raw = pd.read_csv(file_obj, encoding=enc, sep=None, engine='python', header=None)
                break
            except Exception:
                continue
    else:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        sheet = wb.active
        df_raw = pd.DataFrame(list(sheet.values))
        
    df_raw = df_raw.dropna(how='all').reset_index(drop=True)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = [str(c).upper() for c in row if pd.notna(c)]
        if any("DEUDOR" in c or "VALOR" in c or "IMPORTE" in c or "SALDO" in c for c in row_str):
            header_idx = idx
            break
            
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = [str(c).strip() if pd.notna(c) else f"COL_{i}" for i, c in enumerate(df_raw.iloc[header_idx].values)]
    df = df.loc[:, df.columns.notna()].copy()
    df = df.loc[:, ~df.columns.duplicated(keep='first')].copy()
    
    mapa = {}
    for col in df.columns:
        c_str = col.strip().upper()
        if 'DEUDOR SAP' in c_str or 'DEUDOR SAP' in c_str or 'CODIGO CLIENT' in c_str or 'SAP' in c_str:
            if "Deudor_SAP" not in mapa.values(): mapa[col] = "Deudor_SAP"
        elif 'TITULAR ZONA' in c_str or 'NOMBRE DEL RESPONSABLE' in c_str or 'DEUDOR' in c_str or 'TITULAR' in c_str:
            if 'Nombre_Deudor' not in mapa.values(): mapa[col] = 'Nombre_Deudor'
        elif 'FECHA EN QUE SE GENER' in c_str or 'STR' in c_str or 'FECHA' in c_str or 'FECHA GENER' in c_str:
            if 'Fecha_Genera' not in mapa.values(): mapa[col] = 'Fecha_Genera'
        elif 'VALOR DEL FALTAN' in c_str or c_str == 'IMPORTE' or 'VALOR ORIGINAL' in c_str or 'VALOR' in c_str:
            if 'Valor_Original' not in mapa.values(): mapa[col] = 'Valor_Original'
        elif 'ABONO' in c_str:
            if 'Abonos' not in mapa.values(): mapa[col] = 'Abonos'
        elif 'SALDO' in c_str:
            if 'Saldo_Pendiente' not in mapa.values(): mapa[col] = 'Saldo_Pendiente'
            
    df = df.rename(columns=mapa)
    return df

def limpiar_monto(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return int(round(val))
    val_str = str(val).replace('$', '').replace('-', '0').strip()
    if not val_str or val_str.lower() in ['none', 'nan']: return 0
    val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return int(round(float(val_str)))
    except:
        return 0

def extraer_col(df, col_name, default=''):
    if col_name in df.columns:
        res = df[col_name]
        return res.iloc[:, 0] if isinstance(res, pd.DataFrame) else res
    return pd.Series([default] * len(df))


# --- MOTOR DE AUDITORÍA Y DASHBOARD ---
def procesar_archivos(file_hist, file_anexo):
    if file_hist is None or file_anexo is None:
        return "⚠️ Por favor sube ambos archivos para iniciar la auditoría.", None, None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_hist = cargar_y_limpiar(file_hist)
    df_anexo = cargar_y_limpiar(file_anexo)
    
    for df in [df_hist, df_anexo]:
        df['Valor_Original_OK'] = extraer_col(df, 'Valor_Original').apply(limpiar_monto)
        df['Abonos_OK'] = extraer_col(df, 'Abonos').apply(limpiar_monto)
        df['Saldo_OK'] = extraer_col(df, 'Saldo_Pendiente').apply(limpiar_monto)
        
        mask_no_saldo = (df['Saldo_OK'] == 0) & (df['Valor_Original_OK'] > 0)
        df.loc[mask_no_saldo, 'Saldo_OK'] = df.loc[mask_no_saldo, 'Valor_Original_OK'] - df.loc[mask_no_saldo, 'Abonos_OK']
        
        df['Deudor_SAP_OK'] = extraer_col(df, 'Deudor_SAP').astype(str).str.split('.').str[0].str.strip()
        df['Nombre_Deudor_OK'] = extraer_col(df, 'Nombre_Deudor').astype(str).str.strip()
        df['Nombre_Deudor_OK'] = df['Nombre_Deudor_OK'].replace(['nan', 'None', '', 'NaN'], 'Sin Nombre Registrado')
        
        raw_fechas = extraer_col(df, 'Fecha_Genera')
        df['Fecha_Genera_DT'] = pd.to_datetime(raw_fechas, errors='coerce', dayfirst=True)
        df['Fecha_Genera_OK'] = df['Fecha_Genera_DT'].dt.strftime('%d/%m/%Y').fillna(raw_fechas.astype(str))
        df['Fecha_Genera_OK'] = df['Fecha_Genera_OK'].replace(['nan', 'None', 'NaT'], 'Sin Fecha')
        
    df_hist = df_hist.tail(2000).copy()
    
    df_hist_activos = df_hist[
        (df_hist['Saldo_OK'] > 0) &
        (df_hist['Deudor_SAP_OK'] != '') &
        (df_hist['Deudor_SAP_OK'].str.lower() != 'none') &
        (df_hist['Deudor_SAP_OK'].str.lower() != 'nan')
    ].copy()
    
    df_anexo_activos = df_anexo[
        (df_anexo['Saldo_OK'] > 0) &
        (df_anexo['Deudor_SAP_OK'] != '') &
        (df_anexo['Deudor_SAP_OK'].str.lower() != 'none') &
        (df_anexo['Deudor_SAP_OK'].str.lower() != 'nan')
    ].copy()
    
    fecha_corte_dt = df_anexo_activos['Fecha_Genera_DT'].max()
    fecha_corte_str = fecha_corte_dt.strftime('%d/%m/%Y') if pd.notna(fecha_corte_dt) else "Corte del Día"
    
    nuevos = df_anexo_activos[df_anexo_activos['Fecha_Genera_DT'] == fecha_corte_dt].copy()
    sin_cambios = df_anexo_activos[df_anexo_activos['Fecha_Genera_DT'] < fecha_corte_dt].copy()
    
    anexo_counts = sin_cambios.groupby(['Deudor_SAP_OK', 'Saldo_OK']).size().to_dict()
    saldados_list = []
    for idx, row in df_hist_activos.iterrows():
        clave = (row['Deudor_SAP_OK'], row['Saldo_OK'])
        if anexo_counts.get(clave, 0) > 0:
            anexo_counts[clave] -= 1
        else:
            saldados_list.append(row)
            
    saldados = pd.DataFrame(saldados_list) if saldados_list else pd.DataFrame(columns=df_hist_activos.columns)
    if not saldados.empty and 'Fecha_Genera_DT' in saldados.columns:
        saldados = saldados[saldados['Fecha_Genera_DT'] != fecha_corte_dt].copy()
        
    def preparar_df_ui(df_in, col_monto_name='Saldo ($)'):
        if df_in.empty:
            return pd.DataFrame(columns=['Nombre Deudor', 'SAP', 'Fecha', col_monto_name])
        out = df_in[['Nombre_Deudor_OK', 'Deudor_SAP_OK', 'Fecha_Genera_OK', 'Saldo_OK']].copy()
        out.columns = ['Nombre Deudor', 'SAP', 'Fecha', col_monto_name]
        out[col_monto_name] = out[col_monto_name].apply(lambda x: f"${int(x):,}")
        return out
        
    df_nuevos_ui = preparar_df_ui(nuevos, 'Monto Nuevo ($)')
    df_saldados_ui = preparar_df_ui(saldados, 'Monto Saldado ($)')
    df_pendientes_ui = preparar_df_ui(sin_cambios, 'Saldo Pendiente ($)')
    
    tot_nuevos = nuevos['Saldo_OK'].sum() if not nuevos.empty else 0
    tot_saldados = saldados['Saldo_OK'].sum() if not saldados.empty else 0
    tot_pendientes = sin_cambios['Saldo_OK'].sum() if not sin_cambios.empty else 0
    
    m_str = f"## 🥤 **RESUMEN AUDITORÍA AL CORTE: {fecha_corte_str}**\n\n"
    m_str += f"- 🔴 **Nuevos Faltantes Registrados:** {len(nuevos)} caso(s) | **${int(tot_nuevos):,}**\n"
    m_str += f"- 🟢 **Faltantes Saldados / Pagados:** {len(saldados)} caso(s) | **${int(tot_saldados):,}**\n"
    m_str += f"- 🟣 **Pendientes Activos Sin Cambios:** {len(sin_cambios)} caso(s) | **${int(tot_pendientes):,}**\n"
    
    df_summary = pd.DataFrame({
        'Categoría': ['Nuevos Faltantes', 'Faltantes Saldados', 'Pendientes Sin Cambios'],
        'Cantidad': [len(nuevos), len(saldados), len(sin_cambios)],
        'Monto': [tot_nuevos, tot_saldados, tot_pendientes]
    })
    
    fig_pie = px.pie(
        df_summary, values='Cantidad', names='Categoría',
        title="<b>Distribución por Cantidad de Casos</b>",
        color='Categoría',
        color_discrete_map={'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Sin Cambios': '#AB63FA'},
        hole=0.4
    )
    
    fig_bar = px.bar(
        df_summary, x='Categoría', y='Monto', text_auto='.0f',
        title="<b>Impacto Financiero ($ COP)</b>",
        color='Categoría',
        color_discrete_map={'Nuevos Faltantes': '#EF553B', 'Faltantes Saldados': '#00CC96', 'Pendientes Sin Cambios': '#AB63FA'}
    )
    fig_bar.update_layout(showlegend=False, yaxis_title="Monto ($)")
    
    return m_str, fig_pie, fig_bar, df_nuevos_ui, df_saldados_ui, df_pendientes_ui


# --- INTERFAZ EN STREAMLIT ---
st.markdown("# 🥤 **POSTOBÓN - AUDITORÍA DE FALTANTES EN CAJA (POPAYÁN)**")
st.markdown("### Plataforma Interactiva para Aprendices de Auditoría")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Cargar Archivos")
    file_hist = st.file_uploader("1. Histórico Maestro (.csv / .xlsx)", type=['csv', 'xlsx'])
    file_anexo = st.file_uploader("2. Anexo del Día a Auditar (.xlsm / .xlsx / .csv)", type=['xlsm', 'xlsx', 'csv'])

with col2:
    st.subheader("Ejecución")
    ejecutar = st.button("🚀 EJECUTAR AUDITORÍA", type="primary")

if ejecutar:
    if file_hist is None or file_anexo is None:
        st.warning("⚠️ Por favor sube ambos archivos para iniciar la auditoría.")
    else:
        with st.spinner("Procesando histórico y anexo de Popayán..."):
            m_str, fig_pie, fig_bar, df_nuevos, df_saldados, df_pendientes = procesar_archivos(file_hist, file_anexo)
            
        st.markdown("---")
        st.markdown(m_str)
        
        gcol1, gcol2 = st.columns(2)
        with gcol1:
            st.plotly_chart(fig_pie, use_container_width=True)
        with gcol2:
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("## 📄 **Detalle Completo de Registros**")
        
        tab1, tab2, tab3 = st.tabs(["🔴 Nuevos Faltantes", "🟢 Faltantes Saldados / Pagados", "🟣 Pendientes Sin Cambios"])
        
        with tab1:
            st.dataframe(df_nuevos, use_container_width=True)
        with tab2:
            st.dataframe(df_saldados, use_container_width=True)
        with tab3:
            st.dataframe(df_pendientes, use_container_width=True)
else:
    st.info("💡 Sube ambos archivos arriba y presiona **EJECUTAR AUDITORÍA** para ver el informe de Popayán.")
