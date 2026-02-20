import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# --- FUNCIONES INTERNAS DE LIMPIEZA ---
def obtener_enlace_directo_drive(url):
    if "drive.google.com/file/d/" in url:
        file_id = url.split("/d/")[1].split("/")[0]
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url

def procesar_compras(file, nombre_agencia):
    df = pd.read_excel(file, header=None)
    datos = []
    factura, fecha, proveedor, comprador = None, None, None, None
    for _, row in df.iterrows():
        val_a = str(row[0]).strip()
        if "FACTURA:" in val_a:
            try:
                factura = val_a.split("FACTURA:")[1].strip()
                if "FECHA FACT:" in str(row[2]): fecha = str(row[2]).split("FECHA FACT:")[1].strip()
                if "PROVEEDOR:" in str(row[3]): proveedor = str(row[3]).split("PROVEEDOR:")[1].strip()
                if "COMPRADOR:" in str(row[4]): comprador = str(row[4]).split("COMPRADOR:")[1].strip()
            except: continue
        elif val_a.startswith("CR"): 
            descripcion = str(row[3]).strip() 
            if descripcion and descripcion.lower() != 'nan':
                datos.append({
                    "AGENCIA": nombre_agencia, "FACTURA": factura, "FECHA": fecha,
                    "PROVEEDOR": proveedor, "COMPRADOR": comprador, "NP": row[2],
                    "DESCRIPCION": row[3], "CANTIDAD": row[4], "COSTO_UNIT": row[5], "TOTAL": row[9] 
                })
    return pd.DataFrame(datos)

def procesar_traspasos(file, nombre_agencia):
    df = pd.read_excel(file, header=None)
    datos = []
    destino_actual, referencia, fecha_mov, usuario = None, None, None, None
    for _, row in df.iterrows():
        val_a = str(row[0]).strip().upper()
        if val_a.startswith("SALIDA"):
            if "HACIA" in val_a:
                try: destino_actual = val_a.split("HACIA")[1].strip()
                except: continue
            elif "SALIDA DE ALMACEN POR TRASPASO" in val_a:
                destino_actual = f"SALIDA DE ALMACEN POR TRASPASO {nombre_agencia}"
        elif "REFERENCIA:" in val_a:
            try:
                referencia = val_a.split("REFERENCIA:")[1].strip()
                if "FECHA MOV:" in str(row[2]).upper(): fecha_mov = str(row[2]).upper().split("FECHA MOV:")[1].strip()
                if "USUARIO:" in str(row[3]).upper(): usuario = str(row[3]).upper().split("USUARIO:")[1].strip()
            except: continue
        elif val_a.startswith("TRAS") and destino_actual:
            try:
                descripcion = str(row[3]).strip()
                if descripcion and descripcion.lower() != 'nan':
                    cantidad = float(row[4]) 
                    costo = float(row[5])    
                    datos.append({
                        "AGENCIA": nombre_agencia, "DESTINO": destino_actual, "REFERENCIA": referencia,
                        "FECHA_MOV": fecha_mov, "USUARIO": usuario, "NP": row[2],                 
                        "DESCRIPCION": row[3], "CANTIDAD": abs(cantidad), "COSTO_UNIT": costo, "TOTAL_COSTO": abs(cantidad) * costo 
                    })
            except: continue
    return pd.DataFrame(datos)

def limpiar_fecha_robusta(valor, es_tulti):
    if pd.isna(valor) or str(valor).strip() == "": return pd.NaT
    valor_str = str(valor).lower().strip()
    reemplazos = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06',
                  'julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12',
                  ' de ':'/',' del ':'/',',':''}
    for d in ['lunes','martes','miércoles','miercoles','jueves','viernes','sábado','sabado','domingo']: valor_str = valor_str.replace(d, '')
    for txt, num in reemplazos.items(): valor_str = valor_str.replace(txt, str(num))
    valor_str = re.sub(r'\s+', ' ', valor_str).strip().replace('-', '/')
    try:
        return pd.to_datetime(valor_str, dayfirst=not es_tulti)
    except: return pd.NaT

# --- LA INTERFAZ Y LÓGICA QUE SE LLAMA DESDE APP.PY ---
def render():
    st.title("🚀 Auto-Limpieza y Cruce (End-to-End)")
    st.markdown("Sube tus bases **sucias** y pega el link de Ventas Master. El sistema limpiará, filtrará el 2026 y generará la base final para Power BI.")
    
    with st.expander("1️⃣ Carga de Archivos Sucios (BPro y Drive)", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**BPro Compras**")
            f_cc = st.file_uploader("Compras Cuautitlán", type=["xls", "xlsx"], key="e2e_cc")
            f_ct = st.file_uploader("Compras Tultitlán", type=["xls", "xlsx"], key="e2e_ct")
        with c2:
            st.markdown("**BPro Traspasos**")
            f_tc = st.file_uploader("Traspasos Cuautitlán", type=["xls", "xlsx"], key="e2e_tc")
            f_tt = st.file_uploader("Traspasos Tultitlán", type=["xls", "xlsx"], key="e2e_tt")
        with c3:
            st.markdown("**Drive Solicitudes (CSV)**")
            f_dc = st.file_uploader("Drive Cuautitlán", type=["csv"], key="e2e_dc")
            f_dt = st.file_uploader("Drive Tultitlán", type=["csv"], key="e2e_dt")

    with st.expander("2️⃣ Conexión a Ventas Master (Nube)", expanded=True):
        url_ventas = st.text_input("🔗 Pega el enlace de compartir de Google Drive (Ventas Master):")

    if st.button("⚙️ Ejecutar Magia (Limpiar y Cruzar)", type="primary"):
        if f_cc and f_ct and f_dc and f_dt and url_ventas:
            with st.spinner("Procesando y cruzando bases de datos..."):
                try:
                    # Fase 1: Compras
                    df_compras = pd.concat([procesar_compras(f_cc, "CUAUTITLAN"), procesar_compras(f_ct, "TULTITLAN")], ignore_index=True)
                    df_compras['FECHA'] = pd.to_datetime(df_compras['FECHA'], dayfirst=True, errors='coerce')
                    
                    # Fase 2: Traspasos
                    dfs_trasp = []
                    if f_tc: dfs_trasp.append(procesar_traspasos(f_tc, "CUAUTITLAN"))
                    if f_tt: dfs_trasp.append(procesar_traspasos(f_tt, "TULTITLAN"))
                    df_traspasos = pd.concat(dfs_trasp, ignore_index=True) if dfs_trasp else pd.DataFrame(columns=['AGENCIA', 'NP', 'CANTIDAD'])

                    # Fase 3: Drive
                    dfs_sol = []
                    df_c = pd.read_csv(f_dc, header=1, encoding='latin1')
                    df_c.columns = df_c.columns.str.strip() 
                    if 'CANCELAR (X)' in df_c.columns: df_c = df_c[df_c['CANCELAR (X)'].isna()]
                    df_c = df_c.rename(columns={'Fecha': 'FECHA_SOLICITUD', 'Vendedor': 'VENDEDOR', 'No. De Parte': 'NP', 'Descripcion': 'DESCRIPCION', 'Cantidad': 'CANTIDAD', 'Orden de Compra': 'ORDEN_COMPRA'})
                    df_c['AGENCIA'] = 'CUAUTITLAN'
                    df_c['FECHA_SOLICITUD_DT'] = df_c['FECHA_SOLICITUD'].apply(lambda x: limpiar_fecha_robusta(x, es_tulti=False))
                    dfs_sol.append(df_c)
                    
                    df_t = pd.read_csv(f_dt, header=6, encoding='latin1')
                    df_t.columns = df_t.columns.str.strip()
                    df_t = df_t.rename(columns={'Fecha': 'FECHA_SOLICITUD', 'Vendedor': 'VENDEDOR', 'No. De Parte': 'NP', 'Descripcion': 'DESCRIPCION', 'Cantidad': 'CANTIDAD', 'Observaciones': 'ORDEN_COMPRA'})
                    df_t['AGENCIA'] = 'TULTITLAN'
                    df_t['FECHA_SOLICITUD_DT'] = df_t['FECHA_SOLICITUD'].apply(lambda x: limpiar_fecha_robusta(x, es_tulti=True))
                    dfs_sol.append(df_t)
                    
                    df_drive = pd.concat(dfs_sol, ignore_index=True)
                    for col in ['FECHA_SOLICITUD_DT', 'VENDEDOR', 'NP']: df_drive[col] = df_drive[col].replace(r'^\s*$', np.nan, regex=True)
                    df_drive = df_drive.dropna(subset=['FECHA_SOLICITUD_DT', 'VENDEDOR', 'NP'])
                    df_drive = df_drive[df_drive['FECHA_SOLICITUD_DT'].dt.year == 2026]

                    # Fase 4: Ventas Master
                    st.toast('Descargando Ventas Master...', icon='☁️')
                    link_directo = obtener_enlace_directo_drive(url_ventas)
                    df_ventas = pd.read_csv(link_directo, encoding='latin1') if "csv" in url_ventas.lower() else pd.read_excel(link_directo)
                    df_ventas['FECHA'] = pd.to_datetime(df_ventas['FECHA'], dayfirst=True, errors='coerce')
                    df_ventas = df_ventas[df_ventas['FECHA'].dt.year == 2026]
                    df_ventas['ALMACEN_NORM'] = df_ventas['ALMACEN'].astype(str).str.upper()
                    df_ventas['AGENCIA'] = df_ventas['ALMACEN_NORM'].apply(lambda x: 'CUAUTITLAN' if 'CUAUTI' in x else ('TULTITLAN' if 'TULTI' in x else x))

                    # Fase 5: Cruce
                    st.toast('Generando Análisis...', icon='🔗')
                    agg_com = df_compras.groupby(['AGENCIA', 'NP']).agg({'CANTIDAD': 'sum', 'DESCRIPCION': 'first', 'FECHA': 'max'}).reset_index().rename(columns={'CANTIDAD': 'COMPRADO', 'FECHA': 'ULT_COMPRA'})
                    agg_tra = df_traspasos.groupby(['AGENCIA', 'NP']).agg({'CANTIDAD': 'sum'}).reset_index().rename(columns={'CANTIDAD': 'TRASPASADO'})
                    agg_ven = df_ventas.groupby(['AGENCIA', 'NP']).agg({'CANTIDAD': 'sum', 'FECHA': 'max'}).reset_index().rename(columns={'CANTIDAD': 'VENDIDO', 'FECHA': 'ULT_VENTA'})
                    
                    hoja_gral = pd.merge(agg_com, agg_tra, on=['AGENCIA', 'NP'], how='left')
                    hoja_gral = pd.merge(hoja_gral, agg_ven, on=['AGENCIA', 'NP'], how='left').fillna(0)
                    hoja_gral['ULT_COMPRA'] = hoja_gral['ULT_COMPRA'].dt.strftime('%d/%m/%Y').replace('NaT', 'Sin Fecha')
                    hoja_gral['ULT_VENTA'] = pd.to_datetime(hoja_gral['ULT_VENTA']).dt.strftime('%d/%m/%Y').replace('NaT', 'Sin Venta')
                    hoja_gral = hoja_gral[['AGENCIA', 'NP', 'DESCRIPCION', 'COMPRADO', 'TRASPASADO', 'VENDIDO', 'ULT_COMPRA', 'ULT_VENTA']]
                    
                    analisis = []
                    for _, row in df_drive.iterrows():
                        agencia, np_val, fecha_sol, cant_pedida = row['AGENCIA'], row['NP'], row['FECHA_SOLICITUD_DT'], row['CANTIDAD']
                        ventas_post = df_ventas[(df_ventas['AGENCIA'] == agencia) & (df_ventas['NP'] == np_val) & (df_ventas['FECHA'] >= fecha_sol)]
                        total_vendido = ventas_post['CANTIDAD'].sum()
                        
                        if cant_pedida - total_vendido > 0:
                            fila = row.drop('FECHA_SOLICITUD_DT').to_dict()
                            fila['FECHA_SOLICITUD'] = fecha_sol.strftime('%d/%m/%Y')
                            fila['CANTIDAD_VENDIDA'] = total_vendido
                            fila['SOBRANTE'] = cant_pedida - total_vendido
                            analisis.append(fila)
                            
                    hoja_drive = pd.DataFrame(analisis)
                    if not hoja_drive.empty:
                        hoja_drive = hoja_drive[['AGENCIA', 'FECHA_SOLICITUD', 'VENDEDOR', 'NP', 'DESCRIPCION', 'CANTIDAD', 'CANTIDAD_VENDIDA', 'SOBRANTE', 'ORDEN_COMPRA']]
                        hoja_drive = hoja_drive.sort_values(by='SOBRANTE', ascending=False)
                    
                    # Generar Excel
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
                        hoja_gral.to_excel(w, sheet_name='General_Compras_Ventas', index=False)
                        if not hoja_drive.empty: hoja_drive.to_excel(w, sheet_name='Drive_No_Vendido', index=False)
                        else: pd.DataFrame({'Msg': ['Todo vendido']}).to_excel(w, sheet_name='Drive_No_Vendido', index=False)
                    
                    st.success("¡Análisis Completado!")
                    st.balloons()
                    st.download_button("⬇️ Descargar Reporte Final Power BI", buf.getvalue(), "Base_Final_PowerBI.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.warning("⚠️ Sube todos los archivos de BPro, Drive y pega el link de Ventas Master.")
