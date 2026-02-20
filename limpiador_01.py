import streamlit as st
import pandas as pd
import io
import re
from babel.dates import get_month_names

# --- HELPER: VALIDACIÓN SEGURA ---
def es_dataframe_valido(df):
    return isinstance(df, pd.DataFrame) and not df.empty

# --- PARSERS VIEJOS (MULTI-HOJA) ---
@st.cache_data
def procesar_compras(file_content, nomenclatura):
    try:
        xls = pd.ExcelFile(file_content)
        compras_list = []
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                fecha_actual = pd.NaT
                for row in df.itertuples(index=False, name=None):
                    celda_a = str(row[0] or '').strip().upper()
                    celda_c = str(row[2] or '').strip()
                    if "FACTURA:" in celda_a:
                        match = re.search(r'(\d{2}/\d{2}/\d{4})', celda_c)
                        if match: fecha_actual = pd.to_datetime(match.group(1), format='%d/%m/%Y', errors='coerce')
                        continue
                    if nomenclatura in celda_a:
                        try:
                            id_part = row[2]
                            description = row[3]
                            product_line = row[11]
                            cantidad = pd.to_numeric(row[4], errors='coerce')
                            total = pd.to_numeric(row[7], errors='coerce')
                            if pd.notna(id_part) and id_part != 0 and pd.notna(cantidad) and pd.notna(fecha_actual):
                                compras_list.append({'ID PART': id_part, 'DESCRIPTION': description, 'PRODUCT LINE': product_line, 'CANTIDAD COMPRADA': cantidad, 'TOTAL COMPRADO': total, 'Fecha': fecha_actual})
                        except: continue
            except: continue
        return pd.DataFrame(compras_list).dropna(subset=['Fecha']) if compras_list else pd.DataFrame()
    except Exception as e: 
        st.error(f"Error global ({nomenclatura}): {e}"); return pd.DataFrame()

@st.cache_data
def parsear_traspasos_detallado(file_content, nomenclaturas_agencia):
    try:
        xls = pd.ExcelFile(file_content)
        traspasos_combinados = {}
        for sheet_name in xls.sheet_names:
            try:
                df_crudo = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                lista_destinos = []
                limit_search = min(50, len(df_crudo)) 
                for index in range(4, limit_search):
                    celda_a = str(df_crudo.iloc[index, 0]).strip()
                    if celda_a.upper() == "TOTALES": break
                    if celda_a: lista_destinos.append(celda_a)
                
                if not lista_destinos: continue 

                destino_actual = None
                nombre_limpio_destino = None
                fecha_actual = pd.NaT
                
                for index, row in df_crudo.iterrows():
                    celda_a = str(row.get(0, '')).strip()
                    celda_c = str(row.get(2, '')).strip()
                    if not celda_a: continue
                    if celda_a.upper().startswith("REFERENCIA:"):
                        match = re.search(r'(\d{2}/\d{2}/\d{4})', celda_c)
                        if match: fecha_actual = pd.to_datetime(match.group(1), format='%d/%m/%Y', errors='coerce')
                        continue
                    if celda_a in lista_destinos:
                        destino_actual = celda_a
                        nombre_limpio_destino = re.split(r'hacia', destino_actual, flags=re.IGNORECASE)[-1].strip() if "hacia" in destino_actual.lower() else destino_actual
                        continue
                    if destino_actual and any(nom in celda_a for nom in nomenclaturas_agencia):
                        id_part = row.get(2)
                        cant = pd.to_numeric(row.get(4), errors='coerce')
                        if pd.notna(id_part) and pd.notna(cant) and id_part != 0 and pd.notna(fecha_actual):
                            traspasos_combinados.setdefault(nombre_limpio_destino, []).append({'ID PART': id_part, 'Cantidad Traspasada': abs(cant), 'Fecha': fecha_actual})
            except: continue
        return {d: pd.DataFrame(r).dropna(subset=['Fecha']) for d, r in traspasos_combinados.items() if r}
    except Exception as e: st.error(f"Error traspasos: {e}"); return {}

@st.cache_data
def procesar_archivo_venta_individual(file_content, nomenclaturas):
    try:
        xls = pd.ExcelFile(file_content)
        ventas_list = []
        for sheet_name in xls.sheet_names:
            try:
                df_crudo = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                fecha_actual = pd.NaT
                for row in df_crudo.itertuples(index=False, name=None):
                    celda_a = str(row[0] or '').strip().upper()
                    celda_e = str(row[4] or '').strip()
                    if "FACTURA/REFERENCIA:" in celda_a:
                        match = re.search(r'(\d{2}/\d{2}/\d{4})', celda_e)
                        if match: fecha_actual = pd.to_datetime(match.group(1), format='%d/%m/%Y', errors='coerce')
                        continue
                    if any(nom in celda_a for nom in nomenclaturas):
                        try:
                            id_part = row[2]
                            cantidad = pd.to_numeric(row[4], errors='coerce')
                            total = pd.to_numeric(row[6], errors='coerce')
                            if pd.notna(id_part) and id_part != 0 and pd.notna(cantidad) and pd.notna(fecha_actual):
                                ventas_list.append({'ID PART': id_part, 'Cantidad Vendida': cantidad, 'Total Vendido': total, 'Fecha': fecha_actual})
                        except: continue
            except: continue
        return pd.DataFrame(ventas_list).dropna(subset=['Fecha']) if ventas_list else pd.DataFrame()
    except Exception as e: st.error(f"Error ventas: {e}"); return pd.DataFrame()

# --- AGREGACIÓN ---
def agregar_compras(df_raw):
    if not es_dataframe_valido(df_raw): return pd.DataFrame()
    agg = df_raw.groupby('ID PART').agg({'CANTIDAD COMPRADA': 'sum', 'TOTAL COMPRADO': 'sum', 'DESCRIPTION': 'first', 'PRODUCT LINE': 'first', 'Fecha': 'max'}).reset_index()
    agg.rename(columns={'Fecha': 'Fecha Ult. Comp.'}, inplace=True)
    agg['Fecha Ult. Comp.'] = agg['Fecha Ult. Comp.'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else 'N/A')
    return agg

def agregar_datos_simples(df_raw, col_cantidad, col_total=None):
    if not es_dataframe_valido(df_raw): return pd.DataFrame()
    agg_dict = {col_cantidad: 'sum'}
    if col_total and col_total in df_raw.columns: agg_dict[col_total] = 'sum'
    return df_raw.groupby('ID PART').agg(agg_dict).reset_index()

def agregar_dict_datos(dict_raw, col_cantidad, col_total=None):
    if not dict_raw: return {}
    dict_agg = {}
    for key, df_raw in dict_raw.items():
        if es_dataframe_valido(df_raw): dict_agg[key] = agregar_datos_simples(df_raw, col_cantidad, col_total)
    return dict_agg

# --- GENERACIÓN DE REPORTES (CORE VIEJO) ---
def generar_reporte_agencia(df_compras, traspasos_data, seleccion_almacenes, ventas_gral, ventas_manuales):
    if not es_dataframe_valido(df_compras): return pd.DataFrame()
    df = df_compras.copy()
    ids_por_analizar = set()
    
    all_trasp = [d for d in traspasos_data.values() if es_dataframe_valido(d)]
    if all_trasp:
        agg = pd.concat(all_trasp).groupby('ID PART')['Cantidad Traspasada'].sum().reset_index().rename(columns={'Cantidad Traspasada': 'TOTAL TRASPASOS'})
        df = pd.merge(df, agg, on='ID PART', how='left')
    else: df['TOTAL TRASPASOS'] = 0
    df['TOTAL TRASPASOS'] = df['TOTAL TRASPASOS'].fillna(0)

    ignorar_list = []
    if es_dataframe_valido(seleccion_almacenes):
        lista_no = seleccion_almacenes[seleccion_almacenes['Acción'] == 'No Considerar']['Almacén Destino'].tolist()
        ignorar_list = [traspasos_data.get(a) for a in lista_no if traspasos_data.get(a) is not None]
    valid_ignorar = [x for x in ignorar_list if es_dataframe_valido(x)]
    if valid_ignorar:
        ign_agg = pd.concat(valid_ignorar).groupby('ID PART')['Cantidad Traspasada'].sum().reset_index().rename(columns={'Cantidad Traspasada': 'Ignorada'})
        df = pd.merge(df, ign_agg, on='ID PART', how='left').fillna(0)
        costo = (df['TOTAL COMPRADO']/df['CANTIDAD COMPRADA']).replace([float('inf'), -float('inf')], 0).fillna(0)
        df['CANTIDAD COMPRADA'] -= df['Ignorada']
        df['TOTAL COMPRADO'] = df['CANTIDAD COMPRADA'] * costo
        df.drop(columns=['Ignorada'], inplace=True)
    df = df[df['CANTIDAD COMPRADA'] > 0].copy()
    if df.empty: return pd.DataFrame()

    if es_dataframe_valido(ventas_gral):
        df = pd.merge(df, ventas_gral, on="ID PART", how="left")
        df.rename(columns={"Cantidad Vendida": "ALMACEN GENERAL_Venta Directa", "Total Vendido": "ALMACEN GENERAL_Total Vendido"}, inplace=True)
    else: df["ALMACEN GENERAL_Venta Directa"], df["ALMACEN GENERAL_Total Vendido"] = 0, 0

    if es_dataframe_valido(seleccion_almacenes):
        for _, row in seleccion_almacenes[seleccion_almacenes['Acción'] != 'No Considerar'].iterrows():
            alm, accion = row['Almacén Destino'], row['Acción']
            t_df = traspasos_data.get(alm)
            if not es_dataframe_valido(t_df): continue
            temp = pd.merge(df[['ID PART']], t_df, on="ID PART", how="left").fillna(0)
            
            if accion == "Venta Exitosa":
                temp['Cantidad Vendida'] = temp['Cantidad Traspasada']
                costo = pd.merge(temp[['ID PART']], df[['ID PART', 'TOTAL COMPRADO', 'CANTIDAD COMPRADA']], on='ID PART', how='left')
                u_cost = (costo['TOTAL COMPRADO']/costo['CANTIDAD COMPRADA']).fillna(0).replace([float('inf'), -float('inf')], 0)
                temp['Total Vendido'] = temp['Cantidad Vendida'] * u_cost
            elif accion == "Considerar":
                v_df = ventas_manuales.get(alm)
                if es_dataframe_valido(v_df):
                    temp = pd.merge(temp, v_df, on="ID PART", how="left").fillna(0)
                    precio = (temp['Total Vendido']/temp['Cantidad Vendida']).fillna(0).replace([float('inf'), -float('inf')], 0)
                    temp['Cantidad Vendida'] = temp.apply(lambda r: min(r['Cantidad Vendida'], r['Cantidad Traspasada']), axis=1)
                    temp['Total Vendido'] = temp['Cantidad Vendida'] * precio
                else: temp['Cantidad Vendida'], temp['Total Vendido'] = 0, 0
            elif accion == "Por analizar":
                ids_por_analizar.update(t_df[t_df['Cantidad Traspasada']>0]['ID PART'].unique())
                temp['Cantidad Vendida'], temp['Total Vendido'] = 0, 0
            
            if not temp.empty:
                temp.rename(columns={"Cantidad Traspasada": f"{alm}_Traspasada", "Cantidad Vendida": f"{alm}_Vendida", "Total Vendido": f"{alm}_Total Vendido"}, inplace=True)
                df = pd.merge(df, temp.drop(columns=['Fecha'], errors='ignore'), on="ID PART", how="left")

    df.fillna(0, inplace=True)
    c_v = [c for c in df.columns if '_Vendida' in c or '_Venta Directa' in c]
    c_t = [c for c in df.columns if '_Total Vendido' in c]
    df['CANTIDAD VENDIDA'] = df[c_v].sum(axis=1)
    df['TOTAL VENDIDO'] = df[c_t].sum(axis=1)
    if ids_por_analizar:
        df['CANTIDAD VENDIDA'] = df['CANTIDAD VENDIDA'].astype(object)
        df.loc[df['ID PART'].isin(ids_por_analizar), 'CANTIDAD VENDIDA'] = "Por analizar"
    return df

def escribir_excel(writer, df, hoja):
    if not es_dataframe_valido(df): return
    base = ['ID PART', 'DESCRIPTION', 'PRODUCT LINE', 'TOTAL COMPRADO', 'CANTIDAD COMPRADA', 'CANTIDAD VENDIDA', 'TOTAL TRASPASOS', 'TOTAL VENDIDO', 'Fecha Ult. Comp.']
    for c in base: 
        if c not in df.columns: df[c] = 0
    
    gral = sorted([c for c in df.columns if 'ALMACEN GENERAL' in c])
    otros = sorted([c for c in df.columns if c not in base and c not in gral])
    df = df.reindex(columns=base + gral + otros, fill_value=0)
    
    df_main = df[base]
    df_rest = df.drop(columns=base)
    cols = []
    if not df_rest.empty:
        for c in df_rest.columns:
            p = c.split('_')
            cols.append(('_'.join(p[:-1]), p[-1]) if len(p)>=2 else (c, ''))
        df_rest.columns = pd.MultiIndex.from_tuples(cols)
        pd.concat([df_main, df_rest], axis=1).to_excel(writer, sheet_name=hoja, index=False)
    else: df_main.to_excel(writer, sheet_name=hoja, index=False)

def generar_df_remanentes(df_compras_raw, df_ventas_gral, traspasos_dict, manuales_dict, df_config_almacenes):
    if not es_dataframe_valido(df_compras_raw): return pd.DataFrame()
    mapa_acciones = {}
    if es_dataframe_valido(df_config_almacenes):
        for _, r in df_config_almacenes.iterrows():
            mapa_acciones[str(r['Almacén Destino']).strip()] = str(r['Acción']).strip()

    total_salidas = {}
    if es_dataframe_valido(df_ventas_gral):
        s = df_ventas_gral.groupby('ID PART')['Cantidad Vendida'].sum()
        for pid, cant in s.items(): total_salidas[pid] = total_salidas.get(pid, 0) + cant
    
    for nombre_almacen, t_df in traspasos_dict.items():
        if not es_dataframe_valido(t_df): continue
        accion = mapa_acciones.get(nombre_almacen, "Considerar") 
        if accion == "Venta Exitosa":
            s = t_df.groupby('ID PART')['Cantidad Traspasada'].sum()
            for pid, cant in s.items(): total_salidas[pid] = total_salidas.get(pid, 0) + cant
        elif accion == "Considerar":
            v_manual = manuales_dict.get(nombre_almacen)
            if es_dataframe_valido(v_manual):
                s = v_manual.groupby('ID PART')['Cantidad Vendida'].sum()
                for pid, cant in s.items(): total_salidas[pid] = total_salidas.get(pid, 0) + cant

    df_compras = df_compras_raw.copy()
    df_compras['Periodo'] = df_compras['Fecha'].dt.to_period('M')
    
    compras_mensuales = df_compras.groupby(['ID PART', 'Periodo']).agg({
        'CANTIDAD COMPRADA': 'sum',   
        'TOTAL COMPRADO': 'sum',
        'DESCRIPTION': 'first',
        'PRODUCT LINE': 'first',
        'Fecha': 'max'
    }).reset_index()
    
    compras_mensuales = compras_mensuales.sort_values(by=['Periodo'])
    filas_remanentes = []
    
    for row in compras_mensuales.to_dict('records'):
        pid = row['ID PART']
        qty_mes = row['CANTIDAD COMPRADA'] 
        costo_total_mes = row['TOTAL COMPRADO']
        costo_unit = costo_total_mes / qty_mes if qty_mes > 0 else 0
        
        salidas_acumuladas = total_salidas.get(pid, 0)
        if salidas_acumuladas >= qty_mes:
            total_salidas[pid] -= qty_mes
        else:
            residuo = qty_mes - salidas_acumuladas
            total_salidas[pid] = 0 
            row['CANTIDAD COMPRADA'] = residuo
            row['TOTAL COMPRADO'] = residuo * costo_unit
            row['CANTIDAD VENDIDA'] = 0
            row['TOTAL TRASPASOS'] = 0
            row['TOTAL VENDIDO'] = 0
            filas_remanentes.append(row)
    
    return pd.DataFrame(filas_remanentes)

# --- LA INTERFAZ QUE SE LLAMA DESDE APP.PY ---
def render():
    st.title("📊 Análisis Integrado Viejo (Lógica FIFO)")
    
    if 'init_main_app' not in st.session_state:
        st.session_state.update({
            'init_main_app': True,
            'df_compras_cua_raw': pd.DataFrame(), 'df_compras_tul_raw': pd.DataFrame(),
            'traspasos_cua_data_raw': {}, 'traspasos_tul_data_raw': {},
            'ventas_gral_cua_raw': pd.DataFrame(), 'ventas_gral_tul_raw': pd.DataFrame(),
            'ventas_manuales_raw': {}, 'reporte_final_bytes': None,
            'base_almacenes_cua': pd.DataFrame(), 'base_almacenes_tul': pd.DataFrame(), 
            'final_almacenes_cua': pd.DataFrame(), 'final_almacenes_tul': pd.DataFrame(),
            'show_balloons': False,
            'last_id_cua': None, 'last_id_tul': None
        })

    with st.expander("✅ PASO 1: Cargar Archivos de Compras", expanded=True):
        col1, col2 = st.columns(2)
        c_cua = col1.file_uploader("📂 Compras **Cuautitlán**", type=['xlsx', 'xls'])
        if c_cua: st.session_state.df_compras_cua_raw = procesar_compras(c_cua, "CRCU")
        if es_dataframe_valido(st.session_state.df_compras_cua_raw): col1.success(f"Cuautitlán: {len(st.session_state.df_compras_cua_raw)} items.")

        c_tul = col2.file_uploader("📂 Compras **Tultitlán**", type=['xlsx', 'xls'])
        if c_tul: st.session_state.df_compras_tul_raw = procesar_compras(c_tul, "CRTU")
        if es_dataframe_valido(st.session_state.df_compras_tul_raw): col2.success(f"Tultitlán: {len(st.session_state.df_compras_tul_raw)} items.")

    with st.expander("✅ PASO 2: Cargar Traspasos y Definir Almacenes"):
        col3, col4 = st.columns(2)
        with col3:
            st.subheader("Traspasos Cuautitlán")
            trasp_cua = st.file_uploader("📂 Sube traspasos **Cuautitlán**", type=['xlsx', 'xls'], key="up_traspasos_cua")
            if trasp_cua:
                fid = f"{trasp_cua.name}_{trasp_cua.size}"
                if st.session_state.last_id_cua != fid:
                    raw = parsear_traspasos_detallado(trasp_cua, ["TRASUCCU", "TRASAPROCU"])
                    st.session_state.traspasos_cua_data_raw = raw
                    st.session_state.last_id_cua = fid
                    if raw:
                        nuevos = sorted(list(raw.keys()))
                        st.session_state.base_almacenes_cua = pd.DataFrame({"Almacén Destino": nuevos, "Acción": ["Considerar"] * len(nuevos)})
                        st.session_state.final_almacenes_cua = st.session_state.base_almacenes_cua.copy()
            if st.session_state.traspasos_cua_data_raw:
                col3.info(f"Se encontraron {len(st.session_state.traspasos_cua_data_raw)} almacenes.")
                if es_dataframe_valido(st.session_state.base_almacenes_cua):
                    edited_cua = st.data_editor(st.session_state.base_almacenes_cua, column_config={"Acción": st.column_config.SelectboxColumn("Acción", options=["Considerar", "Venta Exitosa", "Por analizar", "No Considerar"], required=True)}, hide_index=True, key="editor_cua_safe", use_container_width=True)
                    st.session_state.final_almacenes_cua = edited_cua

        with col4:
            st.subheader("Traspasos Tultitlán")
            trasp_tul = st.file_uploader("📂 Sube traspasos **Tultitlán**", type=['xlsx', 'xls'], key="up_traspasos_tul")
            if trasp_tul:
                fid = f"{trasp_tul.name}_{trasp_tul.size}"
                if st.session_state.last_id_tul != fid:
                    raw = parsear_traspasos_detallado(trasp_tul, ["TRASUCTU", "TRASAPROTU"])
                    st.session_state.traspasos_tul_data_raw = raw
                    st.session_state.last_id_tul = fid
                    if raw:
                        nuevos = sorted(list(raw.keys()))
                        st.session_state.base_almacenes_tul = pd.DataFrame({"Almacén Destino": nuevos, "Acción": ["Considerar"] * len(nuevos)})
                        st.session_state.final_almacenes_tul = st.session_state.base_almacenes_tul.copy()
            if st.session_state.traspasos_tul_data_raw:
                col4.info(f"Se encontraron {len(st.session_state.traspasos_tul_data_raw)} almacenes.")
                if es_dataframe_valido(st.session_state.base_almacenes_tul):
                    edited_tul = st.data_editor(st.session_state.base_almacenes_tul, column_config={"Acción": st.column_config.SelectboxColumn("Acción", options=["Considerar", "Venta Exitosa", "Por analizar", "No Considerar"], required=True)}, hide_index=True, key="editor_tul_safe", use_container_width=True)
                    st.session_state.final_almacenes_tul = edited_tul

    with st.expander("✅ PASO 3: Cargar Ventas Directas (Almacén General)"):
        c5, c6 = st.columns(2)
        v_cua = c5.file_uploader("📦 Ventas **Cuautitlán**", type=['xlsx', 'xls'])
        if v_cua: st.session_state.ventas_gral_cua_raw = procesar_archivo_venta_individual(v_cua, ["VRCU"])
        if es_dataframe_valido(st.session_state.ventas_gral_cua_raw): c5.success("OK Cuautitlán.")

        v_tul = c6.file_uploader("📦 Ventas **Tultitlán**", type=['xlsx', 'xls'])
        if v_tul: st.session_state.ventas_gral_tul_raw = procesar_archivo_venta_individual(v_tul, ["VRTU"])
        if es_dataframe_valido(st.session_state.ventas_gral_tul_raw): c6.success("OK Tultitlán.")

    use_cua, use_tul = st.session_state.final_almacenes_cua, st.session_state.final_almacenes_tul
    if es_dataframe_valido(use_cua) or es_dataframe_valido(use_tul):
        with st.expander("✅ PASO 4: Cargar Ventas por Traspaso (Manual)", expanded=True):
            st.info("Arrastra los archivos para los almacenes marcados como 'Considerar'.")
            lista_cua = use_cua[use_cua['Acción'] == 'Considerar']['Almacén Destino'].tolist() if es_dataframe_valido(use_cua) else []
            lista_tul = use_tul[use_tul['Acción'] == 'Considerar']['Almacén Destino'].tolist() if es_dataframe_valido(use_tul) else []
            total_req = len(lista_cua) + len(lista_tul)
            c7, c8 = st.columns(2)
            with c7:
                if lista_cua: st.subheader("Agencia Cuautitlán")
                for alm in lista_cua:
                    f = st.file_uploader(f"📂 Venta para: {alm}", key=f"m_c_{alm}")
                    if f: st.session_state.ventas_manuales_raw[alm] = procesar_archivo_venta_individual(f, ["VRCU", "VRTU"])
            with c8:
                if lista_tul: st.subheader("Agencia Tultitlán")
                for alm in lista_tul:
                    f = st.file_uploader(f"📂 Venta para: {alm}", key=f"m_t_{alm}")
                    if f: st.session_state.ventas_manuales_raw[alm] = procesar_archivo_venta_individual(f, ["VRCU", "VRTU"])
            cargados = sum(1 for a in lista_cua + lista_tul if a in st.session_state.ventas_manuales_raw and es_dataframe_valido(st.session_state.ventas_manuales_raw[a]))
            if total_req > 0:
                st.progress(min(cargados/total_req, 1.0), text=f"Archivos: {cargados}/{total_req}")
                if cargados >= total_req: st.success("¡Listo para generar!")

    st.divider()

    listos = all([
        es_dataframe_valido(st.session_state.df_compras_cua_raw),
        es_dataframe_valido(st.session_state.df_compras_tul_raw),
        st.session_state.traspasos_cua_data_raw,
        st.session_state.traspasos_tul_data_raw
    ])

    col_g, col_m, col_sv = st.columns(3)
    
    with col_g:
        if st.button("🚀 Reporte General", type="primary", use_container_width=True, disabled=not listos):
            with st.spinner("Procesando..."):
                c_c = agregar_compras(st.session_state.df_compras_cua_raw)
                c_t = agregar_compras(st.session_state.df_compras_tul_raw)
                t_c = agregar_dict_datos(st.session_state.traspasos_cua_data_raw, 'Cantidad Traspasada')
                t_t = agregar_dict_datos(st.session_state.traspasos_tul_data_raw, 'Cantidad Traspasada')
                v_g_c = agregar_datos_simples(st.session_state.ventas_gral_cua_raw, 'Cantidad Vendida', 'Total Vendido')
                v_g_t = agregar_datos_simples(st.session_state.ventas_gral_tul_raw, 'Cantidad Vendida', 'Total Vendido')
                v_m = agregar_dict_datos(st.session_state.ventas_manuales_raw, 'Cantidad Vendida', 'Total Vendido')
                fin_c = generar_reporte_agencia(c_c, t_c, st.session_state.final_almacenes_cua, v_g_c, v_m)
                fin_t = generar_reporte_agencia(c_t, t_t, st.session_state.final_almacenes_tul, v_g_t, v_m)
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    escribir_excel(w, fin_c, "Detalle_Cuautitlan")
                    escribir_excel(w, fin_t, "Detalle_Tultitlan")
                st.session_state.reporte_final_bytes = out.getvalue()
                st.session_state.show_balloons = True

    with col_m:
        if st.button("📅 Reporte Mensual", type="secondary", use_container_width=True, disabled=not listos):
            with st.spinner("Procesando..."):
                dates = []
                if es_dataframe_valido(st.session_state.df_compras_cua_raw): dates.append(st.session_state.df_compras_cua_raw['Fecha'])
                if es_dataframe_valido(st.session_state.df_compras_tul_raw): dates.append(st.session_state.df_compras_tul_raw['Fecha'])
                for d in st.session_state.traspasos_cua_data_raw.values(): dates.append(d['Fecha'])
                if dates:
                    months = sorted(pd.concat(dates).dt.to_period('M').unique())
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                        for p in months:
                            name = f"{get_month_names('wide', locale='es_ES')[p.month].capitalize()}_{p.year}"
                            cc = st.session_state.df_compras_cua_raw
                            cc = cc[cc['Fecha'].dt.to_period('M')==p] if es_dataframe_valido(cc) else pd.DataFrame()
                            ct = st.session_state.df_compras_tul_raw
                            ct = ct[ct['Fecha'].dt.to_period('M')==p] if es_dataframe_valido(ct) else pd.DataFrame()
                            tc = {k:v[v['Fecha'].dt.to_period('M')==p] for k,v in st.session_state.traspasos_cua_data_raw.items()}
                            tt = {k:v[v['Fecha'].dt.to_period('M')==p] for k,v in st.session_state.traspasos_tul_data_raw.items()}
                            vc = st.session_state.ventas_gral_cua_raw
                            vc = vc[vc['Fecha'].dt.to_period('M')==p] if es_dataframe_valido(vc) else pd.DataFrame()
                            vt = st.session_state.ventas_gral_tul_raw
                            vt = vt[vt['Fecha'].dt.to_period('M')==p] if es_dataframe_valido(vt) else pd.DataFrame()
                            vm = {k:v[v['Fecha'].dt.to_period('M')==p] for k,v in st.session_state.ventas_manuales_raw.items()}
                            fc = generar_reporte_agencia(agregar_compras(cc), agregar_dict_datos(tc,'Cantidad Traspasada'), st.session_state.final_almacenes_cua, agregar_datos_simples(vc,'Cantidad Vendida','Total Vendido'), agregar_dict_datos(vm,'Cantidad Vendida','Total Vendido'))
                            ft = generar_reporte_agencia(agregar_compras(ct), agregar_dict_datos(tt,'Cantidad Traspasada'), st.session_state.final_almacenes_tul, agregar_datos_simples(vt,'Cantidad Vendida','Total Vendido'), agregar_dict_datos(vm,'Cantidad Vendida','Total Vendido'))
                            escribir_excel(w, fc, f"Cuautitlan_{name}")
                            escribir_excel(w, ft, f"Tultitlan_{name}")
                    st.session_state.reporte_final_bytes = out.getvalue()
                    st.session_state.show_balloons = True
                else: st.warning("No hay fechas.")

    with col_sv:
        if st.button("🚫 Compras sin Venta Exitosa", type="primary", use_container_width=True, disabled=not listos):
            with st.spinner("Calculando residuos de inventario..."):
                rem_cua = generar_df_remanentes(st.session_state.df_compras_cua_raw, st.session_state.ventas_gral_cua_raw, st.session_state.traspasos_cua_data_raw, st.session_state.ventas_manuales_raw, st.session_state.final_almacenes_cua)
                rem_tul = generar_df_remanentes(st.session_state.df_compras_tul_raw, st.session_state.ventas_gral_tul_raw, st.session_state.traspasos_tul_data_raw, st.session_state.ventas_manuales_raw, st.session_state.final_almacenes_tul)
                out = io.BytesIO()
                hay_datos = False
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    dates_rem = []
                    if not rem_cua.empty: dates_rem.append(rem_cua['Fecha'])
                    if not rem_tul.empty: dates_rem.append(rem_tul['Fecha'])
                    if dates_rem:
                        hay_datos = True
                        months_rem = sorted(pd.concat(dates_rem).dt.to_period('M').unique())
                        for p in months_rem:
                            name = f"{get_month_names('wide', locale='es_ES')[p.month].capitalize()}_{p.year}"
                            rc_mes = rem_cua[rem_cua['Fecha'].dt.to_period('M')==p].copy() if not rem_cua.empty else pd.DataFrame()
                            rt_mes = rem_tul[rem_tul['Fecha'].dt.to_period('M')==p].copy() if not rem_tul.empty else pd.DataFrame()
                            if not rc_mes.empty:
                                rc_mes['Fecha Ult. Comp.'] = rc_mes['Fecha'].dt.strftime('%d/%m/%Y')
                                rc_mes.drop(columns=['Fecha', 'Periodo'], inplace=True, errors='ignore')
                            if not rt_mes.empty:
                                rt_mes['Fecha Ult. Comp.'] = rt_mes['Fecha'].dt.strftime('%d/%m/%Y')
                                rt_mes.drop(columns=['Fecha', 'Periodo'], inplace=True, errors='ignore')
                            escribir_excel(w, rc_mes, f"Cuautitlan_{name}")
                            escribir_excel(w, rt_mes, f"Tultitlan_{name}")
                if hay_datos:
                    st.session_state.reporte_final_bytes = out.getvalue()
                    st.session_state.show_balloons = True
                else:
                    st.warning("¡Todo el inventario histórico ha sido vendido! No hay residuos.")

    if st.session_state.reporte_final_bytes:
        if st.session_state.show_balloons:
            st.balloons()
            st.session_state.show_balloons = False
        st.download_button("📥 Descargar Excel", st.session_state.reporte_final_bytes, "Reporte.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
