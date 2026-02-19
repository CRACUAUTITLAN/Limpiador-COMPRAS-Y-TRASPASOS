import streamlit as st
import pandas as pd
import io

# --- FUNCIÓN 1: LIMPIEZA DE COMPRAS (BPRO) ---
def procesar_compras(file, nombre_agencia):
    """
    Limpia el archivo de compras evitando filas basura o subtotales.
    """
    df = pd.read_excel(file, header=None)
    datos = []
    
    # Variables de estado
    factura, fecha, proveedor, comprador = None, None, None, None
    
    for _, row in df.iterrows():
        val_a = str(row[0]).strip()
        
        # Detectar Encabezados
        if "FACTURA:" in val_a:
            try:
                factura = val_a.split("FACTURA:")[1].strip()
                
                if "FECHA FACT:" in str(row[2]):
                    fecha = str(row[2]).split("FECHA FACT:")[1].strip()
                if "PROVEEDOR:" in str(row[3]):
                    proveedor = str(row[3]).split("PROVEEDOR:")[1].strip()
                if "COMPRADOR:" in str(row[4]):
                    comprador = str(row[4]).split("COMPRADOR:")[1].strip()
            except:
                continue

        # Detectar Ítems (CRCU / CRTU / o cualquier código de compra)
        elif val_a.startswith("CR"): 
            # --- CANDADO DE LIMPIEZA ---
            descripcion = str(row[3]).strip() # Columna D
            
            # Si la descripción existe y no es 'nan' (vacía), guardamos la fila
            if descripcion and descripcion.lower() != 'nan':
                datos.append({
                    "AGENCIA": nombre_agencia,
                    "FACTURA": factura,
                    "FECHA": fecha,
                    "PROVEEDOR": proveedor,
                    "COMPRADOR": comprador,
                    "NP": row[2],
                    "DESCRIPCION": row[3],
                    "CANTIDAD": row[4],
                    "COSTO_UNIT": row[5],
                    "TOTAL": row[9] 
                })
            
    return pd.DataFrame(datos)

# --- FUNCIÓN 2: LIMPIEZA DE TRASPASOS (BPRO) ---
def procesar_traspasos(file, nombre_agencia):
    """
    Limpia traspasos abarcando todas las variantes de salidas y evita subtotales finales.
    """
    df = pd.read_excel(file, header=None)
    datos = []
    
    # Variables de estado (Niveles jerárquicos)
    destino_actual = None
    referencia, fecha_mov, usuario = None, None, None
    
    for _, row in df.iterrows():
        val_a = str(row[0]).strip().upper()
        
        # --- NIVEL 1: DETECTAR DESTINO (SALIDA...) ---
        # Detecta cualquier variante que empiece con SALIDA
        if val_a.startswith("SALIDA"):
            
            # CASO A: Tiene un destino explícito con la palabra "HACIA"
            if "HACIA" in val_a:
                try:
                    destino_sucio = val_a.split("HACIA")[1] 
                    destino_actual = destino_sucio.strip()
                except:
                    continue
                    
            # CASO B: Es una salida por traspaso genérica (sin "HACIA")
            elif "SALIDA DE ALMACEN POR TRASPASO" in val_a:
                destino_actual = f"SALIDA DE ALMACEN POR TRASPASO {nombre_agencia}"

        # --- NIVEL 2: DETECTAR CABECERA (REFERENCIA / FECHA / USUARIO) ---
        elif "REFERENCIA:" in val_a:
            try:
                referencia = val_a.split("REFERENCIA:")[1].strip()
                
                if "FECHA MOV:" in str(row[2]).upper():
                    fecha_mov = str(row[2]).upper().split("FECHA MOV:")[1].strip()
                
                if "USUARIO:" in str(row[3]).upper():
                    usuario = str(row[3]).upper().split("USUARIO:")[1].strip()
            except:
                continue
                
        # --- NIVEL 3: DETECTAR ÍTEMS (TRAS...) ---
        elif val_a.startswith("TRAS"):
            if destino_actual:
                try:
                    # --- CANDADO DE LIMPIEZA ---
                    descripcion = str(row[3]).strip()
                    
                    if descripcion and descripcion.lower() != 'nan':
                        cantidad = float(row[4]) 
                        costo = float(row[5])    
                        
                        datos.append({
                            "AGENCIA": nombre_agencia,    
                            "DESTINO": destino_actual,    
                            "REFERENCIA": referencia,
                            "FECHA_MOV": fecha_mov,
                            "USUARIO": usuario,
                            "NP": row[2],                 
                            "DESCRIPCION": row[3],        
                            "CANTIDAD": abs(cantidad),    
                            "COSTO_UNIT": costo,          
                            "TOTAL_COSTO": abs(cantidad) * costo 
                        })
                except:
                    continue

    return pd.DataFrame(datos)

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Limpiador BPro", layout="wide")
st.title("🛠️ Limpiador de Reportes BPro - Compras y Traspasos")

tab1, tab2 = st.tabs(["📦 Módulo de COMPRAS", "🚚 Módulo de TRASPASOS"])

# --- PESTAÑA 1: COMPRAS ---
with tab1:
    st.markdown("### Cargar Reportes de Compras (Stock)")
    col1, col2 = st.columns(2)
    
    with col1:
        file_compras_cuauti = st.file_uploader("Subir Compras CUAUTITLÁN", type=["xls", "xlsx"], key="cc")
    with col2:
        file_compras_tulti = st.file_uploader("Subir Compras TULTITLÁN", type=["xls", "xlsx"], key="ct")
        
    if st.button("Procesar Compras", type="primary"):
        dfs_compras = []
        
        if file_compras_cuauti:
            dfs_compras.append(procesar_compras(file_compras_cuauti, "CUAUTITLAN"))
        if file_compras_tulti:
            dfs_compras.append(procesar_compras(file_compras_tulti, "TULTITLAN"))
            
        if dfs_compras:
            df_final_compras = pd.concat(dfs_compras, ignore_index=True)
            st.success(f"¡Base Generada! {len(df_final_compras)} registros encontrados.")
            st.dataframe(df_final_compras.head())
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final_compras.to_excel(writer, index=False)
            
            st.download_button("⬇️ Descargar Base Unificada COMPRAS", buffer, "Master_Compras.xlsx")
        else:
            st.warning("Sube al menos un archivo de compras.")

# --- PESTAÑA 2: TRASPASOS ---
with tab2:
    st.markdown("### Cargar Reportes de Traspasos (Salidas)")
    col3, col4 = st.columns(2)
    
    with col3:
        file_trasp_cuauti = st.file_uploader("Subir Traspasos CUAUTITLÁN", type=["xls", "xlsx"], key="tc")
    with col4:
        file_trasp_tulti = st.file_uploader("Subir Traspasos TULTITLÁN", type=["xls", "xlsx"], key="tt")
        
    if st.button("Procesar Traspasos", type="primary"):
        dfs_trasp = []
        
        if file_trasp_cuauti:
            dfs_trasp.append(procesar_traspasos(file_trasp_cuauti, "CUAUTITLAN"))
        if file_trasp_tulti:
            dfs_trasp.append(procesar_traspasos(file_trasp_tulti, "TULTITLAN"))
            
        if dfs_trasp:
            df_final_trasp = pd.concat(dfs_trasp, ignore_index=True)
            st.success(f"¡Base Generada! {len(df_final_trasp)} movimientos encontrados.")
            st.dataframe(df_final_trasp.head())
            
            buffer2 = io.BytesIO()
            with pd.ExcelWriter(buffer2, engine='xlsxwriter') as writer:
                df_final_trasp.to_excel(writer, index=False)
            
            st.download_button("⬇️ Descargar Base Unificada TRASPASOS", buffer2, "Master_Traspasos.xlsx")
        else:
            st.warning("Sube al menos un archivo de traspasos.")
