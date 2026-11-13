import streamlit as st
import pandas as pd
import numpy as np
import io

# Librerías para conexión con Google Drive API
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# =====================================================================
# CONFIGURACIÓN INTERNA DE GOOGLE DRIVE
# =====================================================================
# Pega aquí tu enlace completo de la carpeta de Drive para que quede fijo en el código:
DRIVE_FOLDER_URL = st.secrets.get(
    "DRIVE_FOLDER_URL",
    "https://drive.google.com/drive/folders/1OZ2vUuPgmwPmVeNdXd7FR7MgiCfedr2a"
)


# =====================================================================
# FUNCIONES UTILITARIAS DE LIMPIEZA Y FORMATO
# =====================================================================
def limpiar_np(serie):
    """Convierte números de parte a texto en MAYÚSCULAS y elimina cualquier espacio."""
    return (
        serie.astype(str)
        .str.upper()
        .str.replace(r"\s+", "", regex=True)
        .replace({"NAN": np.nan, "NONE": np.nan, "": np.nan})
    )


def limpiar_texto(serie):
    """Estandariza nombres de vendedores, sucursales y almacenes."""
    return (
        serie.astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )


def parsear_fecha_latam(serie):
    """Convierte fechas en formato LATAM (DD/MM/YYYY) a datetime."""
    return pd.to_datetime(serie, dayfirst=True, errors="coerce")


# =====================================================================
# FUNCIÓN 1: LIMPIEZA DE COMPRAS (BPRO)
# =====================================================================
def procesar_compras(file, nombre_agencia):
    """
    Limpia el archivo de compras evitando filas basura o subtotales.
    Procesa todas las hojas manteniendo la continuidad de los encabezados.
    """
    dict_dfs = pd.read_excel(file, sheet_name=None, header=None)
    datos = []
    factura, fecha, proveedor, comprador = None, None, None, None

    for hoja, df in dict_dfs.items():
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
                except Exception:
                    continue

            # Detectar Ítems (CRCU / CRTU / o cualquier código de compra)
            elif val_a.startswith("CR"):
                descripcion = str(row[3]).strip()
                if descripcion and descripcion.lower() != "nan":
                    datos.append({
                        "AGENCIA": nombre_agencia,
                        "FACTURA": factura,
                        "FECHA": fecha,
                        "PROVEEDOR": proveedor,
                        "COMPRADOR": comprador,
                        "NP": row[2],
                        "DESCRIPCION": row[3],
                        "CANTIDAD": row[4],
                        "COSTO UNITARIO": row[5],
                        "SUBTOTAL": row[7],
                        "IVA": row[8],
                        "TOTAL": row[9],
                    })

    return pd.DataFrame(datos)


# =====================================================================
# FUNCIÓN 2: LIMPIEZA DE TRASPASOS - SALIDAS (BPRO)
# =====================================================================
def procesar_traspasos(file, nombre_agencia):
    """
    Limpia traspasos de salida abarcando todas las variantes y evita subtotales finales.
    Procesa todas las hojas manteniendo la continuidad de los encabezados.
    """
    dict_dfs = pd.read_excel(file, sheet_name=None, header=None)
    datos = []
    destino_actual = None
    referencia, fecha_mov, usuario = None, None, None

    for hoja, df in dict_dfs.items():
        for _, row in df.iterrows():
            val_a = str(row[0]).strip().upper()

            # NIVEL 1: DETECTAR DESTINO (SALIDA...)
            if val_a.startswith("SALIDA"):
                if "HACIA" in val_a:
                    try:
                        destino_sucio = val_a.split("HACIA")[1]
                        destino_actual = destino_sucio.strip()
                    except Exception:
                        continue
                elif "SALIDA DE ALMACEN POR TRASPASO" in val_a:
                    destino_actual = f"SALIDA DE ALMACEN POR TRASPASO {nombre_agencia}"

            # NIVEL 2: DETECTAR CABECERA (REFERENCIA / FECHA / USUARIO)
            elif "REFERENCIA:" in val_a:
                try:
                    referencia = val_a.split("REFERENCIA:")[1].strip()
                    if "FECHA MOV:" in str(row[2]).upper():
                        fecha_mov = str(row[2]).upper().split("FECHA MOV:")[1].strip()
                    if "USUARIO:" in str(row[3]).upper():
                        usuario = str(row[3]).upper().split("USUARIO:")[1].strip()
                except Exception:
                    continue

            # NIVEL 3: DETECTAR ÍTEMS (TRAS...)
            elif val_a.startswith("TRAS"):
                if destino_actual:
                    try:
                        descripcion = str(row[3]).strip()
                        if descripcion and descripcion.lower() != "nan":
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
                                "TOTAL_COSTO": abs(cantidad) * costo,
                            })
                    except Exception:
                        continue

    return pd.DataFrame(datos)


# =====================================================================
# FUNCIÓN 3: LIMPIEZA DE TRASPASOS - ENTRADAS (BPRO)
# =====================================================================
def procesar_entradas_traspasos(file, nombre_agencia="CUAUTITLAN"):
    """
    Limpia el reporte de ENTRADAS por traspaso de BPro.
    Cambia el rubro DESTINO por ORIGEN extrayendo el almacén después de 'DESDE'
    (Ej: 'Entrada a Alm. Gral. Cuautitlan desde Alm. Gral. Tultitlan' -> 'ALM. GRAL. TULTITLAN').
    """
    df_preview = pd.read_excel(file, nrows=5)
    cols_upper = [str(c).strip().upper() for c in df_preview.columns]
    if "FECHA_MOV" in cols_upper and "NP" in cols_upper:
        file.seek(0)
        df_directo = pd.read_excel(file)
        df_directo.columns = [str(c).strip().upper() for c in df_directo.columns]
        if "DESTINO" in df_directo.columns and "ORIGEN" not in df_directo.columns:
            df_directo.rename(columns={"DESTINO": "ORIGEN"}, inplace=True)
        return df_directo

    file.seek(0)
    dict_dfs = pd.read_excel(file, sheet_name=None, header=None)
    datos = []

    origen_actual = None
    referencia, fecha_mov, usuario = None, None, None

    for hoja, df in dict_dfs.items():
        for _, row in df.iterrows():
            val_a = str(row[0]).strip().upper()

            # NIVEL 1: DETECTAR ORIGEN (ENTRADA...)
            if val_a.startswith("ENTRADA"):
                if "DESDE" in val_a:
                    try:
                        origen_sucio = val_a.split("DESDE")[1]
                        origen_actual = origen_sucio.strip()
                    except Exception:
                        continue
                elif "TRASPASO" in val_a:
                    origen_actual = f"ENTRADA DE ALMACEN POR TRASPASO {nombre_agencia}"

            # NIVEL 2: DETECTAR CABECERA (REFERENCIA / FECHA / USUARIO)
            elif "REFERENCIA:" in val_a:
                try:
                    referencia = val_a.split("REFERENCIA:")[1].strip()
                    if "FECHA MOV:" in str(row[2]).upper():
                        fecha_mov = str(row[2]).upper().split("FECHA MOV:")[1].strip()
                    if "USUARIO:" in str(row[3]).upper():
                        usuario = str(row[3]).upper().split("USUARIO:")[1].strip()
                except Exception:
                    continue

            # NIVEL 3: DETECTAR ÍTEMS (TRAS / ENTR / ETRA...)
            elif val_a.startswith(("TRAS", "ENTR", "ETRA")):
                if origen_actual:
                    try:
                        descripcion = str(row[3]).strip()
                        if descripcion and descripcion.lower() != "nan":
                            cantidad = float(row[4])
                            costo = float(row[5])
                            datos.append({
                                "AGENCIA": nombre_agencia,
                                "ORIGEN": origen_actual,
                                "REFERENCIA": referencia,
                                "FECHA_MOV": fecha_mov,
                                "USUARIO": usuario,
                                "NP": row[2],
                                "DESCRIPCION": row[3],
                                "CANTIDAD": abs(cantidad),
                                "COSTO_UNIT": costo,
                                "TOTAL_COSTO": abs(cantidad) * costo,
                            })
                    except Exception:
                        continue

    return pd.DataFrame(datos)


# =====================================================================
# FUNCIÓN 4: CONEXIÓN INTERNA A GOOGLE DRIVE PARA BASES DE VENTAS MASTER
# =====================================================================
def extraer_id_carpeta(url_o_id):
    """Extrae el ID limpio desde el link fijo de Google Drive."""
    texto = str(url_o_id).strip()
    if "/folders/" in texto:
        texto = texto.split("/folders/")[1].split("?")[0].split("/")[0]
    return texto.strip()


@st.cache_data(ttl=600, show_spinner=False)
def cargar_ventas_desde_drive():
    """
    Se conecta automáticamente a la carpeta fija en Google Drive usando st.secrets,
    descarga los archivos 'AGENCIA_Ventas_AÑO_MASTER' y los une en un DataFrame.
    """
    folder_id = extraer_id_carpeta(DRIVE_FOLDER_URL)
    if not folder_id or "PEGA_AQUI" in folder_id:
        raise ValueError(
            "Configura el enlace de tu carpeta en la constante DRIVE_FOLDER_URL del código."
        )

    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    service = build("drive", "v3", credentials=creds)

    query = (
        f"'{folder_id}' in parents and trashed = false "
        f"and name contains '_Ventas_' and name contains '_MASTER'"
    )
    results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
    archivos = results.get("files", [])

    if not archivos:
        raise FileNotFoundError(
            "No se encontraron archivos con el patrón '_Ventas_..._MASTER' en la carpeta de Drive."
        )

    lista_dfs = []
    for archivo in archivos:
        nombre = archivo["name"].upper()
        if "CUAUTITLAN" not in nombre and any("CUAUTITLAN" in a["name"].upper() for a in archivos):
            continue

        request = service.files().get_media(fileId=archivo["id"])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        df_temp = pd.read_excel(fh)
        lista_dfs.append(df_temp)

    if not lista_dfs:
        return pd.DataFrame()

    return pd.concat(lista_dfs, ignore_index=True)


# =====================================================================
# FUNCIÓN 5: MOTOR DE CRUCES "MODELO LUIS"
# =====================================================================
def construir_modelo_luis(df_solicitudes_raw, df_inventario_raw, df_entradas_raw, df_ventas_raw):
    # 1. PREPARAR BASE 1: SOLICITUDES DE TRASPASO
    df_sol = df_solicitudes_raw.copy()
    df_sol.columns = [str(c).strip().upper() for c in df_sol.columns]

    col_fecha_sol = "C<" if "C<" in df_sol.columns else df_sol.columns[0]
    col_np_sol = "N° DE PARTE" if "N° DE PARTE" in df_sol.columns else "NP"

    df_base = pd.DataFrame()
    df_base["AGENCIA"] = ["CRA CUAUTITLAN"] * len(df_sol)
    df_base["VENDEDOR"] = limpiar_texto(df_sol["SOLICITA"])
    df_base["FECHA SOLICITUD"] = parsear_fecha_latam(df_sol[col_fecha_sol])
    df_base["N° DE PARTE"] = limpiar_np(df_sol[col_np_sol])
    df_base["CANTIDAD SOLICITADA"] = pd.to_numeric(df_sol["CANTIDAD"], errors="coerce").fillna(0)

    df_base = df_base.dropna(subset=["N° DE PARTE"]).reset_index(drop=True)

    # 2. PREPARAR Y CRUZAR BASE 2: BASE INVENTARIO
    df_inv = df_inventario_raw.copy()
    df_inv.columns = [str(c).strip().upper() for c in df_inv.columns]

    df_inv["SUCURSAL_CLEAN"] = limpiar_texto(df_inv["SUCURSAL"])
    df_inv["ALMACEN_CLEAN"] = limpiar_texto(df_inv["ALMACEN"])
    df_inv["NP_CLEAN"] = limpiar_np(df_inv["NP"])

    df_inv_filtrado = df_inv[
        (df_inv["SUCURSAL_CLEAN"] == "CUAUTITLAN")
        & (df_inv["ALMACEN_CLEAN"] == "ALM. GENERAL")
    ].copy()

    df_inv_filtrado["EXISTENCIA"] = pd.to_numeric(
        df_inv_filtrado["EXISTENCIA"], errors="coerce"
    ).fillna(0)

    df_inv_unico = (
        df_inv_filtrado.groupby("NP_CLEAN", as_index=False)
        .agg({"DESCRIPCION": "first", "EXISTENCIA": "sum"})
    )

    df_base = df_base.merge(
        df_inv_unico,
        left_on="N° DE PARTE",
        right_on="NP_CLEAN",
        how="left",
    )
    df_base["DESCRIPCION"] = df_base["DESCRIPCION"].fillna("SIN DESCRIPCIÓN EN INVENTARIO")
    df_base["EXISTENCIA"] = df_base["EXISTENCIA"].fillna(0)
    df_base.drop(columns=["NP_CLEAN"], inplace=True, errors="ignore")

    # 3. PREPARAR Y CRUZAR BASE 3: ENTRADAS (INGRESO Y CANTIDAD TRASPASADA)
    df_ent = df_entradas_raw.copy()
    df_ent.columns = [str(c).strip().upper() for c in df_ent.columns]

    if "DESTINO" in df_ent.columns and "ORIGEN" not in df_ent.columns:
        df_ent.rename(columns={"DESTINO": "ORIGEN"}, inplace=True)

    df_ent["AGENCIA_CLEAN"] = limpiar_texto(df_ent["AGENCIA"])
    df_ent["ORIGEN_CLEAN"] = limpiar_texto(df_ent["ORIGEN"])
    df_ent["NP_CLEAN"] = limpiar_np(df_ent["NP"])
    df_ent["FECHA_MOV_DT"] = parsear_fecha_latam(df_ent["FECHA_MOV"])
    df_ent["CANTIDAD_NUM"] = pd.to_numeric(df_ent["CANTIDAD"], errors="coerce").fillna(0)

    df_ent_filtrado = df_ent[
        (df_ent["AGENCIA_CLEAN"] == "CUAUTITLAN")
        & (df_ent["ORIGEN_CLEAN"].str.contains("ALM. GRAL. TULTITLAN", na=False))
    ].dropna(subset=["NP_CLEAN", "FECHA_MOV_DT"]).sort_values("FECHA_MOV_DT", ascending=True)

    ingresos_encontrados = []
    cant_traspasadas = []

    for _, row in df_base.iterrows():
        np_actual = row["N° DE PARTE"]
        fecha_sol = row["FECHA SOLICITUD"]

        if pd.isna(fecha_sol):
            ingresos_encontrados.append(pd.NaT)
            cant_traspasadas.append(0.0)
            continue

        candidatos = df_ent_filtrado[
            (df_ent_filtrado["NP_CLEAN"] == np_actual)
            & (df_ent_filtrado["FECHA_MOV_DT"] >= fecha_sol)
        ]

        if not candidatos.empty:
            primer_ingreso = candidatos.iloc[0]
            ingresos_encontrados.append(primer_ingreso["FECHA_MOV_DT"])
            cant_traspasadas.append(float(primer_ingreso["CANTIDAD_NUM"]))
        else:
            ingresos_encontrados.append(pd.NaT)
            cant_traspasadas.append(0.0)

    df_base["INGRESO"] = ingresos_encontrados
    df_base["CANTIDAD TRASPASADA"] = cant_traspasadas

    # 4. PREPARAR Y CRUZAR BASE 4: VENTAS MASTER
    df_vta = df_ventas_raw.copy()
    df_vta.columns = [str(c).strip().upper() for c in df_vta.columns]

    df_vta["SUCURSAL_CLEAN"] = limpiar_texto(df_vta["SUCURSAL"])
    df_vta["NP_CLEAN"] = limpiar_np(df_vta["NP"])
    df_vta["VENDEDOR_CLEAN"] = limpiar_texto(df_vta["VENDEDOR"])
    df_vta["FECHA_DT"] = parsear_fecha_latam(df_vta["FECHA"])
    df_vta["CANTIDAD_NUM"] = pd.to_numeric(df_vta["CANTIDAD"], errors="coerce").fillna(0)

    df_vta_cuauti = df_vta[
        df_vta["SUCURSAL_CLEAN"] == "CUAUTITLAN"
    ].dropna(subset=["NP_CLEAN", "FECHA_DT"]).sort_values("FECHA_DT", ascending=True)

    fechas_venta = []
    cantidades_vendidas = []

    for _, row in df_base.iterrows():
        np_actual = row["N° DE PARTE"]
        fecha_ing = row["INGRESO"]
        vendedor_solicita = row["VENDEDOR"]

        if pd.isna(fecha_ing):
            fechas_venta.append(pd.NaT)
            cantidades_vendidas.append(0.0)
            continue

        ventas_np = df_vta_cuauti[
            (df_vta_cuauti["NP_CLEAN"] == np_actual)
            & (df_vta_cuauti["FECHA_DT"] >= fecha_ing)
        ]

        if ventas_np.empty:
            fechas_venta.append(pd.NaT)
            cantidades_vendidas.append(0.0)
        else:
            ventas_vendedor = ventas_np[ventas_np["VENDEDOR_CLEAN"] == vendedor_solicita]
            if not ventas_vendedor.empty:
                primera_venta_vend = ventas_vendedor.iloc[0]
                fechas_venta.append(primera_venta_vend["FECHA_DT"])
                cantidades_vendidas.append(float(primera_venta_vend["CANTIDAD_NUM"]))
            else:
                fechas_venta.append(ventas_np.iloc[0]["FECHA_DT"])
                cantidades_vendidas.append(0.0)

    df_base["FECHA VENTA"] = fechas_venta
    df_base["CANTIDAD VENDIDA"] = cantidades_vendidas

    # 5. CÁLCULO DE MÉTRICAS FINALES
    df_base["DIAS FACTURACIÓN"] = (df_base["FECHA VENTA"] - df_base["INGRESO"]).dt.days

    def evaluar_estatus(row):
        fecha_sol = row["FECHA SOLICITUD"]
        fecha_ing = row["INGRESO"]
        cant_sol = row["CANTIDAD SOLICITADA"]
        cant_ven = row["CANTIDAD VENDIDA"]

        if pd.isna(fecha_ing):
            return "SIN INGRESO"

        dias_ingreso = (fecha_ing - fecha_sol).days if pd.notna(fecha_sol) else 0

        if dias_ingreso > 3:
            return "INGRESO TARDÍO"
        if cant_ven == 0:
            return "VENTA FRACASO"
        if cant_ven >= cant_sol and cant_sol > 0:
            return "VENTA EXITOSA"
        if 0 < cant_ven < cant_sol:
            return "VENTA PARCIAL"

        return "SIN CLASIFICAR"

    df_base["ESTATUS"] = df_base.apply(evaluar_estatus, axis=1)

    df_base["TASA ÉXITO"] = np.where(
        df_base["CANTIDAD SOLICITADA"] > 0,
        (df_base["CANTIDAD VENDIDA"] / df_base["CANTIDAD SOLICITADA"]).round(4),
        0.0,
    )

    for col_f in ["FECHA SOLICITUD", "INGRESO", "FECHA VENTA"]:
        df_base[col_f] = df_base[col_f].dt.strftime("%d/%m/%Y").fillna("")

    columnas_finales = [
        "AGENCIA",
        "VENDEDOR",
        "FECHA SOLICITUD",
        "N° DE PARTE",
        "DESCRIPCION",
        "CANTIDAD SOLICITADA",
        "INGRESO",
        "CANTIDAD TRASPASADA",
        "FECHA VENTA",
        "CANTIDAD VENDIDA",
        "EXISTENCIA",
        "DIAS FACTURACIÓN",
        "ESTATUS",
        "TASA ÉXITO",
    ]

    return df_base[columnas_finales]


# =====================================================================
# INTERFAZ STREAMLIT
# =====================================================================
st.set_page_config(page_title="Limpiador BPro & Modelo Luis", layout="wide")
st.title("🛠️ Limpiador de Reportes BPro y Modelo de Traspasos")

tab1, tab2, tab3 = st.tabs([
    "📦 Módulo de COMPRAS",
    "🚚 Módulo de TRASPASOS",
    "📊 MODELO LUIS",
])

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
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_final_compras.to_excel(writer, index=False)

            st.download_button(
                "⬇️ Descargar Base Unificada COMPRAS",
                buffer.getvalue(),
                "Master_Compras.xlsx",
            )
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
            with pd.ExcelWriter(buffer2, engine="xlsxwriter") as writer:
                df_final_trasp.to_excel(writer, index=False)

            st.download_button(
                "⬇️ Descargar Base Unificada TRASPASOS",
                buffer2.getvalue(),
                "Master_Traspasos.xlsx",
            )
        else:
            st.warning("Sube al menos un archivo de traspasos.")

# --- PESTAÑA 3: MODELO LUIS ---
with tab3:
    st.markdown("### 📊 Generador de Base Madre - MODELO LUIS")
    st.caption(
        "Sube las 3 bases locales; el sistema se conectará automáticamente a Google Drive "
        "para cruzar contra las bases Master de Ventas."
    )

    col_m1, col_m2, col_m3 = st.columns(3)

    with col_m1:
        file_solicitudes = st.file_uploader(
            "1️⃣ SOLICITUDES DE TRASPASO",
            type=["xls", "xlsx"],
            key="ml_sol",
        )
    with col_m2:
        file_inventario = st.file_uploader(
            "2️⃣ BASE INVENTARIO",
            type=["xls", "xlsx"],
            key="ml_inv",
        )
    with col_m3:
        file_entradas = st.file_uploader(
            "3️⃣ ENTRADAS TRASPASO (Cuautitlán)",
            type=["xls", "xlsx"],
            key="ml_ent",
        )

    if st.button("🚀 Generar Base Madre (MODELO LUIS)", type="primary"):
        if not (file_solicitudes and file_inventario and file_entradas):
            st.warning("⚠️ Por favor sube las 3 bases (Solicitudes, Inventario y Entradas).")
        else:
            with st.spinner("Procesando archivos y consultando Ventas Master en Google Drive..."):
                try:
                    # 1. Leer Solicitudes e Inventario
                    df_sol_raw = pd.read_excel(file_solicitudes)
                    df_inv_raw = pd.read_excel(file_inventario)

                    # 2. Procesar y limpiar Entradas de Cuautitlán
                    df_ent_raw = procesar_entradas_traspasos(file_entradas, "CUAUTITLAN")

                    # 3. Cargar Ventas automáticamente desde la carpeta interna de Google Drive
                    df_vta_raw = cargar_ventas_desde_drive()

                    # 4. Construir Base Madre
                    df_modelo_luis = construir_modelo_luis(
                        df_sol_raw,
                        df_inv_raw,
                        df_ent_raw,
                        df_vta_raw,
                    )

                    st.success(
                        f"✅ ¡Base Madre generada con éxito! Total de solicitudes procesadas: {len(df_modelo_luis)}"
                    )
                    st.dataframe(df_modelo_luis, use_container_width=True)

                    # 5. Exportar a Excel con formato
                    buffer_luis = io.BytesIO()
                    with pd.ExcelWriter(buffer_luis, engine="xlsxwriter") as writer:
                        df_modelo_luis.to_excel(writer, sheet_name="MODELO_LUIS", index=False)
                        workbook = writer.book
                        worksheet = writer.sheets["MODELO_LUIS"]
                        pct_fmt = workbook.add_format({"num_format": "0.0%"})
                        worksheet.set_column(13, 13, 14, pct_fmt)

                    st.download_button(
                        label="⬇️ Descargar Base Final MODELO LUIS (.xlsx)",
                        data=buffer_luis.getvalue(),
                        file_name="Base_Madre_Modelo_Luis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"❌ Ocurrió un error durante el procesamiento: {e}")
