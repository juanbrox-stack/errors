import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Marketplace Error Analyst Pro", layout="wide")

st.title("🚀 Analizador Universal de Marketplaces")
st.markdown("Compatible con Carrefour, Privalia, Mediamarkt, Leroy Merlin y otros basados en Mirakl.")

# 1. Carga de archivo
uploaded_file = st.file_uploader("Sube tu CSV de errores", type="csv")

if uploaded_file is not None:
    # --- DETECCIÓN DE NOMBRE DEL MARKETPLACE ---
    # Intentamos adivinarlo por el nombre del archivo
    file_name_raw = uploaded_file.name.lower()
    default_market = "Marketplace"
    if "carrefour" in file_name_raw: default_market = "Carrefour"
    elif "privalia" in file_name_raw or "veepee" in file_name_raw: default_market = "Privalia"
    elif "mediamarkt" in file_name_raw: default_market = "MediaMarkt"
    elif "leroy" in file_name_raw: default_market = "LeroyMerlin"

    # Permitir al usuario confirmar o cambiar el nombre en el sidebar
    st.sidebar.header("⚙️ Configuración")
    mkt_name = st.sidebar.text_input("Nombre del Marketplace", value=default_market)
    
    # 2. Lectura del archivo
    preview = uploaded_file.getvalue().decode('utf-8', errors='ignore')[:1000]
    separator = ';' if preview.count(';') > preview.count(',') else ','
    uploaded_file.seek(0)
    
    # Leemos todo como string para proteger GTINs y SKUs
    df = pd.read_csv(uploaded_file, sep=separator, dtype=str)
    
    # --- MAPEADO DE COLUMNAS (BÚSQUEDA EXHAUSTIVA) ---
    cols = {c.lower(): c for c in df.columns}
    
    # 1. Buscar Error (incluimos ahora los de Mirakl/Carrefour)
    error_candidates = [
        'mirakl-integration-errors', 'error_description', 'errors', 
        'mirakl-rejection-message', 'error', 'status_details'
    ]
    error_col = next((cols[c] for c in error_candidates if c in cols), None)
            
    # 2. Buscar Identificador (GTIN o SKU)
    id_candidates = [
        'mirakl-product-sku', 'gtin', 'ean', 'shop_sku', 'sku', 'product_id'
    ]
    id_col = next((cols[c] for c in id_candidates if c in cols), None)

    if id_col and error_col:
        st.success(f"✅ Detectado en {mkt_name}: ID → `{id_col}` | Errores → `{error_col}`")
        
        # --- LIMPIEZA DE IDENTIFICADORES (Evitar formato científico) ---
        # Si es un EAN o GTIN, limpiamos el formato
        df[id_col] = df[id_col].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # --- PROCESAMIENTO DE ERRORES ---
        # Algunos errores de Carrefour vienen en JSON o strings largos, separamos por '|' o salto de línea
        df['temp_errors'] = df[error_col].fillna('Sin Error').astype(str).str.split(r'\||\n')
        df_exploded = df.explode('temp_errors')
        df_exploded['temp_errors'] = df_exploded['temp_errors'].str.strip()
        # Eliminar filas donde el error quedó vacío tras el split
        df_exploded = df_exploded[df_exploded['temp_errors'] != '']
        
        # --- FILTROS ---
        st.sidebar.subheader("🔍 Filtros de Datos")
        search_id = st.sidebar.text_input(f"Buscar {id_col} específico")
        
        unique_errors = sorted([e for e in df_exploded['temp_errors'].unique() if e and e != 'nan'])
        error_search = st.sidebar.text_input("Filtrar lista de errores por palabra")
        
        if error_search:
            unique_errors = [e for e in unique_errors if error_search.lower() in e.lower()]
            
        selected_errors = st.sidebar.multiselect("Errores a incluir en el reporte:", unique_errors)

        # Aplicar filtros
        df_final = df_exploded.copy()
        if search_id:
            df_final = df_final[df_final[id_col].str.contains(search_id, case=False)]
        if selected_errors:
            df_final = df_final[df_final['temp_errors'].isin(selected_errors)]

        # --- TABLAS DE RESULTADOS ---
        c1, c2 = st.columns([2, 3])
        
        with c1:
            st.subheader("📊 Tabla Dinámica (Resumen)")
            resumen = df_final.groupby('temp_errors')[id_col].nunique().reset_index()
            resumen.columns = ['Mensaje de Error', 'Productos Afectados']
            resumen = resumen.sort_values(by='Productos Afectados', ascending=False)
            st.dataframe(resumen, use_container_width=True, hide_index=True)
            
        with c2:
            st.subheader("📌 Detalle por Producto")
            st.dataframe(df_final[[id_col, 'temp_errors']], use_container_width=True, hide_index=True)

        # --- EXPORTACIÓN ---
        st.divider()
        if st.button(f"🎁 Generar Reporte de {mkt_name} en Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Pestaña resumen
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN_ERRORES')
                
                # Una pestaña por cada error
                errors_to_sheet = df_final['temp_errors'].unique()
                for i, err_name in enumerate(errors_to_sheet):
                    if not err_name or str(err_name) == 'nan': continue
                    
                    subset = df_final[df_final['temp_errors'] == err_name][[id_col, 'temp_errors']]
                    
                    # Limpieza nombre de pestaña (reglas Excel)
                    clean_sheet = re.sub(r'[\[\]\:\*\?\/\\]', '', str(err_name))
                    clean_sheet = clean_sheet.strip("'")[:30]
                    if not clean_sheet: clean_sheet = f"Error_{i}"
                    
                    try:
                        subset.to_excel(writer, index=False, sheet_name=clean_sheet)
                    except:
                        subset.to_excel(writer, index=False, sheet_name=f"Detalle_{i}")
            
            # Nombre de archivo dinámico
            final_filename = f"Analisis_{mkt_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
            
            st.download_button(
                label="⬇️ Descargar Archivo Excel",
                data=output.getvalue(),
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error(f"❌ No se detectaron las columnas necesarias.")
        st.write("Columnas encontradas en el archivo:", list(df.columns))
        st.info("Buscaba algo parecido a: 'mirakl-integration-errors' o 'error_description' para los fallos.")