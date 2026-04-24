import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Marketplace Error Analyst", layout="wide")

st.title("🛠️ Analizador Universal de Errores")
st.markdown("Sube tu CSV. El sistema detectará automáticamente las columnas de SKU y Errores sin importar si están en mayúsculas o minúsculas.")

uploaded_file = st.file_uploader("Sube tu archivo CSV", type="csv")

if uploaded_file is not None:
    # 1. Detección inteligente de separador
    # Leemos las primeras líneas para ver si abunda el ';' o la ','
    preview = uploaded_file.getvalue().decode('utf-8', errors='ignore')[:1000]
    separator = ';' if preview.count(';') > preview.count(',') else ','
    
    uploaded_file.seek(0) # Resetear puntero del archivo
    df = pd.read_csv(uploaded_file, sep=separator)
    
    # --- NORMALIZACIÓN DE COLUMNAS ---
    # Creamos un diccionario para mapear nombres reales a nombres estándar
    col_map = {c.lower(): c for c in df.columns}
    
    # Buscamos la columna de errores
    error_col_real = None
    if 'errors' in col_map:
        error_col_real = col_map['errors']
    
    # Buscamos la columna de SKU (Prioridades)
    sku_col_real = None
    for candidate in ['shop_sku', 'shop-sku', 'sku', 'offer_sku', 'referencia']:
        if candidate in col_map:
            sku_col_real = col_map[candidate]
            break
            
    # Si aún no la encontramos, buscamos cualquier cosa que contenga 'sku'
    if not sku_col_real:
        for c_lower, c_real in col_map.items():
            if 'sku' in c_lower:
                sku_col_real = c_real
                break

    if sku_col_real and error_col_real:
        st.success(f"Columnas detectadas: SKU → `{sku_col_real}` | Errores → `{error_col_real}`")
        
        # --- PROCESAMIENTO ---
        # Aseguramos que tratamos con strings y separamos por '|'
        df['temp_errors'] = df[error_col_real].fillna('Sin Error').astype(str).str.split('|')
        df_exploded = df.explode('temp_errors')
        df_exploded['temp_errors'] = df_exploded['temp_errors'].str.strip()
        
        # --- FILTROS ---
        st.sidebar.header("🔍 Filtros")
        sku_filter = st.sidebar.text_input("Buscar por SKU específico")
        
        # Filtro de texto para errores
        error_text_search = st.sidebar.text_input("Filtrar errores por palabra (ej: 'EAN')")
        
        all_unique_errors = sorted(df_exploded['temp_errors'].unique().tolist())
        
        # Pre-filtrar lista de errores si hay búsqueda por texto
        if error_text_search:
            all_unique_errors = [e for e in all_unique_errors if error_text_search.lower() in e.lower()]
            
        selected_errors = st.sidebar.multiselect("Seleccionar errores:", all_unique_errors)

        # Aplicar lógica de filtrado al DF explotado
        df_final = df_exploded.copy()
        if sku_filter:
            df_final = df_final[df_final[sku_col_real].astype(str).str.contains(sku_filter, case=False)]
        if selected_errors:
            df_final = df_final[df_final['temp_errors'].isin(selected_errors)]
        elif error_text_search:
            df_final = df_final[df_final['temp_errors'].str.contains(error_text_search, case=False)]

        # --- TABLAS ---
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📋 Resumen Agrupado")
            # Agrupación con conteo de SKUs únicos
            resumen = df_final.groupby('temp_errors')[sku_col_real].nunique().reset_index()
            resumen.columns = ['Tipo de Error', 'Cantidad SKUs']
            st.dataframe(resumen, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("📌 Detalle SKU - Error")
            st.dataframe(df_final[[sku_col_real, 'temp_errors']], use_container_width=True, hide_index=True)

        # --- EXCEL ---
        st.divider()
        if st.button("📦 Generar Excel por Pestañas"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Pestaña Resumen
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN_GLOBAL')
                
                # Iterar por los errores únicos en el set actual
                for i, error_val in enumerate(df_final['temp_errors'].unique()):
                    subset = df_final[df_final['temp_errors'] == error_val][[sku_col_real, 'temp_errors']]
                    
                    # Limpieza extrema de nombre de pestaña para evitar errores de Excel
                    clean_name = re.sub(r'[\[\]\:\*\?\/\\]', '', str(error_val))
                    clean_name = clean_name.strip("'")[:30]
                    if not clean_name: clean_name = f"Error_{i}"
                    
                    try:
                        subset.to_excel(writer, index=False, sheet_name=clean_name)
                    except:
                        subset.to_excel(writer, index=False, sheet_name=f"Error_Hoja_{i}")
            
            st.download_button(
                label="⬇️ Descargar Reporte en Excel",
                data=output.getvalue(),
                file_name="analisis_marketplaces.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    else:
        st.error(f"No se detectaron columnas válidas. Busqué 'errors' y algo similar a 'sku'. Columnas en tu archivo: {list(df.columns)}")