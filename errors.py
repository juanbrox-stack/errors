import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Marketplace Error Master", layout="wide")

st.title("📊 Analizador Multierror de Marketplace")
st.markdown("Carga tu CSV. Los errores separados por `|` se analizarán de forma individual.")

uploaded_file = st.file_uploader("Sube tu CSV", type="csv")

if uploaded_file is not None:
    # 1. Lectura inicial
    df = pd.read_csv(uploaded_file, sep=';')
    
    if 'errors' in df.columns and 'shop_sku' in df.columns:
        
        # --- PROCESAMIENTO DE ERRORES MÚLTIPLES ---
        # Convertimos la columna de errores en una lista separando por '|'
        df['errors_list'] = df['errors'].fillna('Sin Error').astype(str).str.split('|')
        
        # 'Explotamos' la lista: una fila por cada error individual
        df_exploded = df.explode('errors_list')
        
        # Limpieza de espacios en blanco que puedan quedar tras el split
        df_exploded['errors_list'] = df_exploded['errors_list'].str.strip()
        
        # --- PANEL DE FILTROS ---
        st.sidebar.header("🔍 Filtros")
        sku_search = st.sidebar.text_input("Buscar SKU")
        
        all_unique_errors = sorted(df_exploded['errors_list'].unique().tolist())
        selected_errors = st.sidebar.multiselect(
            "Selecciona errores para analizar:", 
            options=all_unique_errors
        )

        # Aplicar filtros
        df_to_show = df_exploded.copy()
        if sku_search:
            df_to_show = df_to_show[df_to_show['shop_sku'].astype(str).str.contains(sku_search)]
        if selected_errors:
            df_to_show = df_to_show[df_to_show['errors_list'].isin(selected_errors)]

        # --- VISUALIZACIÓN ---
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📋 Resumen por Error Individual")
            # Agrupamos por el error individual y contamos SKUs únicos
            summary = df_to_show.groupby('errors_list')['shop_sku'].nunique().reset_index()
            summary.columns = ['Detalle del Error', 'Cantidad SKUs']
            st.dataframe(summary, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("📌 SKUs Afectados")
            # Mostramos el SKU y el error individual asociado
            st.dataframe(df_to_show[['shop_sku', 'errors_list']], use_container_width=True, hide_index=True)

        # --- EXPORTACIÓN ---
        if st.button("🚀 Generar Excel por Pestañas"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Pestaña de resumen
                summary.to_excel(writer, index=False, sheet_name='RESUMEN_DINAMICO')
                
                # Una pestaña por cada error individual seleccionado
                errors_to_export = selected_errors if selected_errors else all_unique_errors
                
                for i, err in enumerate(errors_to_export):
                    # Filtrar datos para este error específico
                    temp_df = df_to_show[df_to_show['errors_list'] == err][['shop_sku', 'errors_list']]
                    
                    # Limpieza de nombre de pestaña (regla de la comilla simple corregida)
                    clean_name = re.sub(r'[\[\]\:\*\?\/\\]', '', str(err))
                    clean_name = clean_name.strip("'")[:30]
                    
                    if not clean_name: clean_name = f"Error_{i}"
                    
                    try:
                        temp_df.to_excel(writer, index=False, sheet_name=clean_name)
                    except:
                        temp_df.to_excel(writer, index=False, sheet_name=f"E_{i}")

            st.download_button(
                label="📥 Descargar Reporte Segmentado",
                data=output.getvalue(),
                file_name="errores_segmentados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Columnas 'errors' o 'shop_sku' no encontradas.")