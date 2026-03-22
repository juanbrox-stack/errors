import streamlit as st
import pandas as pd
import io

# Configuración visual
st.set_page_config(page_title="Marketplace Error Master", layout="wide")

st.title("🚀 Optimizador de Errores Marketplace")
st.markdown("""
Esta herramienta agrupa tus errores y genera un Excel con una pestaña 
dedicada a cada tipo de error para facilitar la corrección masiva.
""")

# 1. Carga de archivo con detección de separador
uploaded_file = st.file_uploader("Sube tu CSV de errores", type="csv")

if uploaded_file is not None:
    # Leemos el archivo usando el separador detectado en tus ejemplos
    df = pd.read_csv(uploaded_file, sep=';')
    
    # Limpieza básica: asegurar que las columnas existen
    if 'errors' in df.columns and 'shop_sku' in df.columns:
        
        # Filtros en la barra lateral
        st.sidebar.header("🔍 Panel de Filtros")
        
        # Buscador por SKU
        sku_to_find = st.sidebar.text_input("Buscar un SKU concreto")
        
        # Selector de errores múltiples
        all_errors = df['errors'].dropna().unique().tolist()
        selected_errors = st.sidebar.multiselect(
            "Selecciona errores para el informe:", 
            options=all_errors,
            default=all_errors[:3] if len(all_errors) > 3 else all_errors
        )

        # Aplicar lógica de filtrado
        df_filtered = df.copy()
        if sku_to_find:
            df_filtered = df_filtered[df_filtered['shop_sku'].astype(str).str.contains(sku_to_find)]
        if selected_errors:
            df_filtered = df_filtered[df_filtered['errors'].isin(selected_errors)]

        # --- SECCIÓN DE ANÁLISIS ---
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📋 Resumen de Incidencias")
            # Agrupación por error y conteo
            summary = df_filtered.groupby('errors')['shop_sku'].count().reset_index()
            summary.columns = ['Tipo de Error', 'Total SKUs']
            st.dataframe(summary, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("📌 Vista Detallada")
            st.dataframe(df_filtered[['shop_sku', 'errors']], use_container_width=True, hide_index=True)

        # --- GENERACIÓN DE EXCEL POR PESTAÑAS ---
        st.divider()
        st.subheader("📦 Generar Reporte de Salida")
        
        if st.button("Preparar Excel para Descarga"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Pestaña 1: Resumen General
                summary.to_excel(writer, index=False, sheet_name='RESUMEN_GENERAL')
                
                # Pestañas Dinámicas: Un error por pestaña
                for error_name in selected_errors:
                    # Filtrar datos de ese error
                    specific_df = df_filtered[df_filtered['errors'] == error_name][['shop_sku', 'errors']]
                    
                    # Limpiar nombre de pestaña (Excel no permite > 31 caracteres o ciertos símbolos)
                    clean_name = str(error_name)[:30].replace(':', '').replace('/', '').replace('|', '-')
                    specific_df.to_excel(writer, index=False, sheet_name=clean_name)
            
            processed_data = output.getvalue()
            
            st.download_button(
                label="⬇️ Descargar Excel (Pestañas por Error)",
                data=processed_data,
                file_name="correccion_marketplaces.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("¡Excel generado con éxito! Cada error tiene su propia pestaña.")
            
    else:
        st.error("Error: No se han encontrado las columnas 'errors' o 'shop_sku'. Revisa el formato del CSV.")