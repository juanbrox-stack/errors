import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Analizador de Errores Marketplace", layout="wide")

st.title("📊 Analizador de Errores de Marketplace")
st.markdown("Sube tu archivo CSV para identificar SKUs afectados y agrupar errores.")

# 1. Carga de archivo
uploaded_file = st.file_uploader("Elige un archivo CSV", type="csv")

if uploaded_file is not None:
    # Leemos el archivo. Nota: El archivo de ejemplo usa sep=';'
    df = pd.read_csv(uploaded_file, sep=';')
    
    # Verificación de columnas necesarias
    if 'errors' in df.columns and 'shop_sku' in df.columns:
        
        # --- FILTROS ---
        st.sidebar.header("Filtros")
        
        # Buscador de SKU
        sku_search = st.sidebar.text_input("Buscar SKU específico")
        
        # Filtro de Errores (eliminando nulos para la lista)
        lista_errores = df['errors'].dropna().unique().tolist()
        selected_errors = st.sidebar.multiselect("Filtrar por tipo de error", lista_errores)

        # Aplicar filtros al DataFrame original
        df_filtered = df.copy()
        if sku_search:
            df_filtered = df_filtered[df_filtered['shop_sku'].astype(str).str.contains(sku_search)]
        if selected_errors:
            df_filtered = df_filtered[df_filtered['errors'].isin(selected_errors)]

        # --- CUADRO DE MANDO (RESUMEN) ---
        st.subheader("Resumen de Errores (Tabla Dinámica)")
        
        # Agrupación por error y conteo de SKUs
        error_summary = df_filtered.groupby('errors')['shop_sku'].count().reset_index()
        error_summary.columns = ['Error', 'Cantidad de SKUs afectados']
        
        col1, col2 = st.columns([1, 2])
        with col1:
            st.dataframe(error_summary, use_container_width=True)
        
        with col2:
            st.subheader("Detalle de SKUs")
            st.dataframe(df_filtered[['shop_sku', 'errors']], use_container_width=True)

        # --- EXPORTACIÓN A EXCEL ---
        st.divider()
        st.subheader("Descargar Resultados")
        
        # Creamos el archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtered[['shop_sku', 'errors']].to_excel(writer, index=False, sheet_name='Detalle_Errores')
            error_summary.to_excel(writer, index=False, sheet_name='Resumen_Agrupado')
            
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Descargar Excel con errores y SKUs",
            data=excel_data,
            file_name="analisis_errores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        st.error("El archivo no contiene las columnas requeridas: 'errors' y 'shop_sku'.")

else:
    st.info("Por favor, sube un archivo CSV para comenzar el análisis.")