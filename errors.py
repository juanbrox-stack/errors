import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Marketplace Error Analyst Pro", layout="wide")

st.title("🛠️ Analizador Universal (Privalia, MM, Leroy)")

uploaded_file = st.file_uploader("Sube tu CSV de errores", type="csv")

if uploaded_file is not None:
    # 1. Detección de separador
    preview = uploaded_file.getvalue().decode('utf-8', errors='ignore')[:1000]
    separator = ';' if preview.count(';') > preview.count(',') else ','
    uploaded_file.seek(0)
    
    df = pd.read_csv(uploaded_file, sep=separator, dtype=str) # Leemos todo como str para evitar pérdida de datos
    
    # --- MAPEADO FLEXIBLE DE COLUMNAS ---
    cols = {c.lower(): c for c in df.columns}
    
    # 1. Buscar Error (ahora incluye 'error_description')
    error_col = None
    for cand in ['errors', 'error_description', 'error', 'status_details']:
        if cand in cols:
            error_col = cols[cand]
            break
            
    # 2. Buscar Identificador (GTIN o SKU)
    id_col = None
    # Prioridad: gtin (para Privalia), luego shop_sku, luego sku
    for cand in ['gtin', 'ean', 'shop_sku', 'sku', 'product_id']:
        if cand in cols:
            id_col = cols[cand]
            break

    if id_col and error_col:
        st.success(f"Detectado: ID → `{id_col}` | Errores → `{error_col}`")
        
        # --- LIMPIEZA DE GTIN (Evitar formato científico) ---
        if 'gtin' in id_col.lower() or 'ean' in id_col.lower():
            # Quitamos decimales y espacios, convertimos a string limpio
            df[id_col] = df[id_col].str.replace(r'\.0$', '', regex=True).fillna('')
            st.info("💡 Columna GTIN detectada: Se ha forzado el formato numérico sin decimales.")

        # --- PROCESAMIENTO ---
        # Separar por '|' o por salto de línea (algunos marketplaces usan saltos)
        df['temp_errors'] = df[error_col].fillna('Sin Error').astype(str).str.split(r'\||\n')
        df_exploded = df.explode('temp_errors')
        df_exploded['temp_errors'] = df_exploded['temp_errors'].str.strip()
        
        # --- FILTROS ---
        st.sidebar.header("🔍 Filtros")
        search_id = st.sidebar.text_input(f"Buscar {id_col}")
        
        unique_errors = sorted([e for e in df_exploded['temp_errors'].unique() if e])
        selected_errors = st.sidebar.multiselect("Filtrar por errores:", unique_errors)

        # Aplicar filtros
        df_final = df_exploded.copy()
        if search_id:
            df_final = df_final[df_final[id_col].str.contains(search_id, case=False)]
        if selected_errors:
            df_final = df_final[df_final['temp_errors'].isin(selected_errors)]

        # --- VISUALIZACIÓN ---
        c1, c2 = st.columns(2)
        
        with c1:
            st.subheader("📊 Resumen")
            resumen = df_final.groupby('temp_errors')[id_col].nunique().reset_index()
            resumen.columns = ['Error', 'Total Afectados']
            st.dataframe(resumen, use_container_width=True, hide_index=True)
            
        with c2:
            st.subheader("📌 Listado")
            st.dataframe(df_final[[id_col, 'temp_errors']], use_container_width=True, hide_index=True)

        # --- EXPORTACIÓN ---
        if st.button("📦 Descargar Excel Segmentado"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Resumen
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN')
                
                # Una pestaña por error
                for i, err_name in enumerate(df_final['temp_errors'].unique()):
                    if not err_name or err_name == 'nan': continue
                    
                    subset = df_final[df_final['temp_errors'] == err_name][[id_col, 'temp_errors']]
                    
                    # Limpieza de nombre de pestaña (Excel rules)
                    clean_sheet = re.sub(r'[\[\]\:\*\?\/\\]', '', str(err_name))
                    clean_sheet = clean_sheet.strip("'")[:30]
                    
                    try:
                        subset.to_excel(writer, index=False, sheet_name=clean_sheet if clean_sheet else f"Error_{i}")
                    except:
                        subset.to_excel(writer, index=False, sheet_name=f"Hoja_{i}")
            
            st.download_button(
                label="⬇️ Descargar Reporte",
                data=output.getvalue(),
                file_name="reporte_errores_privalia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error(f"No encontré columnas de Error o ID. Columnas detectadas: {list(df.columns)}")