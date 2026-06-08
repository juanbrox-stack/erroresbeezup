import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io

# Configuración de la página en modo ancho
st.set_page_config(page_title="Extractor de Errores Marketplaces", page_icon="🛒", layout="wide")

st.title("🛒 Extractor Agrupado e Interactivo de Errores")
st.markdown("Pega la URL del reporte HTML para obtener los SKUs organizados con enlaces directos en Excel.")

# --- SECCIÓN DE CONFIGURACIÓN DE MARKETPLACE ---
col_mp, col_url = st.columns([1, 3])

with col_mp:
    marketplace_options = ["MediaMarkt", "CDiscount / Octopia", "Carrefour", "Leroy Merlin", "Decathlon", "Conforama", "Otro (Personalizado)"]
    mp_selected = st.selectbox("Selecciona el Marketplace:", marketplace_options)
    
    if mp_selected == "Otro (Personalizado)":
        marketplace = st.text_input("Escribe el nombre del Marketplace:", "MiMarketplace")
    else:
        marketplace = mp_selected

# Limpiar el nombre para el archivo físico
mp_clean = marketplace.lower().replace(" ", "_").replace("/", "_")

with col_url:
    url = st.text_input(f"URL del reporte HTML de {marketplace}:", placeholder="https://...")

# --- FUNCIÓN DE EXTRACCIÓN ---
def extract_data(url_report):
    try:
        response = requests.get(url_report)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        results = []
        current_error_msg = "Error general"
        
        for element in soup.find_all(['h4', 'li']):
            text = element.get_text(separator=" ", strip=True)
            if element.name == 'h4' and ("error" in text.lower() or "attribute" in text.lower() or "field" in text.lower() or "valeur" in text.lower() or "champ" in text.lower()):
                current_error_msg = text
            
            if text.startswith("SKU "):
                if " : " in text:
                    parts = text.split(" : ", 1)
                    sku = parts[0].replace("SKU ", "").strip()
                    detail = parts[1].strip()
                else:
                    sku = text.replace("SKU ", "").strip()
                    detail = current_error_msg
                
                results.append({'SKU': sku, 'Error': detail})
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Error al procesar la URL: {e}")
        return None

# --- LÓGICA PRINCIPAL ---
if url:
    with st.spinner(f'Procesando y vinculando errores de {marketplace}...'):
        df = extract_data(url)
        
        if df is not None and not df.empty:
            st.success(f"¡Se han encontrado {len(df)} errores en total para {marketplace}!")
            
            # Generar frecuencias para la vista web y el Excel
            df_counts = df['Error'].value_counts().reset_index()
            df_counts.columns = ['Tipo de Error', 'Cantidad de SKUs Afectados']
            
            # --- VISTA PREVIA EN LA WEB APP ---
            st.subheader("🔍 Resumen de errores encontrados")
            st.dataframe(df_counts, use_container_width=True)
            
            st.subheader("📋 Detalle por Bloques de Error")
            # Agrupar en el orden del dataframe de conteos (de mayor a menor)
            for idx, row in df_counts.iterrows():
                error_msg = row['Tipo de Error']
                cant = row['Cantidad de SKUs Afectados']
                group = df[df['Error'] == error_msg]
                
                with st.expander(f"❌ [{idx + 1}] {error_msg} ({cant} SKUs afectados)"):
                    st.dataframe(group[['SKU']].reset_index(drop=True), use_container_width=True)
            
            # --- GENERACIÓN DE EXCEL INTERACTIVO CON XLSXWRITER ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                
                # 1. Pestaña global por si se quiere un volcado general
                df.to_excel(writer, sheet_name='Todos los SKUs', index=False)
                
                # 2. Generar pestañas individuales mapeadas por número (Error_1, Error_2...)
                error_to_sheet = {}
                for i, row in df_counts.iterrows():
                    error_msg = row['Tipo de Error']
                    group = df[df['Error'] == error_msg]
                    sheet_name = f"Error_{i + 1}"
                    
                    group[['SKU']].to_excel(writer, sheet_name=sheet_name, index=False)
                    error_to_sheet[error_msg] = sheet_name
                
                # 3. Pestaña de Resumen Inicial
                df_counts.to_excel(writer, sheet_name='Resumen Errores', index=False)
                
                # Componentes internos de xlsxwriter para inyectar hipervínculos
                workbook  = writer.book
                worksheet = writer.sheets['Resumen Errores']
                
                # Formato visual de enlace (Azul y subrayado)
                link_format = workbook.add_format({
                    'font_color': 'blue',
                    'underline': 1,
                    'font_name': 'Calibri',
                    'font_size': 11
                })
                
                # Escribir las URLs de salto interno en la columna A de la hoja de Resumen
                for idx, row in df_counts.iterrows():
                    error_text = row['Tipo de Error']
                    target_sheet = error_to_sheet.get(error_text, 'Todos los SKUs')
                    row_excel = idx + 1 # +1 debido a la fila de cabecera
                    
                    # Genera la fórmula internal:'NombrePestaña'!A1
                    worksheet.write_url(
                        row_excel, 0, 
                        f"internal:'{target_sheet}'!A1", 
                        string=error_text, 
                        cell_format=link_format
                    )
            
            st.markdown("---")
            st.subheader("💾 Descargar Reporte Optimizado")
            
            file_name_dynamic = f"errores_{mp_clean}.xlsx"
            st.markdown(f"El archivo se descargará listo para usar como: `{file_name_dynamic}`")
            
            st.download_button(
                label=f"📥 Descargar Excel de {marketplace} con Vínculos (.xlsx)",
                data=buffer.getvalue(),
                file_name=file_name_dynamic,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se detectaron datos con el patrón de SKUs en esta URL.")
