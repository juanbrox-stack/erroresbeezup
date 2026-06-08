import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io

st.set_page_config(page_title="Extractor de Errores Marketplaces", page_icon="🛒", layout="wide")

st.title("🛒 Extractor Agrupado de Errores BeezUP")
st.markdown("Pega la URL del reporte HTML para obtener los SKUs agrupados por el tipo de error.")

# --- NUEVA SECCIÓN: SELECCIÓN DE MARKETPLACE ---
col_mp, col_url = st.columns([1, 3])

with col_mp:
    marketplace_options = ["MediaMarkt", "Carrefour", "Leroy Merlin", "Decathlon", "Conforama", "Cdiscoungt", "Otro (Personalizado)"]
    mp_selected = st.selectbox("Selecciona el Marketplace:", marketplace_options)
    
    # Si elige "Otro", permitimos que escriba el nombre a mano
    if mp_selected == "Otro (Personalizado)":
        marketplace = st.text_input("Escribe el nombre del Marketplace:", "MiMarketplace")
    else:
        marketplace = mp_selected

# Limpiamos el nombre para evitar problemas en el nombre del archivo (quitar espacios)
mp_clean = marketplace.lower().replace(" ", "_")

with col_url:
    url = st.text_input(f"URL del reporte HTML de {marketplace}:", placeholder="https://beezupmp2publication3.blob.core.windows.net/...")

def extract_data(url_report):
    try:
        response = requests.get(url_report)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        results = []
        current_error_msg = "Error general"
        
        for element in soup.find_all(['h4', 'li']):
            text = element.get_text(separator=" ", strip=True)
            if element.name == 'h4' and ("error" in text.lower() or "attribute" in text.lower()):
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

if url:
    with st.spinner(f'Procesando y agrupando errores de {marketplace}...'):
        df = extract_data(url)
        
        if df is not None and not df.empty:
            st.success(f"¡Se han encontrado {len(df)} errores en total para {marketplace}!")
            
            # --- SECCIÓN DE VISTA PREVIA AGRUPADA ---
            st.subheader("🔍 Resumen de errores encontrados")
            
            df_counts = df['Error'].value_counts().reset_index()
            df_counts.columns = ['Tipo de Error', 'Cantidad de SKUs Afectados']
            st.dataframe(df_counts, use_container_width=True)
            
            st.subheader("📋 Detalle por Bloques de Error")
            for error_msg, group in df.groupby('Error'):
                with st.expander(f"❌ {error_msg} ({len(group)} SKUs afectados)"):
                    st.write("SKUs afectados:")
                    st.dataframe(group[['SKU']].reset_index(drop=True), use_container_width=True)
            
            # --- GENERACIÓN DE EXCEL ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_counts.to_excel(writer, sheet_name='Resumen Errores', index=False)
                df.to_excel(writer, sheet_name='Todos los SKUs', index=False)
                
                for i, (error_msg, group) in enumerate(df.groupby('Error'), 1):
                    sheet_name = f"Error_{i}" 
                    group[['SKU']].to_excel(writer, sheet_name=sheet_name, index=False)
            
            st.markdown("---")
            st.subheader("💾 Descargar Reporte Optimizado")
            
            # NOMBRE DEL ARCHIVO DINÁMICO
            file_name_dynamic = f"errores_{mp_clean}.xlsx"
            
            st.markdown(f"El archivo se descargará automáticamente como: `{file_name_dynamic}`")
            
            st.download_button(
                label=f"📥 Descargar Excel de {marketplace} (.xlsx)",
                data=buffer.getvalue(),
                file_name=file_name_dynamic,
                mime="application/vnd.ms-excel"
            )
        else:
            st.warning("No se encontraron SKUs en esta URL.")
