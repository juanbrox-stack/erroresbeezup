import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io

st.set_page_config(page_title="Extractor de Errores MediaMarkt", page_icon="🛒")

st.title("🛒 Extractor de Errores Mirakl/MediaMarkt")
st.markdown("Pega la URL del reporte HTML de BeezUP para extraer los SKUs y sus errores en una tabla limpia.")

# Input de la URL
url = st.text_input("URL del reporte HTML:", placeholder="https://beezupmp2publication3.blob.core.windows.net/...")

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
    with st.spinner('Extrayendo datos...'):
        df = extract_data(url)
        
        if df is not None and not df.empty:
            st.success(f"¡Se han encontrado {len(df)} errores!")
            
            # Mostrar vista previa
            st.dataframe(df, use_container_width=True)
            
            # Botón para descargar en EXCEL (para evitar el problema de las comas)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Errores')
            
            st.download_button(
                label="📥 Descargar como Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="errores_marketplaces.xlsx",
                mime="application/vnd.ms-excel"
            )
        else:
            st.warning("No se encontraron SKUs en esta URL. Asegúrate de que el formato sea el correcto.")