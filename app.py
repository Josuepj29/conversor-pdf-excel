import streamlit as st
import pdfplumber
import pandas as pd
import io

# 1. Configuración de la página y metadatos
st.set_page_config(
    page_title="Banca a Excel | Conversor",
    page_icon="📊",
    layout="centered"
)

# 2. CSS Personalizado para un frontend profesional
st.markdown("""
<style>
    /* Ocultar menú y pie de página predeterminados de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Estilizar la zona de carga de archivos */
    .stFileUploader > div > div {
        background-color: #f8f9fa;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 20px;
    }
    
    /* Botón principal */
    .stButton>button {
        width: 100%;
        background-color: #0f172a;
        color: white;
        border-radius: 8px;
        padding: 12px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #334155;
        color: white;
        border-color: #334155;
    }
    
    /* Botón de descarga */
    .stDownloadButton>button {
        width: 100%;
        background-color: #10b981;
        color: white;
        border-radius: 8px;
        padding: 12px;
        font-weight: 600;
    }
    .stDownloadButton>button:hover {
        background-color: #059669;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# 3. Encabezado de la interfaz
st.markdown("<h1 style='text-align: center; color: #0f172a;'>Conversor de Estados de Cuenta</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; font-size: 1.1rem;'>Transforma tus PDFs bancarios en bases de datos Excel limpias en segundos. 100% privado: procesado en memoria.</p>", unsafe_allow_html=True)
st.divider()

# 4. Zona interactiva
archivo_pdf = st.file_uploader("Arrastra aquí tu estado de cuenta (PDF)", type=["pdf"])

if archivo_pdf:
    if st.button("Procesar Documento"):
        with st.spinner('Extrayendo y estructurando datos...'):
            datos_extraidos = []
            
            try:
                # Procesar PDF en memoria
                with pdfplumber.open(archivo_pdf) as pdf:
                    for pagina in pdf.pages:
                        tabla = pagina.extract_table(
                            table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}
                        )
                        if tabla:
                            datos_extraidos.extend(tabla)
                
                if datos_extraidos:
                    df = pd.DataFrame(datos_extraidos)
                    
                    # Limpieza básica de filas vacías
                    df = df.dropna(how='all')
                    
                    # Convertir a Excel en memoria
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, header=False)
                    
                    st.success("✅ ¡Documento estructurado con éxito!")
                    
                    st.download_button(
                        label="📥 Descargar Excel Limpio",
                        data=buffer.getvalue(),
                        file_name="Estado_Cuenta_Procesado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("⚠️ No se detectaron tablas estructuradas en este PDF.")
            except Exception as e:
                st.error("Ocurrió un error al leer el archivo. Verifica que sea un PDF válido.")