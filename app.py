import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Extractor de PDF", layout="centered")

# --- EL LOGO SIEMPRE ARRIBA ---
# Al ponerlo en 900, le quitamos las columnas para que ocupe un buen espacio
st.image("logo.jpg", width=900) 
st.write("---") 

# --- 1. SISTEMA DE LOGIN ---
# Correos oficiales del equipo. Puedes cambiar "sintesis2026" por sus contraseñas reales.
USUARIOS_PERMITIDOS = {
    "karinao@sintesis.com": "acceso123",
    "nicolf@sintesis.com.bo": "Sintesis123",
    "manuelm@sintesis.com.bo": "Junio2026"
}

if 'logeado' not in st.session_state:
    st.session_state.logeado = False

# Si NO está logeado, mostramos solo el formulario
if not st.session_state.logeado:
    st.title("🔒 Acceso Restringido")
    st.write("Por favor, inicia sesión con tu correo corporativo.")
    
    email = st.text_input("Correo electrónico")
    password = st.text_input("Contraseña", type="password")
    
    if st.button("Iniciar Sesión"):
        if email in USUARIOS_PERMITIDOS and USUARIOS_PERMITIDOS[email] == password:
            st.session_state.logeado = True
            st.rerun()
        else:
            st.error("❌ Correo o contraseña incorrectos.")

# --- 2. EL PORTAL (Solo se muestra si el login fue exitoso) ---
else:
    # Botón para cerrar sesión en la esquina superior derecha
    col_vacia, col_salir = st.columns([4, 1])
    with col_salir:
        if st.button("Cerrar Sesión"):
            st.session_state.logeado = False
            st.session_state.procesado = False
            st.session_state.excel_data = None
            
            # Forzamos un refresco real del navegador para que el logo vuelva a cargar
            st.components.v1.html("""
                <script>
                    window.parent.location.reload();
                </script>
            """, height=0)
            st.stop()

    st.title("Extractor de Extractos Bancarios a Excel 📊")
    st.write("Sube tu archivo PDF y selecciona las páginas que deseas extraer.")

    if 'procesado' not in st.session_state:
        st.session_state.procesado = False
        st.session_state.excel_data = None
        st.session_state.total_filas = 0
        
    if 'uploader_key' not in st.session_state:
        st.session_state.uploader_key = 1

    uploaded_file = st.file_uploader("1. Sube tu archivo PDF aquí", type="pdf", key=f"uploader_{st.session_state.uploader_key}")

    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input("Página Inicial", min_value=1, value=1334)
    with col2:
        end_page = st.number_input("Página Final", min_value=1, value=1465)

    if uploaded_file is not None:
        if not st.session_state.procesado:
            if st.button("2. Procesar y Extraer Datos"):
                st.info("Procesando el texto de las páginas... ¡No cierres esta ventana!")
                all_data = []
                
                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        total_pages = len(pdf.pages)
                        
                        if start_page > total_pages or end_page > total_pages or start_page > end_page:
                            st.error(f"Error en las páginas. El PDF solo tiene {total_pages} páginas.")
                        else:
                            patron = r"^(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(.+?)\s+([\-\d,]+\.\d{2})\s+([\-\d,]+\.\d{2})\s+([\-\d,]+\.\d{2})\s*$"
                            
                            for i in range(start_page - 1, end_page):
                                page = pdf.pages[i]
                                text = page.extract_text()
                                
                                if text:
                                    lineas = text.split('\n')
                                    for linea in lineas:
                                        match = re.match(patron, linea.strip())
                                        if match:
                                            all_data.append(list(match.groups()))
                    
                    if all_data:
                        columnas = ["FECHA", "N.DOC", "DESCRIPCION", "DEBE", "HABER", "SALDO"]
                        df = pd.DataFrame(all_data, columns=columnas)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Extracto')
                        
                        st.session_state.excel_data = output.getvalue()
                        st.session_state.total_filas = len(df)
                        st.session_state.procesado = True
                        st.rerun()
                        
                    else:
                        st.warning("No se encontraron transacciones en estas páginas.")
                        
                except Exception as e:
                    st.error(f"Ocurrió un error: {e}")

        if st.session_state.procesado:
            st.success(f"¡Éxito! Se extrajeron {st.session_state.total_filas} transacciones listas para Excel.")
            
            st.download_button(
                label="📥 3. Descargar archivo Excel",
                data=st.session_state.excel_data,
                file_name=f"Extracto_{start_page}_a_{end_page}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.write("---")
            
            if st.button("🔄 Procesar nuevo archivo / Limpiar pantalla"):
                st.session_state.procesado = False
                st.session_state.excel_data = None
                st.session_state.total_filas = 0
                st.session_state.uploader_key += 1 
                st.rerun()
