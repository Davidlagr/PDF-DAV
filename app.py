import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
import tempfile
import os
from io import BytesIO

st.set_page_config(page_title="Gestor PDF", page_icon="📄", layout="centered")

st.sidebar.title("Módulos PDF")
opcion = st.sidebar.radio("Selecciona una herramienta:", 
                          ["Comprimir PDF", "Unir y Reorganizar", "Firmar PDF", "Editar y Rellenar PDF", "Convertir a Word"])

# ==========================================
# MÓDULO 1: COMPRIMIR PDF
# ==========================================
if opcion == "Comprimir PDF":
    st.header("🗜️ Comprimir PDF")
    archivo_pdf = st.file_uploader("Sube tu archivo PDF", type=["pdf"])
    
    if archivo_pdf is not None:
        nivel = st.radio("Nivel de compresión:", ["Medio (Buena calidad)", "Máximo (Calidad media)"])
        
        if st.button("Comprimir"):
            with st.spinner("Comprimiendo..."):
                doc = fitz.open(stream=archivo_pdf.getvalue(), filetype="pdf")
                
                if nivel == "Medio (Buena calidad)":
                    garbage_level = 3
                    deflate = True
                else:
                    garbage_level = 4
                    deflate = True
                
                salida_pdf = BytesIO()
                doc.save(salida_pdf, garbage=garbage_level, deflate=deflate)
                doc.close()
                
                st.success("¡Compresión finalizada!")
                st.download_button(
                    label="Descargar PDF Comprimido",
                    data=salida_pdf.getvalue(),
                    file_name=f"comprimido_{archivo_pdf.name}",
                    mime="application/pdf"
                )

# ==========================================
# MÓDULO 2: UNIR Y REORGANIZAR
# ==========================================
elif opcion == "Unir y Reorganizar":
    st.header("🔗 Unir y Reorganizar PDFs")
    archivos_pdf = st.file_uploader("Sube los PDFs a unir (en el orden deseado)", type=["pdf"], accept_multiple_files=True)
    
    if archivos_pdf:
        st.write("Archivos cargados:")
        for i, f in enumerate(archivos_pdf):
            st.write(f"{i+1}. {f.name}")
            
        orden_hojas = st.text_input("Reorganizar hojas (Ej: 1,3,2,4,5). Deja en blanco para mantener el orden original.")
        
        if st.button("Unir y Procesar"):
            with st.spinner("Uniendo documentos..."):
                doc_final = fitz.open()
                
                for pdf_file in archivos_pdf:
                    doc_temp = fitz.open(stream=pdf_file.getvalue(), filetype="pdf")
                    doc_final.insert_pdf(doc_temp)
                    doc_temp.close()
                
                if orden_hojas:
                    try:
                        nuevo_orden = [int(x.strip()) - 1 for x in orden_hojas.split(",")]
                        doc_final.select(nuevo_orden)
                    except Exception as e:
                        st.error(f"Error en el formato de orden: {e}. Asegúrate de usar números separados por comas y que las páginas existan.")
                        st.stop()
                
                salida_pdf = BytesIO()
                doc_final.save(salida_pdf)
                doc_final.close()
                
                st.success("¡Documento procesado!")
                st.download_button(
                    label="Descargar PDF Unido",
                    data=salida_pdf.getvalue(),
                    file_name="documento_unido.pdf",
                    mime="application/pdf"
                )

# ==========================================
# MÓDULO 3: FIRMAR PDF 
# ==========================================
elif opcion == "Firmar PDF":
    st.header("✍️ Firmar PDF")
    st.info("Ajusta las coordenadas de tu firma usando los controles. La vista previa se actualizará en tiempo real.")
    
    archivo_pdf = st.file_uploader("Sube el PDF a firmar", type=["pdf"])
    archivo_firma = st.file_uploader("Sube la imagen de tu firma", type=["png", "jpg", "jpeg"])
    
    if archivo_pdf and archivo_firma:
        pdf_bytes = archivo_pdf.getvalue()
        firma_bytes = archivo_firma.getvalue()
        
        # Abrir documento temporal solo para leer dimensiones de página
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        col1, col2 = st.columns([1, 1])
        with col1:
            pagina_firma = st.number_input("Página a firmar:", min_value=1, max_value=num_paginas, value=1)
        
        # Obtener dimensiones de la página seleccionada
        page_info = doc_info[pagina_firma - 1]
        max_width = page_info.rect.width
        max_height = page_info
