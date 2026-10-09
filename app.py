import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
import tempfile
import os
from io import BytesIO

st.set_page_config(page_title="Gestor PDF", page_icon="📄", layout="centered")

st.sidebar.title("Módulos PDF")
opcion = st.sidebar.radio("Selecciona una herramienta:", 
                          ["Comprimir PDF", "Unir y Reorganizar", "Firmar PDF", "Convertir a Word"])

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
                doc = fitz.open(stream=archivo_pdf.read(), filetype="pdf")
                
                # Configurar nivel de compresión (garbage collection y deflate)
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
                
                # Unir todos los PDFs
                for pdf_file in archivos_pdf:
                    doc_temp = fitz.open(stream=pdf_file.read(), filetype="pdf")
                    doc_final.insert_pdf(doc_temp)
                    doc_temp.close()
                
                # Reorganizar si se especificó un orden
                if orden_hojas:
                    try:
                        # Convertir a índice basado en cero
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
    st.info("Nota: Esta herramienta aplica una firma visual (imagen). Para firmas criptográficas con certificado digital (.pfx), se requiere configuración en servidor local.")
    
    archivo_pdf = st.file_uploader("Sube el PDF a firmar", type=["pdf"])
    archivo_firma = st.file_uploader("Sube la imagen de tu firma", type=["png", "jpg", "jpeg"])
    
    if archivo_pdf and archivo_firma:
        doc = fitz.open(stream=archivo_pdf.read(), filetype="pdf")
        num_paginas = len(doc)
        
        pagina_firma = st.number_input("¿En qué página quieres firmar?", min_value=1, max_value=num_paginas, value=num_paginas)
        
        # Opciones de posición predefinidas (esquina inferior derecha es estándar)
        st.write("La firma se colocará en la esquina inferior derecha de la página seleccionada.")
        
        if st.button("Aplicar Firma"):
            with st.spinner("Firmando documento..."):
                page = doc[pagina_firma - 1]
                firma_bytes = archivo_firma.read()
                
                # Calcular coordenadas (Esquina inferior derecha)
                rect = page.rect
                ancho_firma, alto_firma = 150, 50 # Tamaño de la imagen insertada
                margen_x, margen_y = 50, 50
                
                rect_firma = fitz.Rect(
                    rect.width - ancho_firma - margen_x, 
                    rect.height - alto_firma - margen_y, 
                    rect.width - margen_x, 
                    rect.height - margen_y
                )
                
                # Insertar imagen
                page.insert_image(rect_firma, stream=firma_bytes)
                
                salida_pdf = BytesIO()
                doc.save(salida_pdf)
                doc.close()
                
                st.success("¡Firma aplicada!")
                st.download_button(
                    label="Descargar PDF Firmado",
                    data=salida_pdf.getvalue(),
                    file_name=f"firmado_{archivo_pdf.name}",
                    mime="application/pdf"
                )

# ==========================================
# MÓDULO 4: CONVERTIR A WORD
# ==========================================
elif opcion == "Convertir a Word":
    st.header("📝 Convertir PDF a Word")
    archivo_pdf = st.file_uploader("Sube tu archivo PDF", type=["pdf"])
    
    if archivo_pdf is not None:
        if st.button("Convertir a Word"):
            with st.spinner("Convirtiendo... Esto puede tardar unos segundos."):
                # pdf2docx requiere rutas físicas, usamos tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(archivo_pdf.read())
                    ruta_pdf = tmp_pdf.name
                    
                ruta_docx = ruta_pdf.replace(".pdf", ".docx")
                
                try:
                    # Convertir
                    cv = Converter(ruta_pdf)
                    cv.convert(ruta_docx, start=0, end=None)
                    cv.close()
                    
                    # Leer el docx para descargarlo
                    with open(ruta_docx, "rb") as docx_file:
                        docx_bytes = docx_file.read()
                        
                    st.success("¡Conversión exitosa!")
                    st.download_button(
                        label="Descargar Word",
                        data=docx_bytes,
                        file_name=f"{archivo_pdf.name.replace('.pdf', '')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Error durante la conversión: {e}")
                finally:
                    # Limpiar archivos temporales
                    if os.path.exists(ruta_pdf):
                        os.remove(ruta_pdf)
                    if os.path.exists(ruta_docx):
                        os.remove(ruta_docx)
