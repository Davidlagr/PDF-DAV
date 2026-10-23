import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
import tempfile
import os
from io import BytesIO
from PIL import Image
from streamlit_image_coordinates import streamlit_image_coordinates

st.set_page_config(page_title="Gestor PDF", page_icon="📄", layout="centered")

# Inicializar variables de estado para guardar las coordenadas de los clics
if "firma_x" not in st.session_state:
    st.session_state.firma_x = 50.0
if "firma_y" not in st.session_state:
    st.session_state.firma_y = 50.0
if "texto_x" not in st.session_state:
    st.session_state.texto_x = 50.0
if "texto_y" not in st.session_state:
    st.session_state.texto_y = 50.0

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
# MÓDULO 3: FIRMAR PDF (Clic interactivo)
# ==========================================
elif opcion == "Firmar PDF":
    st.header("✍️ Firmar PDF")
    st.info("Sube tu documento y tu firma. Haz clic en la vista previa del documento para ubicar la firma.")
    
    archivo_pdf = st.file_uploader("Paso 1: Sube el PDF a firmar", type=["pdf"], key="pdf_firma")
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        st.write("### Selecciona la página")
        pagina_firma = st.number_input("Página a firmar:", min_value=1, max_value=num_paginas, value=1)
        doc_info.close()

        archivo_firma = st.file_uploader("Paso 2: Sube la imagen de tu firma", type=["png", "jpg", "jpeg"], key="img_firma")
        
        if archivo_firma:
            firma_bytes = archivo_firma.getvalue()
            
            st.write("### Tamaño de la Firma")
            col_x, col_y = st.columns(2)
            with col_x:
                ancho_firma = st.slider("Ancho", 10.0, 500.0, 150.0, key="firma_ancho")
            with col_y:
                alto_firma = st.slider("Alto", 10.0, 500.0, 50.0, key="firma_alto")

            st.write("### Haz clic en el documento para colocar tu firma")
            
            # Generar vista previa fresca en cada interacción
            doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_preview = doc_preview[pagina_firma - 1]
            
            # Insertar la firma en la posición guardada en session_state
            rect_firma = fitz.Rect(st.session_state.firma_x, st.session_state.firma_y, 
                                   st.session_state.firma_x + ancho_firma, st.session_state.firma_y + alto_firma)
            page_preview.insert_image(rect_firma, stream=firma_bytes)
            
            # Renderizar página a escala 1:1 para coordenadas precisas
            pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Capturar clic del usuario
            coords = streamlit_image_coordinates(img_pil, key="coordenadas_firma")
            
            # Si el usuario hace clic en una nueva posición, actualizar estado y recargar
            if coords is not None:
                if coords["x"] != st.session_state.firma_x or coords["y"] != st.session_state.firma_y:
                    st.session_state.firma_x = float(coords["x"])
                    st.session_state.firma_y = float(coords["y"])
                    st.rerun()

            doc_preview.close()
            
            if st.button("Aplicar Firma y Descargar"):
                with st.spinner("Procesando documento final..."):
                    doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                    page_final = doc_final[pagina_firma - 1]
                    page_final.insert_image(rect_firma, stream=firma_bytes)
                    
                    salida_pdf = BytesIO()
                    doc_final.save(salida_pdf)
                    doc_final.close()
                    
                    st.success("¡Firma aplicada con éxito!")
                    st.download_button(
                        label="Descargar PDF Firmado",
                        data=salida_pdf.getvalue(),
                        file_name=f"firmado_{archivo_pdf.name}",
                        mime="application/pdf"
                    )

# ==========================================
# MÓDULO 4: EDITAR Y RELLENAR PDF (Clic interactivo)
# ==========================================
elif opcion == "Editar y Rellenar PDF":
    st.header("📝 Editar y Rellenar PDF")
    st.info("Escribe tu texto y haz clic en el documento para ubicarlo.")
    
    archivo_pdf = st.file_uploader("Sube el PDF a rellenar", type=["pdf"])
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        pagina_texto = st.number_input("Página:", min_value=1, max_value=num_paginas, value=1)
        texto_insertar = st.text_input("Texto a insertar:", "Escribe tu texto aquí")
        doc_info.close()

        st.write("### Ajustes del Texto")
        col_x, col_y = st.columns(2)
        with col_x:
            tamano_fuente = st.slider("Tamaño de fuente", 8, 72, 12)
        with col_y:
            color_texto = st.color_picker("Color del texto", "#000000")

        rgb_color = tuple(int(color_texto.lstrip('#')[i:i+2], 16)/255 for i in (0, 2, 4))

        st.write("### Haz clic en el documento para colocar el texto")
        
        doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_preview = doc_preview[pagina_texto - 1]
        
        # PyMuPDF usa el punto inferior izquierdo como inicio del texto
        p = fitz.Point(st.session_state.texto_x, st.session_state.texto_y)
        page_preview.insert_text(p, texto_insertar, fontsize=tamano_fuente, color=rgb_color)
        
        pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
        img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Capturar clic
        coords = streamlit_image_coordinates(img_pil, key="coordenadas_texto")
        
        if coords is not None:
            if coords["x"] != st.session_state.texto_x or coords["y"] != st.session_state.texto_y:
                st.session_state.texto_x = float(coords["x"])
                st.session_state.texto_y = float(coords["y"])
                st.rerun()

        doc_preview.close()
        
        if st.button("Aplicar Texto y Descargar"):
            with st.spinner("Procesando documento final..."):
                doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                page_final = doc_final[pagina_texto - 1]
                page_final.insert_text(p, texto_insertar, fontsize=tamano_fuente, color=rgb_color)
                
                salida_pdf = BytesIO()
                doc_final.save(salida_pdf)
                doc_final.close()
                
                st.success("¡Texto aplicado con éxito!")
                st.download_button(
                    label="Descargar PDF Rellenado",
                    data=salida_pdf.getvalue(),
                    file_name=f"rellenado_{archivo_pdf.name}",
                    mime="application/pdf"
                )

# ==========================================
# MÓDULO 5: CONVERTIR A WORD
# ==========================================
elif opcion == "Convertir a Word":
    st.header("📝 Convertir PDF a Word")
    archivo_pdf = st.file_uploader("Sube tu archivo PDF", type=["pdf"])
    
    if archivo_pdf is not None:
        if st.button("Convertir a Word"):
            with st.spinner("Convirtiendo... Esto puede tardar unos segundos."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(archivo_pdf.getvalue())
                    ruta_pdf = tmp_pdf.name
                    
                ruta_docx = ruta_pdf.replace(".pdf", ".docx")
                
                try:
                    cv = Converter(ruta_pdf)
                    cv.convert(ruta_docx, start=0, end=None)
                    cv.close()
                    
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
                    if os.path.exists(ruta_pdf):
                        os.remove(ruta_pdf)
                    if os.path.exists(ruta_docx):
                        os.remove(ruta_docx)
