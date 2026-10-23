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
# MÓDULO 3: FIRMAR PDF (Corregido y Optimizado)
# ==========================================
elif opcion == "Firmar PDF":
    st.header("✍️ Firmar PDF")
    st.info("Sube tu documento y la imagen de tu firma. Ajusta las coordenadas y visualiza el resultado en tiempo real.")
    
    # 1. Subir solo el PDF primero para dar retroalimentación visual inmediata
    archivo_pdf = st.file_uploader("Paso 1: Sube el PDF a firmar", type=["pdf"], key="pdf_firma")
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        st.write("### Selecciona la página")
        pagina_firma = st.number_input("Página a firmar:", min_value=1, max_value=num_paginas, value=1)
        
        # Dimensiones seguras de la página
        page_info = doc_info[pagina_firma - 1]
        max_width = float(page_info.rect.width)
        max_height = float(page_info.rect.height)
        doc_info.close()

        # 2. Subir la firma
        archivo_firma = st.file_uploader("Paso 2: Sube la imagen de tu firma", type=["png", "jpg", "jpeg"], key="img_firma")
        
        if archivo_firma:
            firma_bytes = archivo_firma.getvalue()
            
            st.write("### Ajustes de Posición y Tamaño")
            col_x, col_y = st.columns(2)
            
            # Cálculos seguros para evitar errores si la página es pequeña
            default_x = max(0.0, min(max_width, max_width - 160.0))
            default_y = max(0.0, min(max_height, max_height - 100.0))
            
            with col_x:
                pos_x = st.slider("Posición Horizontal (X)", 0.0, max_width, default_x, key="firma_x")
                ancho_firma = st.slider("Ancho de la firma", 10.0, max(11.0, max_width), min(150.0, max_width), key="firma_ancho")
            with col_y:
                pos_y = st.slider("Posición Vertical (Y)", 0.0, max_height, default_y, key="firma_y")
                alto_firma = st.slider("Alto de la firma", 10.0, max(11.0, max_height), min(50.0, max_height), key="firma_alto")

            st.write("### Vista Previa con Firma")
            doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_preview = doc_preview[pagina_firma - 1]
            
            rect_firma = fitz.Rect(pos_x, pos_y, pos_x + ancho_firma, pos_y + alto_firma)
            page_preview.insert_image(rect_firma, stream=firma_bytes)
            
            pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            st.image(pix.tobytes("png"), use_container_width=True)
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
        else:
            # Mostrar documento mientras se espera la firma
            st.warning("Sube la imagen de tu firma para poder ajustarla.")
            st.write("### Vista Previa del Documento (Sin firma)")
            doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_preview = doc_preview[pagina_firma - 1]
            pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            st.image(pix.tobytes("png"), use_container_width=True)
            doc_preview.close()

# ==========================================
# MÓDULO 4: EDITAR Y RELLENAR PDF
# ==========================================
elif opcion == "Editar y Rellenar PDF":
    st.header("📝 Editar y Rellenar PDF")
    st.info("Añade texto directamente sobre cualquier página del PDF (ideal para rellenar formularios).")
    
    archivo_pdf = st.file_uploader("Sube el PDF a rellenar", type=["pdf"])
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        pagina_texto = st.number_input("Página donde insertarás el texto:", min_value=1, max_value=num_paginas, value=1)
        texto_insertar = st.text_input("Texto a insertar:", "Ejemplo de texto")
        
        page_info = doc_info[pagina_texto - 1]
        max_width = float(page_info.rect.width)
        max_height = float(page_info.rect.height)
        doc_info.close()

        st.write("### Ajustes de Posición del Texto")
        col_x, col_y = st.columns(2)
        
        default_x_txt = min(max_width, 50.0)
        default_y_txt = min(max_height, 50.0)
        
        with col_x:
            pos_x_txt = st.slider("Posición Horizontal (X)", 0.0, max_width, default_x_txt, key="txt_x")
            tamano_fuente = st.slider("Tamaño de fuente", 8, 72, 12)
        with col_y:
            pos_y_txt = st.slider("Posición Vertical (Y)", 0.0, max_height, default_y_txt, key="txt_y")
            color_texto = st.color_picker("Color del texto", "#000000")

        rgb_color = tuple(int(color_texto.lstrip('#')[i:i+2], 16)/255 for i in (0, 2, 4))

        st.write("### Vista Previa")
        doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_preview = doc_preview[pagina_texto - 1]
        
        p = fitz.Point(pos_x_txt, pos_y_txt)
        page_preview.insert_text(p, texto_insertar, fontsize=tamano_fuente, color=rgb_color)
        
        pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        st.image(pix.tobytes("png"), use_container_width=True)
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
