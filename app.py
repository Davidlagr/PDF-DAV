import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
import tempfile
import os
import random
import string
from io import BytesIO
from PIL import Image, ImageEnhance
from streamlit_image_coordinates import streamlit_image_coordinates
import pdfplumber
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from pyhanko.sign import signers
from pyhanko.pdf_utils.reader import PdfFileReader
from pyhanko.sign import fields

st.set_page_config(page_title="Gestor PDF", page_icon="📄", layout="wide") 

# --- INICIALIZACIÓN DE VARIABLES DE ESTADO ---
if "firma_x" not in st.session_state: st.session_state.firma_x = 50.0
if "firma_y" not in st.session_state: st.session_state.firma_y = 50.0
if "last_click_firma" not in st.session_state: st.session_state.last_click_firma = None

if "texto_x" not in st.session_state: st.session_state.texto_x = 50.0
if "texto_y" not in st.session_state: st.session_state.texto_y = 50.0
if "last_click_texto" not in st.session_state: st.session_state.last_click_texto = None

if "textos_agregados" not in st.session_state: st.session_state.textos_agregados = []
if "pdf_actual_form" not in st.session_state: st.session_state.pdf_actual_form = None

def update_firma_pos():
    st.session_state.firma_x = st.session_state.slider_firma_x
    st.session_state.firma_y = st.session_state.slider_firma_y

def update_texto_pos():
    st.session_state.texto_x = st.session_state.slider_texto_x
    st.session_state.texto_y = st.session_state.slider_texto_y


st.sidebar.title("Módulos PDF")
opcion = st.sidebar.radio("Selecciona una herramienta:", 
                          ["Comprimir PDF", 
                           "Unir y Reorganizar", 
                           "Dividir o Quitar Hojas",
                           "Firmar PDF", 
                           "Editar y Rellenar PDF", 
                           "Marca de Agua",
                           "Proteger / Desbloquear PDF",
                           "Convertir a Word, Excel, PPT"])

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
                garbage_level = 3 if nivel == "Medio (Buena calidad)" else 4
                salida_pdf = BytesIO()
                doc.save(salida_pdf, garbage=garbage_level, deflate=True)
                doc.close()
                st.success("¡Compresión finalizada!")
                st.download_button(label="Descargar PDF Comprimido", data=salida_pdf.getvalue(), file_name=f"comprimido_{archivo_pdf.name}", mime="application/pdf")

# ==========================================
# MÓDULO 2: UNIR Y REORGANIZAR
# ==========================================
elif opcion == "Unir y Reorganizar":
    st.header("🔗 Unir y Reorganizar PDFs")
    archivos_pdf = st.file_uploader("Sube los PDFs a unir", type=["pdf"], accept_multiple_files=True)
    if archivos_pdf:
        st.write("### 1. Ordenar Archivos Subidos")
        nombres_archivos = [f.name for f in archivos_pdf]
        orden_archivos = st.multiselect("Orden de los documentos:", options=nombres_archivos, default=nombres_archivos)
        
        archivos_ordenados = []
        for name in orden_archivos:
            for f in archivos_pdf:
                if f.name == name and f not in archivos_ordenados:
                    archivos_ordenados.append(f)
                    break

        st.write("### 2. Ordenar páginas internas (Opcional)")
        orden_hojas = st.text_input("Reorganizar todas las hojas finales (Ej: 1,3,2,4,5). Deja en blanco para mantener el orden.")
        if st.button("Unir y Procesar"):
            with st.spinner("Uniendo documentos..."):
                doc_final = fitz.open()
                for pdf_file in archivos_ordenados:
                    doc_temp = fitz.open(stream=pdf_file.getvalue(), filetype="pdf")
                    doc_final.insert_pdf(doc_temp)
                    doc_temp.close()
                if orden_hojas:
                    try:
                        nuevo_orden = [int(x.strip()) - 1 for x in orden_hojas.split(",")]
                        doc_final.select(nuevo_orden)
                    except:
                        st.error("Error en el formato de orden.")
                        st.stop()
                salida_pdf = BytesIO()
                doc_final.save(salida_pdf)
                doc_final.close()
                st.success("¡Documento procesado!")
                st.download_button(label="Descargar PDF Unido", data=salida_pdf.getvalue(), file_name="documento_unido.pdf", mime="application/pdf")

# ==========================================
# MÓDULO 3: DIVIDIR O QUITAR HOJAS
# ==========================================
elif opcion == "Dividir o Quitar Hojas":
    st.header("✂️ Dividir o Quitar Hojas de un PDF")
    archivo_pdf = st.file_uploader("Sube tu archivo PDF", type=["pdf"])
    if archivo_pdf:
        doc = fitz.open(stream=archivo_pdf.getvalue(), filetype="pdf")
        num_paginas = len(doc)
        st.info(f"El documento tiene **{num_paginas} páginas** (numeradas del 1 al {num_paginas}).")
        
        modo = st.radio("¿Qué deseas hacer?", ["Conservar páginas (Extraer)", "Eliminar páginas (Quitar)"])
        paginas_input = st.text_input("Ingresa las páginas separadas por comas o rangos (Ej: 1, 3, 5-8):")
        
        if st.button("Procesar Documento"):
            with st.spinner("Procesando páginas..."):
                paginas_seleccionadas = set()
                for part in paginas_input.split(','):
                    part = part.strip()
                    if '-' in part:
                        try:
                            start, end = map(int, part.split('-'))
                            paginas_seleccionadas.update(range(start, end + 1))
                        except: pass
                    else:
                        try:
                            if part: paginas_seleccionadas.add(int(part))
                        except: pass
                
                idx_seleccionados = sorted([p - 1 for p in paginas_seleccionadas if 1 <= p <= num_paginas])
                idx_finales = [i for i in range(num_paginas) if i not in idx_seleccionados] if modo == "Eliminar páginas (Quitar)" else idx_seleccionados
                
                if not idx_finales:
                    st.error("La selección dejó el documento sin páginas.")
                else:
                    doc.select(idx_finales)
                    salida_pdf = BytesIO()
                    doc.save(salida_pdf)
                    doc.close()
                    st.success("¡Páginas procesadas con éxito!")
                    st.download_button(label="Descargar PDF Modificado", data=salida_pdf.getvalue(), file_name=f"modificado_{archivo_pdf.name}", mime="application/pdf")

# ==========================================
# MÓDULO 4: FIRMAR PDF (VISUAL Y DIGITAL)
# ==========================================
elif opcion == "Firmar PDF":
    st.header("✍️ Firmar PDF")
    tipo_firma = st.radio("Elige el tipo de firma:", ["Firma Visual (Imagen en el documento)", "Firma Digital Criptográfica (Archivo .p12 / .pfx)"])
    
    if tipo_firma == "Firma Visual (Imagen en el documento)":
        archivo_pdf = st.file_uploader("Paso 1: Sube el PDF a firmar", type=["pdf"], key="pdf_firma")
        if archivo_pdf:
            pdf_bytes = archivo_pdf.getvalue()
            doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
            num_paginas = len(doc_info)
            
            st.write("### Selecciona la página")
            pagina_firma = st.number_input("Página a firmar:", min_value=1, max_value=num_paginas, value=1)
            page_info = doc_info[pagina_firma - 1]
            max_width = float(page_info.rect.width)
            max_height = float(page_info.rect.height)
            doc_info.close()

            archivo_firma = st.file_uploader("Paso 2: Sube la imagen de tu firma", type=["png", "jpg", "jpeg"], key="img_firma")
            if archivo_firma:
                firma_img = Image.open(archivo_firma).convert("RGBA")
                quitar_fondo = st.checkbox("Quitar fondo blanco de la firma")
                
                if quitar_fondo:
                    datas = firma_img.getdata()
                    newData = []
                    for item in datas:
                        if item[0] > 200 and item[1] > 200 and item[2] > 200: newData.append((255, 255, 255, 0))
                        else: newData.append(item)
                    firma_img.putdata(newData)
                
                img_byte_arr = BytesIO()
                firma_img.save(img_byte_arr, format='PNG')
                firma_bytes = img_byte_arr.getvalue()
                
                st.write("### Controles de Tamaño y Posición")
                col_t1, col_t2 = st.columns(2)
                with col_t1: ancho_firma = st.slider("Ancho de firma", 10.0, 500.0, 150.0, key="firma_ancho")
                with col_t2: alto_firma = st.slider("Alto de firma", 10.0, 500.0, 50.0, key="firma_alto")
                    
                col_p1, col_p2 = st.columns(2)
                with col_p1: st.slider("Arrastrar Horizontalmente (X)", 0.0, max_width, value=st.session_state.firma_x, key="slider_firma_x", on_change=update_firma_pos)
                with col_p2: st.slider("Arrastrar Verticalmente (Y)", 0.0, max_height, value=st.session_state.firma_y, key="slider_firma_y", on_change=update_firma_pos)

                st.write("### Vista Previa Interactiva (Clic para mover)")
                doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
                page_preview = doc_preview[pagina_firma - 1]
                rect_firma = fitz.Rect(st.session_state.firma_x, st.session_state.firma_y, st.session_state.firma_x + ancho_firma, st.session_state.firma_y + alto_firma)
                page_preview.insert_image(rect_firma, stream=firma_bytes)
                
                pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
                img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                coords = streamlit_image_coordinates(img_pil, key="click_firma")
                if coords is not None:
                    click_id = f"{coords['x']}-{coords['y']}"
                    if st.session_state.last_click_firma != click_id:
                        st.session_state.last_click_firma = click_id
                        st.session_state.firma_x = float(coords["x"])
                        st.session_state.firma_y = float(coords["y"])
                        st.rerun()
                doc_preview.close()
                
                st.write("### Opciones de Seguridad")
                bloquear_pdf = st.checkbox("🔒 Bloquear PDF contra modificaciones
