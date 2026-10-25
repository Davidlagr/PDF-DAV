import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
import tempfile
import os
import random
import string
import traceback
from io import BytesIO
from PIL import Image, ImageEnhance
from streamlit_image_coordinates import streamlit_image_coordinates
import pdfplumber
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

# PyHanko y Cryptography
from pyhanko.sign import signers
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from cryptography.hazmat.primitives.serialization import pkcs12

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
# MÓDULO 4: FIRMAR PDF (VISUAL Y DIGITAL CON DIAGNÓSTICO)
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
                bloquear_pdf = st.checkbox("🔒 Bloquear PDF contra modificaciones (Impide que se borre la firma)")

                if st.button("Aplicar Firma y Descargar"):
                    with st.spinner("Procesando documento final..."):
                        doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                        page_final = doc_final[pagina_firma - 1]
                        page_final.insert_image(rect_firma, stream=firma_bytes)
                        salida_pdf = BytesIO()
                        if bloquear_pdf:
                            owner_pw = ''.join(random.choices(string.ascii_letters + string.digits, k=16))
                            perms = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_ACCESSIBILITY
                            doc_final.save(salida_pdf, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=owner_pw, permissions=perms)
                        else:
                            doc_final.save(salida_pdf)
                        doc_final.close()
                        st.success("¡Firma aplicada con éxito!")
                        st.download_button(label="Descargar PDF Firmado", data=salida_pdf.getvalue(), file_name=f"firmado_{archivo_pdf.name}", mime="application/pdf")
    
    else:
        st.info("Aplica una firma digital con validez legal subiendo tu certificado PKCS#12 (.p12 o .pfx).")
        
        col_firma1, col_firma2 = st.columns(2)
        with col_firma1:
            pdf_subido = st.file_uploader("1. Sube el PDF a firmar", type=["pdf"])
        with col_firma2:
            archivo_p12 = st.file_uploader("2. Sube tu certificado (.p12 / .pfx)", type=["p12", "pfx"])
        
        password_certificado = st.text_input("3. Contraseña del certificado:", type="password")

        if archivo_p12 and password_certificado:
            if st.button("Diagnosticar Certificado"):
                try:
                    private_key, cert, extras = pkcs12.load_key_and_certificates(
                        archivo_p12.getvalue(),
                        password_certificado.encode("utf-8")
                    )
                    st.success("Certificado cargado correctamente")
                    st.write("Clave privada:", private_key is not None)
                    st.write("Certificado:", cert is not None)
                    if cert:
                        st.code(str(cert.subject))
                except Exception as e:
                    st.error(f"ERROR REAL: {str(e)}")
                    st.code(traceback.format_exc())

       if pdf_subido and archivo_p12 and password_certificado:
            if st.button("Firmar Documento Digitalmente", type="primary"):
                with st.spinner("Procesando firma criptográfica en memoria..."):
                    try:
                        # 1. Extracción nativa (la misma que funcionó en el diagnóstico)
                        private_key, cert, extras = pkcs12.load_key_and_certificates(
                            archivo_p12.getvalue(),
                            password_certificado.encode("utf-8")
                        )
                        
                        if private_key is None:
                            st.error("El certificado no contiene clave privada")
                            st.stop()
                        if cert is None:
                            st.error("No se encontró certificado")
                            st.stop()
                        
                        # 2. Armar el Signer directamente con los objetos extraídos (Esquivando load_pkcs12)
                        signer = signers.SimpleSigner(
                            signing_cert=cert,
                            signing_key=private_key,
                            cert_registry=signers.SimpleCertificateStore()
                        )
                        
                        # 3. Preparar el PDF
                        pdf_writer = IncrementalPdfFileWriter(
                            BytesIO(pdf_subido.getvalue())
                        )
                        
                        output = BytesIO()
                        
                        # 4. Firmar e inyectar
                        signers.sign_pdf(
                            pdf_writer, 
                            signature_meta=signers.PdfSignatureMetadata(
                                field_name="FirmaDigitalLegal"
                            ), 
                            signer=signer, 
                            out=output
                        )
                        
                        st.success("¡PDF firmado correctamente con certificado digital!")
                        st.download_button(
                            label="⬇️ Descargar PDF Firmado Digitalmente",
                            data=output.getvalue(),
                            file_name=f"firmado_digital_{pdf_subido.name}",
                            mime="application/pdf"
                        )
                        
                    except Exception as e:
                        st.error(f"ERROR REAL: {str(e)}")
                        st.code(traceback.format_exc())

# ==========================================
# MÓDULO 5: EDITAR Y RELLENAR PDF
# ==========================================
elif opcion == "Editar y Rellenar PDF":
    st.header("📝 Rellenar Formularios PDF (Multicampo)")
    archivo_pdf = st.file_uploader("Sube el formulario PDF a rellenar", type=["pdf"])
    
    if archivo_pdf:
        if st.session_state.pdf_actual_form != archivo_pdf.name:
            st.session_state.textos_agregados = []
            st.session_state.pdf_actual_form = archivo_pdf.name
            st.session_state.texto_x = 50.0
            st.session_state.texto_y = 50.0
            st.session_state.last_click_texto = None

        pdf_bytes = archivo_pdf.getvalue()
        doc_info = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_paginas = len(doc_info)
        
        col_herramientas, col_vista = st.columns([1, 2], gap="large")

        with col_herramientas:
            st.write("### 1. Herramientas de Texto")
            pagina_texto = st.number_input("Página actual:", min_value=1, max_value=num_paginas, value=1)
            texto_insertar = st.text_input("Escribe aquí el texto a insertar:", "Ej: Carlos Gómez")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1: tamano_fuente = st.slider("Tamaño letra", 8, 72, 12)
            with col_f2: color_texto = st.color_picker("Color", "#000000")
            
            rgb_color = tuple(int(color_texto.lstrip('#')[i:i+2], 16)/255 for i in (0, 2, 4))
            page_info = doc_info[pagina_texto - 1]
            max_width = float(page_info.rect.width)
            max_height = float(page_info.rect.height)
            
            st.slider("Ajuste Fino X", 0.0, max_width, value=st.session_state.texto_x, key="slider_texto_x", on_change=update_texto_pos)
            st.slider("Ajuste Fino Y", 0.0, max_height, value=st.session_state.texto_y, key="slider_texto_y", on_change=update_texto_pos)
            
            if st.button("➕ Añadir este texto al documento", type="primary", use_container_width=True):
                if texto_insertar.strip():
                    st.session_state.textos_agregados.append({
                        "pagina": pagina_texto, "texto": texto_insertar,
                        "x": st.session_state.texto_x, "y": st.session_state.texto_y,
                        "size": tamano_fuente, "color": rgb_color
                    })
                    st.session_state.texto_y = min(max_height, st.session_state.texto_y + 25.0)
                    st.rerun()

            st.write("---")
            st.write("### 📋 Textos Agregados")
            if st.session_state.textos_agregados:
                for i, txt in enumerate(st.session_state.textos_agregados):
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"**Pg {txt['pagina']}:** {txt['texto']}")
                    if c2.button("🗑️", key=f"del_{i}"):
                        st.session_state.textos_agregados.pop(i)
                        st.rerun()
                
                st.write("---")
                if st.button("⬇️ Aplicar Todos los Textos y Descargar", type="primary"):
                    with st.spinner("Procesando..."):
                        doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                        for txt in st.session_state.textos_agregados:
                            page_final = doc_final[txt["pagina"] - 1]
                            p_final = fitz.Point(txt["x"], txt["y"])
                            page_final.insert_text(p_final, txt["texto"], fontsize=txt["size"], color=txt["color"])
                        
                        salida_pdf = BytesIO()
                        doc_final.save(salida_pdf)
                        doc_final.close()
                        st.success("¡Formulario completado!")
                        st.download_button(label="Descargar PDF Final", data=salida_pdf.getvalue(), file_name=f"formulario_{archivo_pdf.name}", mime="application/pdf")
            else:
                st.info("Aún no has agregado ningún texto.")
        
        with col_vista:
            st.write("### Vista Previa Interactiva (Clic para posicionar)")
            doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_preview = doc_preview[pagina_texto - 1]
            
            for txt in st.session_state.textos_agregados:
                if txt["pagina"] == pagina_texto:
                    p_hist = fitz.Point(txt["x"], txt["y"])
                    page_preview.insert_text(p_hist, txt["texto"], fontsize=txt["size"], color=txt["color"])
            
            if texto_insertar.strip():
                p_actual = fitz.Point(st.session_state.texto_x, st.session_state.texto_y)
                page_preview.insert_text(p_actual, texto_insertar, fontsize=tamano_fuente, color=rgb_color)
            
            pix = page_preview.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            coords = streamlit_image_coordinates(img_pil, key="click_texto")
            if coords is not None:
                click_id = f"{coords['x']}-{coords['y']}"
                if st.session_state.last_click_texto != click_id:
                    st.session_state.last_click_texto = click_id
                    st.session_state.texto_x = float(coords["x"]) / 1.5
                    st.session_state.texto_y = float(coords["y"]) / 1.5
                    st.rerun()

        doc_info.close()

# ==========================================
# MÓDULO 6: MARCA DE AGUA
# ==========================================
elif opcion == "Marca de Agua":
    st.header("©️ Insertar Marca de Agua")
    archivo_pdf = st.file_uploader("Sube el PDF", type=["pdf"])
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        
        st.write("### Configuración de la Marca")
        tipo = st.radio("Tipo de marca:", ["Texto", "Imagen (Logo)"])
        
        if tipo == "Texto":
            texto_marca = st.text_input("Texto de la marca de agua:", "CONFIDENCIAL")
            col1, col2, col3 = st.columns(3)
            with col1: color_marca = st.color_picker("Color del texto", "#808080")
            with col2: size_marca = st.slider("Tamaño de texto", 20, 150, 60)
            with col3: opacidad = st.slider("Opacidad", 0.1, 1.0, 0.5)
            
            rgb_color = tuple(int(color_marca.lstrip('#')[i:i+2], 16)/255 for i in (0, 2, 4))
            
            if st.button("Aplicar Marca de Agua (Todas las hojas)"):
                with st.spinner("Aplicando..."):
                    doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                    for page in doc_final:
                        rect = page.rect
                        x_center = rect.width / 2 - (len(texto_marca) * size_marca / 4)
                        y_center = rect.height / 2
                        p = fitz.Point(x_center, y_center)
                        page.insert_text(p, texto_marca, fontsize=size_marca, color=rgb_color, fill_opacity=opacidad)
                    
                    salida_pdf = BytesIO()
                    doc_final.save(salida_pdf)
                    doc_final.close()
                    st.success("¡Marca de agua de texto aplicada!")
                    st.download_button("Descargar PDF con Marca", data=salida_pdf.getvalue(), file_name=f"marca_agua_{archivo_pdf.name}", mime="application/pdf")
        
        else:
            img_marca = st.file_uploader("Sube tu logo/imagen", type=["png", "jpg", "jpeg"])
            if img_marca:
                col1, col2 = st.columns(2)
                with col1: escala = st.slider("Tamaño de la imagen (%)", 10, 100, 50)
                with col2: opacidad_img = st.slider("Opacidad", 0.1, 1.0, 0.3)
                
                if st.button("Aplicar Logo como Marca de Agua"):
                    with st.spinner("Procesando..."):
                        img = Image.open(img_marca).convert("RGBA")
                        alpha = img.split()[3]
                        alpha = ImageEnhance.Brightness(alpha).enhance(opacidad_img)
                        img.putalpha(alpha)
                        
                        img_byte_arr = BytesIO()
                        img.save(img_byte_arr, format='PNG')
                        img_bytes_transparente = img_byte_arr.getvalue()
                        
                        doc_final = fitz.open(stream=pdf_bytes, filetype="pdf")
                        for page in doc_final:
                            rect = page.rect
                            img_w = (rect.width * escala) / 100
                            img_h = (img.height * img_w) / img.width
                            
                            x0 = (rect.width - img_w) / 2
                            y0 = (rect.height - img_h) / 2
                            rect_img = fitz.Rect(x0, y0, x0 + img_w, y0 + img_h)
                            
                            page.insert_image(rect_img, stream=img_bytes_transparente)
                            
                        salida_pdf = BytesIO()
                        doc_final.save(salida_pdf)
                        doc_final.close()
                        st.success("¡Logo aplicado como marca de agua!")
                        st.download_button("Descargar PDF con Logo", data=salida_pdf.getvalue(), file_name=f"logo_agua_{archivo_pdf.name}", mime="application/pdf")

# ==========================================
# MÓDULO 7: PROTEGER / DESBLOQUEAR
# ==========================================
elif opcion == "Proteger / Desbloquear PDF":
    st.header("🔐 Proteger o Desbloquear PDF")
    modo_seguridad = st.radio("Elige una acción:", ["Proteger (Añadir Contraseña)", "Desbloquear (Quitar Contraseña)"])
    
    archivo_pdf = st.file_uploader(f"Sube el PDF a {modo_seguridad.split(' ')[0].lower()}", type=["pdf"])
    
    if archivo_pdf:
        pdf_bytes = archivo_pdf.getvalue()
        
        if "Proteger" in modo_seguridad:
            col1, col2 = st.columns(2)
            with col1: user_pass = st.text_input("Contraseña de Apertura (Para leer):", type="password")
            with col2: owner_pass = st.text_input("Contraseña Maestra (Para editar/imprimir):", type="password")
            
            st.write("### Permisos para los Lectores")
            perm_imprimir = st.checkbox("Permitir Imprimir", value=True)
            perm_copiar = st.checkbox("Permitir Copiar Texto y Gráficos", value=True)
            perm_editar = st.checkbox("Permitir Modificar el Documento", value=False)
            
            if st.button("Proteger y Descargar"):
                if not user_pass and not owner_pass:
                    st.error("Debes ingresar al menos una contraseña.")
                else:
                    with st.spinner("Encriptando documento con AES-256..."):
                        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                        perms = fitz.PDF_PERM_ACCESSIBILITY 
                        if perm_imprimir: perms |= fitz.PDF_PERM_PRINT
                        if perm_copiar: perms |= fitz.PDF_PERM_COPY
                        if perm_editar: perms |= fitz.PDF_PERM_MODIFY
                        
                        salida_pdf = BytesIO()
                        doc.save(salida_pdf, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=user_pass, owner_pw=owner_pass if owner_pass else user_pass, permissions=perms)
                        doc.close()
                        st.success("¡Documento protegido exitosamente!")
                        st.download_button("Descargar PDF Seguro", data=salida_pdf.getvalue(), file_name=f"protegido_{archivo_pdf.name}", mime="application/pdf")
        
        else:
            pass_actual = st.text_input("Ingresa la contraseña actual del documento:", type="password")
            if st.button("Desbloquear y Descargar"):
                with st.spinner("Intentando desbloquear..."):
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    if doc.is_encrypted:
                        if doc.authenticate(pass_actual):
                            salida_pdf = BytesIO()
                            doc.save(salida_pdf, encryption=fitz.PDF_ENCRYPT_NONE)
                            doc.close()
                            st.success("¡Documento desbloqueado exitosamente!")
                            st.download_button("Descargar PDF Liberado", data=salida_pdf.getvalue(), file_name=f"desbloqueado_{archivo_pdf.name}", mime="application/pdf")
                        else:
                            st.error("❌ Contraseña incorrecta. El documento sigue bloqueado.")
                            doc.close()
                    else:
                        st.warning("Este documento no está protegido por contraseña.")
                        doc.close()

# ==========================================
# MÓDULO 8: CONVERTIR ARCHIVOS
# ==========================================
elif opcion == "Convertir a Word, Excel, PPT":
    st.header("🔄 Convertir PDF a Otros Formatos")
    archivo_pdf = st.file_uploader("Sube tu archivo PDF", type=["pdf"])
    if archivo_pdf is not None:
        formato = st.radio("Selecciona el formato:", ["Word (.docx) - Texto editable", "Excel (.xlsx) - Extraer Tablas", "PowerPoint (.pptx) - Presentación visual"])
        if st.button(f"Convertir a {formato.split(' ')[0]}"):
            if "Word" in formato:
                with st.spinner("Convirtiendo a Word..."):
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
                        st.download_button(label="Descargar Word", data=docx_bytes, file_name=f"{archivo_pdf.name.replace('.pdf', '')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    finally:
                        if os.path.exists(ruta_pdf): os.remove(ruta_pdf)
                        if os.path.exists(ruta_docx): os.remove(ruta_docx)
            elif "Excel" in formato:
                with st.spinner("Escaneando tablas..."):
                    try:
                        todas_las_tablas = []
                        with pdfplumber.open(BytesIO(archivo_pdf.getvalue())) as pdf:
                            for page in pdf.pages:
                                tables = page.extract_tables()
                                for table in tables:
                                    if table: 
                                        df = pd.DataFrame(table[1:], columns=table[0]) if len(table)>1 else pd.DataFrame(table)
                                        todas_las_tablas.append(df)
                        if not todas_las_tablas: st.warning("No se detectaron tablas.")
                        else:
                            salida_excel = BytesIO()
                            with pd.ExcelWriter(salida_excel, engine='openpyxl') as writer:
                                for i, df in enumerate(todas_las_tablas): df.to_excel(writer, sheet_name=f"Tabla_{i+1}", index=False)
                            st.success(f"¡{len(todas_las_tablas)} tablas extraídas!")
                            st.download_button(label="Descargar Excel", data=salida_excel.getvalue(), file_name=f"{archivo_pdf.name.replace('.pdf', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e: st.error(f"Error al extraer: {e}")
            elif "PowerPoint" in formato:
                with st.spinner("Creando diapositivas..."):
                    try:
                        prs = Presentation()
                        blank_slide_layout = prs.slide_layouts[6] 
                        doc = fitz.open(stream=archivo_pdf.getvalue(), filetype="pdf")
                        for page_num in range(len(doc)):
                            page = doc[page_num]
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) 
                            slide = prs.slides.add_slide(blank_slide_layout)
                            slide.shapes.add_picture(BytesIO(pix.tobytes("png")), 0, 0, width=prs.slide_width, height=prs.slide_height)
                        doc.close()
                        salida_ppt = BytesIO()
                        prs.save(salida_ppt)
                        st.success("¡Conversión exitosa!")
                        st.download_button(label="Descargar PowerPoint", data=salida_ppt.getvalue(), file_name=f"{archivo_pdf.name.replace('.pdf', '')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                    except Exception as e: st.error(f"Error creando PowerPoint: {e}")
