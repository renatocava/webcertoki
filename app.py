import streamlit as st
import pandas as pd
import os
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
from tempfile import NamedTemporaryFile

if 'df_procesado' not in st.session_state:
    st.session_state.df_procesado = None
if 'grupos' not in st.session_state:
    st.session_state.grupos = None
if 'plantillas' not in st.session_state:
    st.session_state.plantillas = None
if 'certificados_generados' not in st.session_state:
    st.session_state.certificados_generados = False
if 'zip_buffer' not in st.session_state:
    st.session_state.zip_buffer = None

# Registrar fuente personalizada
def register_custom_font():
    """Registra la fuente Trebuchet MS si está disponible"""
    font_path = os.path.join("fonts", "trebuchet.ttf")
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont('Trebuchet', font_path))
            return True
        except Exception as e:
            st.warning(f"No se pudo cargar la fuente Trebuchet MS: {e}")
            return False
    else:
        st.info("Fuente Trebuchet MS no encontrada. Usando fuente por defecto.")
        return False

TREBUCHET_AVAILABLE = register_custom_font()
styles_config = None

# Diccionario de meses
def mes_en_espanol(fecha):
    meses = {
        'January': 'enero',
        'February': 'febrero',
        'March': 'marzo',
        'April': 'abril',
        'May': 'mayo',
        'June': 'junio',
        'July': 'julio',
        'August': 'agosto',
        'September': 'septiembre',
        'October': 'octubre',
        'November': 'noviembre',
        'December': 'diciembre'
    }

    mes_ingles = fecha.strftime('%B')
    mes_espanol = meses.get(mes_ingles, mes_ingles)
    return fecha.strftime(f"%d de {mes_espanol} del %Y")

# Función para agregar marca de agua (PDF)
def agregar_marca_agua(pdf_bytes, watermark_path):
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_bytes)
        watermark_reader = PyPDF2.PdfReader(watermark_path)
        
        watermark_page = watermark_reader.pages[0]
        
        pdf_writer = PyPDF2.PdfWriter()
        
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            
            # Determinar orientación de la página
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            is_landscape = page_width > page_height
            
            # Crear una copia de la marca de agua para no modificar la original
            if is_landscape:
                landscape_watermark_path = os.path.join("watermarks", "marca_agua_landscape.pdf")
                if os.path.exists(landscape_watermark_path):
                    landscape_watermark_reader = PyPDF2.PdfReader(landscape_watermark_path)
                    watermark = landscape_watermark_reader.pages[0]
                else:
                    watermark = watermark_page
            else:
                watermark = watermark_page
            
            page.merge_page(watermark)
            pdf_writer.add_page(page)
        
        result_pdf = BytesIO()
        pdf_writer.write(result_pdf)
        result_pdf.seek(0)
        
        return result_pdf
    except Exception as e:
        st.error(f"Error al aplicar marca de agua: {e}")
        return pdf_bytes

# Función para cargar plantillas
def cargar_plantillas():
    """Carga las plantillas de fondo desde la carpeta plantillas"""
    plantillas = {}
    plantillas_path = "plantillas"

    if not os.path.exists(plantillas_path):
        st.error(f"❌ La carpeta '{plantillas_path}' no existe. Créala y agrega las imágenes de fondo.")
        return None

    archivos_plantilla = {
        'PROGRESIVO_1P_5S.jpg': 'fondo_1',
        'PARTICIPACION_1P_5S.jpg': 'fondo_2',
        'APROBADO_1P_3P.jpg': 'fondo_3',
        'APROBADO_4P_5S.jpg': 'fondo_4'
    }

    for archivo, clave in archivos_plantilla.items():
        ruta_completa = os.path.join(plantillas_path, archivo)
        if os.path.exists(ruta_completa):
            with open(ruta_completa, 'rb') as f:
                plantillas[clave] = f.read()
        else:
            st.warning(f"⚠️ No se encontró {archivo} en la carpeta plantillas")

    if len(plantillas) == 4:
        return plantillas
    else:
        st.error(f"❌ Se necesitan 4 plantillas, solo se encontraron {len(plantillas)}")
        return None

# Función para clasificar estudiantes por criterios
def clasificar_estudiantes_por_nota(df, nombre_archivo):
    grupos = {
        'grupo_1': pd.DataFrame(),  # Progresivo
        'grupo_2': pd.DataFrame(),  # Nota < 13 / Participación
        'grupo_3': pd.DataFrame(),  # Nota ≥ 13 y Grado = v1
        'grupo_4': pd.DataFrame()  # Nota ≥ 13 y Grado = v2
    }

    if 'nota final' not in df.columns:
        st.error("❌ No se encontró la columna 'NOTA FINAL' en el DataFrame")
        return None

    if 'grado' not in df.columns:
        st.error("❌ No se encontró la columna 'GRADO' en el DataFrame")
        return None

    # Verificar si el archivo empieza con "P"
    archivo_empieza_con_p = nombre_archivo.upper().startswith('P')

    if archivo_empieza_con_p:
        # Si el archivo empieza con "P", todos los estudiantes van al grupo 1 (Progresivo)
        grupos['grupo_1'] = df.copy()
        st.info(f"📋 **Archivo detectado con prefijo 'P'**: Todos los certificados usarán el formato Progresivo")

    else:
        df['nota_final_num'] = pd.to_numeric(df['nota final'], errors='coerce')

        # Grupo 2: Nota < 13 - Participación
        grupos['grupo_2'] = df[df['nota_final_num'] < 13].copy()

        # Grupos 3 y 4: Nota ≥ 13
        df_nota_alta = df[df['nota_final_num'] >= 13].copy()

        grupos['grupo_3'] = df_nota_alta[df_nota_alta['grado'].str.lower().str.strip().isin(['1p', '2p', '3p'])].copy()
        grupos['grupo_4'] = df_nota_alta[
            df_nota_alta['grado'].str.lower().str.strip().isin(['4p', '5p', '1s', '2s', '3s', '4s', '5s'])].copy()

    return grupos

# Función para procesar el archivo Excel Base
def procesar_excel_inicial(uploaded_file):
    """
    Procesa el archivo Excel eliminando las primeras 9 filas y columnas J-N y desde la T
    """
    try:
        df_original = pd.read_excel(uploaded_file)

        # Eliminar las primeras 11 filas (índices 0-10, quedando la fila 12 como cabecera)
        df_procesado = df_original.iloc[10:].copy()

        # Resetear el índice para que la nueva primera fila sea el índice 0
        df_procesado = df_procesado.reset_index(drop=True)

        # Usar la primera fila como nombres de columnas
        df_procesado.columns = df_procesado.iloc[0]
        df_procesado = df_procesado.drop(df_procesado.index[0]).reset_index(drop=True)

        # Conservar solo las columnas A-I y O-S
        if len(df_procesado.columns) > 19:
            columnas_a_conservar = list(range(9)) + list(range(14, 20))
            df_procesado = df_procesado.iloc[:, columnas_a_conservar]

        df_procesado.columns = df_procesado.columns.str.lower()

        # Reemplazar 'NP' por 0 en la columna 'nota final'
        if 'nota final' in df_procesado.columns:
            df_procesado['nota final'] = df_procesado['nota final'].apply(
                lambda x: 0 if isinstance(x, str) and x.strip().upper() == 'NP' else x
            )

        df_procesado['nombre_certificado'] = df_procesado['nombre'].fillna('').str.strip() + ' ' + df_procesado[
            'paterno'].fillna('').str.strip() + ' ' + df_procesado['materno'].fillna('').str.strip()

        columnas = df_procesado.columns.tolist()
        columnas.remove('nombre_certificado')
        posicion_nro = columnas.index('nro')
        columnas.insert(posicion_nro + 1, 'nombre_certificado')
        df_procesado = df_procesado[columnas]

        return df_procesado, True, "Archivo procesado correctamente"

    except Exception as e:
        return None, False, f"Error al procesar el archivo: {str(e)}"

# Acomodar el texto en múltiples líneas para que se ajuste al ancho máximo
def wrap_text_to_width(canvas, text, font_name, font_size, max_width_mm):
    max_width_points = max_width_mm * 2.83465
    words = text.split()
    lines = []
    current_line = []

    for word in words:
        test_line = current_line + [word]
        test_text = ' '.join(test_line)
        text_width = canvas.stringWidth(test_text, font_name, font_size)

        if text_width <= max_width_points:
            current_line = test_line
        else:
            if current_line:
                lines.append(' '.join(current_line))
                current_line = [word]
            else:
                lines.append(word)

    if current_line:
        lines.append(' '.join(current_line))

    return lines

# Dibuja texto multilínea usando la configuración de estilos específica
def draw_multiline_text(canvas, text, style_key, page_width, styles_config, max_width_mm=None):
    style = styles_config[style_key]
    font_name = style['font_family'] if TREBUCHET_AVAILABLE else 'Helvetica'

    is_bold = style.get('bold', False)

    if is_bold:
        try:
            bold_font_name = f"{font_name}-Bold"
            canvas.setFont(bold_font_name, style['font_size'])
            font_name = bold_font_name
        except Exception as e:
            try:
                canvas.setFont(font_name, style['font_size'])
            except Exception as e:
                canvas.setFont('Helvetica-Bold' if is_bold else 'Helvetica', style['font_size'])
                font_name = 'Helvetica-Bold' if is_bold else 'Helvetica'
    else:
        try:
            canvas.setFont(font_name, style['font_size'])
        except Exception as e:
            canvas.setFont('Helvetica', style['font_size'])
            font_name = 'Helvetica'

    canvas.setFillColor(HexColor(style['color']))
    x_points = style['x'] * 2.83465
    y_points = style['y'] * 2.83465

    if max_width_mm is None:
        if style['x'] == 148 or style['x'] == 105:
            text_width = canvas.stringWidth(text, font_name, style['font_size'])
            x_points = (page_width - text_width) / 2
        canvas.drawString(x_points, y_points, text)
        return style['font_size']

    lines = wrap_text_to_width(canvas, text, font_name, style['font_size'], max_width_mm)
    line_height = style['font_size'] * 1.2
    start_y = y_points

    for i, line in enumerate(lines):
        line_y = start_y - (i * line_height)
        if style['x'] == 148 or style['x'] == 105:
            text_width = canvas.stringWidth(line, font_name, style['font_size'])
            line_x = (page_width - text_width) / 2
        else:
            line_x = x_points
        canvas.drawString(line_x, line_y, line)

    return line_height * len(lines)

# Genera certificados para un grupo específico con su plantilla y estilos correspondientes
def generar_certificados_grupo(grupo_df, plantilla_bytes, plantilla_key, nombre_grupo, zip_file, progress_bar,
    estudiantes_base, total_estudiantes, styles_config_by_template):
    certificados_generados = 0

    # Aplicar marca de agua si la segunda letra es 'I' y si esta aprobado
    nombre_archivo = st.session_state.get('nombre_archivo', '')
    aplicar_marca_agua = len(nombre_archivo) >= 2 and nombre_archivo[1].upper() == 'I' and plantilla_key != 'fondo_2'
    
    # Ruta a la marca de agua
    watermark_path = os.path.join("watermarks", "marca_agua.pdf")
    if aplicar_marca_agua and not os.path.exists(watermark_path):
        st.warning(f"⚠️ No se encontró el archivo de marca de agua en {watermark_path}. Se generarán PDFs sin marca de agua.")
        aplicar_marca_agua = False

    # Obtener la configuración de estilos para esta plantilla
    styles_config = styles_config_by_template[plantilla_key]

    # Determinar orientación de página según la plantilla
    if styles_config.get('orientation') == 'portrait':
        page_size = A4
        page_width, page_height = A4
    else:
        page_size = landscape(A4)
        page_width, page_height = landscape(A4)

    for i, row in grupo_df.iterrows():
        try:
            nombre = str(row["nombre_certificado"]).strip().upper()
            curso = str(row["curso"]).strip().upper()
            fecha = mes_en_espanol(datetime.today())
            numero = (
                str(row["numeración"]).strip()
                if "numeración" in row and pd.notnull(row["numeración"])
                else f"GEN-{i + 1:03}"
            )

            # Extraer valores para la variable de horas, sólo si es para "Progresivos" (fondo_1)
            horas = "horas_progresivo"
            horas_progresivo = ""
            if plantilla_key == 'fondo_1' and horas in row and pd.notnull(row[horas]):
                horas_progresivo = str(row[horas])

            # Crear archivo temporal con la plantilla
            with NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                tmp_img.write(plantilla_bytes)
                tmp_img.flush()
                tmp_img_path = tmp_img.name

            # Crear PDF con orientación específica
            pdf_buffer = BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=page_size)

            # Insertar imagen de fondo
            c.drawImage(tmp_img_path, 0, 0, width=page_width, height=page_height)

            # Dibujar texto usando los estilos específicos de la plantilla
            draw_multiline_text(c, nombre, 'nombre', page_width, styles_config, styles_config['nombre']['max_width'])
            draw_multiline_text(c, curso, 'curso', page_width, styles_config, styles_config['curso']['max_width'])
            draw_multiline_text(c, f"Lima, {fecha}", 'fecha', page_width, styles_config)

            # Se considera la variable horas si es para el fondo_1
            if plantilla_key == 'fondo_1' and horas_progresivo:
                draw_multiline_text(c, horas_progresivo, 'horas', page_width, styles_config)
            
            if plantilla_key != 'fondo_2':
                draw_multiline_text(c, f"Certificado Nº {numero}", 'numero', page_width, styles_config)

            c.save()
            pdf_bytes = pdf_buffer.getvalue()

            # Aplicar marca de agua si es necesario
            if aplicar_marca_agua:
                pdf_buffer = agregar_marca_agua(BytesIO(pdf_bytes), watermark_path)
                pdf_bytes = pdf_buffer.getvalue()

            # Añadir al ZIP
            # Si es el grupo 2 (Constancias), guardar en la subcarpeta "Constancias"
            # pdf_name = f"{nombre.replace(' ', '_')}.pdf"
            # zip_file.writestr(pdf_name, pdf_bytes)
            if plantilla_key == 'fondo_2':
                pdf_name = f"Constancias/{nombre.strip().replace(' ', '_') + '_' + curso[0:11].replace(' ', '_')}.pdf"
            else:
                pdf_name = f"{nombre.strip().replace(' ', '_') + '_' + curso[0:11].replace(' ', '_')}.pdf"

            zip_file.writestr(pdf_name, pdf_bytes)

            certificados_generados += 1

            # Actualizar progreso
            progreso_actual = (estudiantes_base + certificados_generados) / total_estudiantes
            progress_bar.progress(min(progreso_actual, 1.0))

            # Limpiar archivo temporal
            try:
                if os.path.exists(tmp_img_path):
                    os.unlink(tmp_img_path)
            except:
                pass

        except Exception as e:
            st.error(f"Error generando certificado para {nombre}: {e}")

    return certificados_generados

# Función para generar todos los certificados
def generar_todos_certificados():
    if st.session_state.grupos and st.session_state.plantillas:
        st.info("Generando certificados por grupos...")

        total_estudiantes = sum(len(grupo) for grupo in st.session_state.grupos.values() if not grupo.empty)
        progress_bar = st.progress(0)
        estudiantes_procesados = 0

        zip_buffer = BytesIO()

        with ZipFile(zip_buffer, "a") as zip_file:
            # Crear directorio para constancias
            zip_file.writestr("Constancias/", "")

            # Configuración de estilos
            styles_config_by_template = {
            "fondo_1": {
                'curso': {
                    'font_family': 'Trebuchet',
                    'font_size': 32,
                    'color': '#000000', #11959f
                    'x': 52,
                    'y': 129,
                    'max_width': 220,
                    'bold': True
                },
                'nombre': {
                    'font_family': 'Trebuchet',
                    'font_size': 25,
                    'color': '#000000', #004064
                    'x': 52,
                    'y': 85,
                    'max_width': 210
                },
                'fecha': {
                    'font_family': 'Trebuchet',
                    'font_size': 18,
                    'color': '#004064',
                    'x': 52,
                    'y': 36,
                    'max_width': None,
                    'bold': True
                },
                'numero': {
                    'font_family': 'Trebuchet',
                    'font_size': 15.5,
                    'color': '#004064',
                    'x': 52,
                    'y': 27,
                    'max_width': None
                },
                'horas': {
                    'font_family': 'Trebuchet',
                    'font_size': 15.5,
                    'color': '#004064',
                    'x': 132.5,
                    'y': 65.2,
                    'max_width': None
                },
                'orientation': 'landscape'  # Orientación horizontal
            },
            "fondo_2": {  # Vertical
                'curso': {
                    'font_family': 'Trebuchet',
                    'font_size': 30.5,
                    'color': '#000000', #11959f
                    'x': 105,
                    'y': 185,
                    'max_width': 160,
                    'bold': True
                },
                'nombre': {
                    'font_family': 'Trebuchet',
                    'font_size': 29,
                    'color': '#000000', #004064
                    'x': 105,
                    'y': 131,
                    'max_width': 160,
                    'bold': True
                },
                'fecha': {
                    'font_family': 'Trebuchet',
                    'font_size': 18,
                    'color': '#004064',
                    'x': 105,
                    'y': 78,
                    'max_width': None
                },
                # No aparece en el certificado, sólo está para evitar errores en f()
                'numero': {
                    'font_family': 'Trebuchet',
                    'font_size': 1,
                    'color': '#ffffff',
                    'x': 0,
                    'y': 0,
                    'max_width': None
                },
                'orientation': 'portrait'  # Orientación vertical
            },
            "fondo_3": {
                'curso': {
                    'font_family': 'Trebuchet',
                    'font_size': 30.5,
                    'color': '#000000', #11959f
                    'x': 148,
                    'y': 117,
                    'max_width': 245,
                    'bold': True
                },
                'nombre': {
                    'font_family': 'Trebuchet',
                    'font_size': 29,
                    'color': '#000000', #004064
                    'x': 148,
                    'y': 75,
                    'max_width': 245,
                    'bold': True
                },
                'fecha': {
                    'font_family': 'Trebuchet',
                    'font_size': 18,
                    'color': '#004064',
                    'x': 20,
                    'y': 41,
                    'max_width': None,
                    'bold': True
                },
                'numero': {
                    'font_family': 'Trebuchet',
                    'font_size': 15.5,
                    'color': '#004064',
                    'x': 20,
                    'y': 32,
                    'max_width': None
                },
                'orientation': 'landscape'
            },
            "fondo_4": {
                'curso': {
                    'font_family': 'Trebuchet',
                    'font_size': 30.5,
                    'color': '#000000', #11959f
                    'x': 148,
                    'y': 117,
                    'max_width': 245,
                    'bold': True
                },
                'nombre': {
                    'font_family': 'Trebuchet',
                    'font_size': 29,
                    'color': '#000000', #004064
                    'x': 148,
                    'y': 75,
                    'max_width': 245,
                    'bold': True
                },
                'fecha': {
                    'font_family': 'Trebuchet',
                    'font_size': 18,
                    'color': '#004064',
                    'x': 20,
                    'y': 41,
                    'max_width': None,
                    'bold': True
                },
                'numero': {
                    'font_family': 'Trebuchet',
                    'font_size': 15.5,
                    'color': '#004064',
                    'x': 20,
                    'y': 32,
                    'max_width': None
                },
                'orientation': 'landscape'
            }
        }

            # Mapeo de grupos a plantillas
            mapeo_plantillas = {
                'grupo_1': 'fondo_1',  # Progresiva
                'grupo_2': 'fondo_2',  # Participación Nota < 13
                'grupo_3': 'fondo_3',  # Base - Nota ≥ 13 y Grado = 1P-3P
                'grupo_4': 'fondo_4'  # Base - Nota ≥ 13 y Grado = 4P-5S
            }

            for grupo_nombre, grupo_df in st.session_state.grupos.items():
                if not grupo_df.empty:
                    plantilla_key = mapeo_plantillas[grupo_nombre]
                    plantilla_bytes = st.session_state.plantillas[plantilla_key]

                    st.write(f"Procesando {grupo_nombre} ({len(grupo_df)} estudiantes) con plantilla {plantilla_key}...")

                    # Generar certificados pasando la configuración de estilos
                    certificados_gen = generar_certificados_grupo(
                        grupo_df,
                        plantilla_bytes,
                        plantilla_key,
                        grupo_nombre,
                        zip_file,
                        progress_bar,
                        estudiantes_procesados,
                        total_estudiantes,
                        styles_config_by_template
                    )

                    estudiantes_procesados += len(grupo_df)

                    st.success(f"✅ {grupo_nombre}: {certificados_gen} certificados generados con estilo {plantilla_key}")

        zip_buffer.seek(0)
        st.success("🎉 Todos los certificados han sido generados correctamente y están listos para su descarga.")
        
        st.session_state.zip_buffer = zip_buffer
        st.session_state.certificados_generados = True
        
        return True
    return False


# Configuración de la web
st.set_page_config(page_title="Generador de Certificados", layout="centered")
st.title("🎓 Generador de Certificados PDF con Plantillas Automáticas")

# Preprocesamiento del Excel
st.header("📤 Subir y procesar archivo Excel")
uploaded_file = st.file_uploader("Selecciona un archivo Excel", type=["xlsx"])

if uploaded_file:
    st.subheader("📊 Vista previa del archivo original")
    df_original = pd.read_excel(uploaded_file)
    st.write(f"**Dimensiones originales:** {df_original.shape[0]} filas x {df_original.shape[1]} columnas")
    st.write(f"**Nombre del archivo:** {uploaded_file.name}")
    st.dataframe(df_original.head(15))

    # Procesar automáticamente el archivo
    with st.spinner("Procesando archivo..."):
        df_procesado, exito, mensaje = procesar_excel_inicial(uploaded_file)
        
        if exito:
            st.session_state.df_procesado = df_procesado
            st.session_state.nombre_archivo = uploaded_file.name
            
            # Resetear estados cuando se procesa un nuevo archivo
            st.session_state.grupos = None
            st.session_state.plantillas = None
            st.session_state.certificados_generados = False
            st.session_state.zip_buffer = None
            
            st.success(mensaje)
            st.subheader("✅ Archivo procesado - Vista previa de datos limpios")
            st.write(f"**Dimensiones procesadas:** {df_procesado.shape[0]} filas x {df_procesado.shape[1]} columnas")
            st.dataframe(df_procesado)
            
            # Cargar plantillas automáticamente
            st.session_state.plantillas = cargar_plantillas()
            
            # Clasificar estudiantes automáticamente
            nombre_archivo = st.session_state.nombre_archivo
            st.session_state.grupos = clasificar_estudiantes_por_nota(st.session_state.df_procesado, nombre_archivo)
            
            # Generar certificados automáticamente
            with st.spinner("Generando certificados..."):
                generar_todos_certificados()
        else:
            st.error(mensaje)

# Mostrar botón de descarga si los certificados fueron generados
if st.session_state.certificados_generados and st.session_state.zip_buffer:
    nombre_archivo = st.session_state.get('nombre_archivo', '')
    nombre_base = os.path.splitext(nombre_archivo)[0] if nombre_archivo else "CERTIFICADOS"
    
    if len(nombre_base) >= 2 and nombre_base[1].upper() == 'P':
        zip_filename = f"{nombre_base}_PRELIMINAR.zip"
    else:
        zip_filename = f"{nombre_base}.zip"
    
    st.download_button(
        label="📥 Descargar todos los certificados (ZIP)",
        data=st.session_state.zip_buffer,
        file_name=zip_filename,
        mime="application/zip"
    )
elif not uploaded_file:
    st.info("👆 Sube un archivo Excel para generar los certificados automáticamente.")