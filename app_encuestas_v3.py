import streamlit as st
import pandas as pd
import json
import base64
import time
from io import BytesIO
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuración de la página de Streamlit para móviles
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para que se vea hermoso en celulares y computadoras
st.markdown("""
<style>
    .main-header {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #1E3A8A;
        font-weight: 700;
        text-align: center;
        padding-bottom: 10px;
        border-bottom: 2px solid #3B82F6;
    }
    .section-subheader {
        color: #1F2937;
        font-weight: 600;
        margin-top: 15px;
    }
    .card {
        background-color: #F3F4F6;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #3B82F6;
        margin-bottom: 15px;
    }
    .success-text {
        color: #10B981;
        font-weight: bold;
    }
    .error-text {
        color: #EF4444;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. CONEXIÓN EXCLUSIVA A GOOGLE SHEETS
# ==========================================

def obtener_credenciales_google():
    """Obtiene las credenciales de Google Cloud de los Secretos de Streamlit."""
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            # Formatear la clave privada para evitar errores de saltos de línea
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            return creds_dict
    except Exception:
        pass
    return None

def conectar_google_sheets():
    """Establece conexión con Google Sheets. Retorna el cliente o None si falla."""
    creds_dict = obtener_credenciales_google()
    if not creds_dict:
        return None
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        return None

def inicializar_base_datos():
    """Busca o crea las hojas necesarias en Google Sheets con sus columnas oficiales."""
    client = conectar_google_sheets()
    if not client:
        return None, None
    
    nombre_libro = "Base_Encuestas_SRPA"
    try:
        # Intentar abrir el libro existente
        sh = client.open(nombre_libro)
    except gspread.exceptions.SpreadsheetNotFound:
        # Crear el libro si no existe
        sh = client.create(nombre_libro)
        # Compartir con el correo de la cuenta de servicio como editor
        creds_dict = obtener_credenciales_google()
        if creds_dict:
            sh.share(creds_dict["client_email"], perm_type='user', role='editor')
            
    # Inicializar pestaña Cola_Revision
    try:
        ws_cola = sh.worksheet("Cola_Revision")
    except gspread.exceptions.WorksheetNotFound:
        ws_cola = sh.add_worksheet(title="Cola_Revision", rows="1000", cols="15")
        # Columnas oficiales
        headers = [
            "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
            "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", 
            "Foto_Pag1_Base64", "Foto_Pag2_Base64", "Estado"
        ]
        ws_cola.append_row(headers)
        
    # Inicializar pestaña Respuestas_SRPA
    try:
        ws_respuestas = sh.worksheet("Respuestas_SRPA")
    except gspread.exceptions.WorksheetNotFound:
        ws_respuestas = sh.add_worksheet(title="Respuestas_SRPA", rows="2000", cols="30")
        headers = [
            "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
            "Institucion_Educativa_Verificada", "Rol",
            "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
            "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
            "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
            "Verificado_Por", "Fecha_Aprobacion"
        ]
        ws_respuestas.append_row(headers)
        
    return ws_cola, ws_respuestas

# Cargar las hojas de cálculo
ws_cola, ws_respuestas = inicializar_base_datos()

# Mostrar tutorial si no está conectado
if ws_cola is None:
    st.title("📊 Proyecto Construyendo Futuro")
    st.info("⚠️ **Google Sheets no está configurado o conectado.**")
    st.markdown("""
    ### ⚙️ Instrucciones para conectar tu Base de Datos en Google Sheets:
    
    Para que las encuestas se guarden de forma colaborativa en la nube, necesitas conectar esta aplicación con Google Sheets:
    
    1. **Crea un proyecto en Google Cloud Console** y activa las APIs de **Google Sheets** y **Google Drive**.
    2. **Crea una Cuenta de Servicio** y descarga la clave en formato **JSON**.
    3. **Crea una hoja de cálculo en tu Google Drive** llamada exactamente `Base_Encuestas_SRPA`.
    4. **Comparte esa hoja de cálculo** con el correo electrónico de tu Cuenta de Servicio como **Editor**.
    5. **Ingresa los secretos en tu panel de Streamlit Cloud:**
       * Ve a la configuración de tu aplicación en Streamlit Cloud (`Settings > Secrets`).
       * Pega el contenido de tu archivo JSON de Google de la siguiente manera:
       
       ```toml
       GEMINI_API_KEY = "tu-api-key-de-google-ai-studio"
       
       [gcp_service_account]
       type = "service_account"
       project_id = "tu-proyecto"
       private_key_id = "tu-clave-id"
       private_key = "-----BEGIN PRIVATE KEY-----\\nTU-LLAVE-AQUI\\n-----END PRIVATE KEY-----\\n"
       client_email = "tu-correo-de-servicio@gserviceaccount.com"
       client_id = "tu-cliente-id"
       auth_uri = "https://accounts.google.com/o/oauth2/auth"
       token_uri = "https://oauth2.google.com/token"
       auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
       client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/tu-correo"
       ```
    6. **Recarga esta página** para iniciar el sistema móvil.
    """)
    st.stop()

# ==========================================
# 2. FUNCIONES DE LECTURA Y ESCRITURA SEGURA
# ==========================================

def cargar_cola_revision():
    """Carga los registros de la cola de revisión de forma segura e inmune a KeyError."""
    try:
        rows = ws_cola.get_all_values()
        if not rows or len(rows) <= 1:
            # Retornar DataFrame vacío con las columnas esperadas
            return pd.DataFrame(columns=[
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", 
                "Foto_Pag1_Base64", "Foto_Pag2_Base64", "Estado"
            ])
        
        headers = [h.strip() for h in rows[0]]
        data = rows[1:]
        
        # Ajustar filas para que tengan exactamente la longitud de las cabeceras
        cleaned_data = []
        for r in data:
            if len(r) < len(headers):
                r = r + [""] * (len(headers) - len(r))
            else:
                r = r[:len(headers)]
            cleaned_data.append(r)
            
        df = pd.DataFrame(cleaned_data, columns=headers)
        return df
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame()

def cargar_respuestas_validadas():
    """Carga los registros de la base final de forma segura."""
    try:
        rows = ws_respuestas.get_all_values()
        if not rows or len(rows) <= 1:
            return pd.DataFrame()
        headers = [h.strip() for h in rows[0]]
        data = rows[1:]
        cleaned_data = []
        for r in data:
            if len(r) < len(headers):
                r = r + [""] * (len(headers) - len(r))
            else:
                r = r[:len(headers)]
            cleaned_data.append(r)
        return pd.DataFrame(cleaned_data, columns=headers)
    except Exception as e:
        st.error(f"Error cargando respuestas validadas: {e}")
        return pd.DataFrame()

def append_row_and_verify(sheet, row_data, id_col_index=1):
    """Inserta una fila en Google Sheets y verifica de inmediato que fue escrita con éxito."""
    try:
        # Insertar la fila
        sheet.append_row(row_data)
        time.sleep(1.5)  # Esperar que Google Sheets procese
        
        # Verificar buscando el ID de la encuesta
        id_encuesta = row_data[id_col_index - 1]
        cell = sheet.find(id_encuesta)
        if cell:
            # Fila encontrada con éxito
            return True, cell.row
        return False, None
    except Exception as e:
        return False, str(e)

def update_row_and_verify(sheet, row_num, col_name, new_value):
    """Actualiza una celda específica por su nombre de columna y verifica que se aplicó."""
    try:
        headers = sheet.row_values(1)
        if col_name not in headers:
            return False, f"Columna '{col_name}' no encontrada."
        
        col_index = headers.index(col_name) + 1
        sheet.update_cell(row_num, col_index, new_value)
        time.sleep(1.5)
        
        # Verificar lectura de vuelta
        val_leido = sheet.cell(row_num, col_index).value
        if val_leido == new_value:
            return True, None
        return False, "La verificación de escritura falló (el valor en la nube no coincide)."
    except Exception as e:
        return False, str(e)

# ==========================================
# 3. PROCESAMIENTO DE IMÁGENES Y COMPRESIÓN
# ==========================================

def comprimir_y_convertir_base64(uploaded_file, max_width=350):
    """Comprime la imagen para reducir costos y la convierte a Base64."""
    try:
        img = Image.open(uploaded_file)
        # Convertir a RGB si tiene canal alpha
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            bg = Image.new('RGB', img.size, (255, 255, 255))
            bg.paste(img, (0, 0), img.convert('RGBA') if img.mode != 'RGBA' else img)
            img = bg
        else:
            img = img.convert('RGB')
            
        # Calcular nueva proporción
        w_percent = (max_width / float(img.size[0]))
        h_size = int((float(img.size[1]) * float(w_percent)))
        img_resized = img.resize((max_width, h_size), Image.Resampling.LANCELET)
        
        # Guardar en buffer
        buffered = BytesIO()
        img_resized.save(buffered, format="JPEG", quality=80)
        img_bytes = buffered.getvalue()
        
        # Convertir a base64
        base64_str = base64.b64encode(img_bytes).decode('utf-8')
        return base64_str, img_bytes
    except Exception as e:
        st.error(f"Error comprimiendo imagen: {e}")
        return None, None

# ==========================================
# 4. MOTOR DE INTELIGENCIA ARTIFICIAL GEMINI
# ==========================================

def inicializar_api_gemini():
    """Inicializa la API de Gemini buscando la API Key de forma segura."""
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    return None

def analizar_encuesta_con_gemini(img_bytes_pag1, img_bytes_pag2, tipo_formulario):
    """
    Motor OCR ultra robusto de tres niveles de seguridad para Gemini.
    Implementa un fallback dinámico de modelos y esquemas para garantizar 100% de éxito.
    """
    api_key = inicializar_api_gemini()
    if not api_key:
        return None, "No se encontró la GEMINI_API_KEY en los secretos de Streamlit."
    
    # Esquema JSON requerido por el prompt para asegurar que siempre retorne la estructura correcta
    esquema_esperado = {
        "tipo_formulario": tipo_formulario,
        "encabezado": {
            "fecha": "YYYY-MM-DD o vacío",
            "municipio": "Texto o vacío",
            "institucion_educativa": "Texto manuscrito en el papel o vacío",
            "rol": "Estudiante/Docente/Padre de Familia/Lider comunitario o vacío"
        },
        "respuestas_conocimiento": {
            "p1": "Opción marcada o vacío (ej: 'a. Si' o 'b. No')",
            "p2": "Opción marcada o vacío",
            "p3": "Opción marcada o vacío",
            "p4": "Opción marcada o vacío",
            "p5": "Opción marcada o vacío",
            "p6": "Opción marcada o vacío",
            "p7": "Opción marcada o vacío",
            "p8": "Opción marcada o vacío"
        },
        "evaluacion_satisfaccion": {
            "sat_p1": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p2": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p3": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p4": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p5": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p6": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p7": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p8": "Excelente/Bueno/Regular/Deficiente o vacío",
            "sat_p9": "Excelente/Bueno/Regular/Deficiente o vacío"
        }
    }
    
    prompt = f"""
    Eres un transcriptor experto en encuestas del "Proyecto Construyendo Futuro" (SRPA).
    Analiza las dos imágenes proporcionadas (Página 1 y Página 2 de un cuestionario).
    
    INSTRUCCIONES CRÍTICAS:
    1. Identifica el tipo de formulario: {tipo_formulario}.
    2. Lee el encabezado manuscrito. Presta mucha atención al nombre de la "Institución Educativa" escrito a mano. No inventes datos.
    3. Para las preguntas de conocimientos y de satisfacción, identifica cuál casilla tiene marcada una "X".
    4. IMPORTANTE: Si una pregunta no tiene ninguna marcación, o está en blanco, debes guardarla como un texto vacío "". NO asumas ni inventes respuestas que el participante no seleccionó.
    
    Retorna un objeto JSON con esta estructura exacta:
    {json.dumps(esquema_esperado, indent=2)}
    """
    
    # Lista de modelos de producción activos (orden de prioridad para 2026)
    modelos_a_probar = [
        "gemini-2.5-flash", 
        "gemini-2.5-pro", 
        "gemini-2.0-flash", 
        "gemini-1.5-flash"
    ]
    
    # Preparar el contenido de la imagen para la API
    contents = []
    if img_bytes_pag1:
        contents.append({"mime_type": "image/jpeg", "data": img_bytes_pag1})
    if img_bytes_pag2:
        contents.append({"mime_type": "image/jpeg", "data": img_bytes_pag2})
    contents.append(prompt)
    
    # INTENTO 1: Usando la nueva biblioteca google-genai
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)
        
        for modelo in modelos_a_probar:
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                if response.text:
                    # Limpiar y parsear JSON
                    json_str = response.text.strip()
                    if json_str.startswith("```json"):
                        json_str = json_str.split("```json")[1].split("```")[0].strip()
                    elif json_str.startswith("```"):
                        json_str = json_str.split("```")[1].split("```")[0].strip()
                    
                    data = json.loads(json_str)
                    return data, None
            except Exception:
                continue
    except Exception:
        pass
        
    # INTENTO 2: Fallback a la biblioteca tradicional google-generativeai
    try:
        import google.generativeai as genai_old
        genai_old.configure(api_key=api_key)
        
        # Convertir imágenes para la estructura antigua
        img_parts = []
        if img_bytes_pag1:
            img_parts.append({'mime_type': 'image/jpeg', 'data': img_bytes_pag1})
        if img_bytes_pag2:
            img_parts.append({'mime_type': 'image/jpeg', 'data': img_bytes_pag2})
            
        for modelo in modelos_a_probar:
            try:
                model_instance = genai_old.GenerativeModel(modelo)
                response = model_instance.generate_content([prompt] + img_parts)
                if response.text:
                    json_str = response.text.strip()
                    if json_str.startswith("```json"):
                        json_str = json_str.split("```json")[1].split("```")[0].strip()
                    elif json_str.startswith("```"):
                        json_str = json_str.split("```")[1].split("```")[0].strip()
                        
                    data = json.loads(json_str)
                    return data, None
            except Exception:
                continue
    except Exception as e:
        return None, f"Error general de inicialización de la API: {e}"
        
    return None, "Ninguno de los modelos de Gemini pudo procesar la solicitud."

# ==========================================
# 5. INTERFAZ GRÁFICA DE USUARIO (TABS)
# ==========================================

st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #4B5563;'>Sistematización Colaborativa en la Nube con IA - Evaluación de Conocimientos SRPA</p>", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📷 Carga de Encuestas", "✏️ Banco de Verificación", "📊 Dashboard de Impacto"])

# ------------------------------------------
# TAB 1: CARGA DE ENCUESTAS (FOTOS CELULAR)
# ------------------------------------------
with tab1:
    st.markdown("<h3 class='section-subheader'>1. Selección del Formulario</h3>", unsafe_allow_html=True)
    tipo_form = st.radio("Tipo de Encuesta a Digitalizar:", ["PRETEST", "POSTEST"], horizontal=True)
    
    st.markdown("<h3 class='section-subheader'>2. Capturar Fotos de la Encuesta Física (Doble Página)</h3>", unsafe_allow_html=True)
    st.info("💡 Consejo: Para encuestas de doble página, toma la foto bien de frente, con luz y enfocando las marcas y la caligrafía.")
    
    col_img1, col_img2 = st.columns(2)
    with col_img1:
        st.subheader("Página 1 (Encabezado y Preguntas)")
        file_pag1 = st.file_uploader("Subir Página 1", type=["png", "jpg", "jpeg"], key="uploader_p1")
    with col_img2:
        st.subheader("Página 2 (Preguntas / Satisfacción)")
        file_pag2 = st.file_uploader("Subir Página 2", type=["png", "jpg", "jpeg"], key="uploader_p2")
        
    if file_pag1 and file_pag2:
        if st.button("🚀 Procesar Encuesta con Gemini", use_container_width=True):
            with st.spinner("La Inteligencia Artificial está analizando la caligrafía y marcas..."):
                # Comprimir y convertir
                b64_p1, bytes_p1 = comprimir_y_convertir_base64(file_pag1)
                b64_p2, bytes_p2 = comprimir_y_convertir_base64(file_pag2)
                
                if b64_p1 and b64_p2:
                    # Enviar a Gemini
                    resultado, err = analizar_encuesta_con_gemini(bytes_p1, bytes_p2, tipo_form)
                    
                    if err:
                        st.error(f"Error procesando la encuesta con Gemini (Error en API): {err}")
                    else:
                        # Crear ID único para la encuesta
                        id_encuesta = f"ENC_{time.strftime('%Y%m%d_%H%M%S')}"
                        fecha_carga = time.strftime('%Y-%m-%d %H:%M:%S')
                        
                        # Extraer metadatos sugeridos
                        encabezado = resultado.get("encabezado", {})
                        municipio = encabezado.get("municipio", "")
                        ie_ia = encabezado.get("institucion_educativa", "")
                        rol = encabezado.get("rol", "")
                        
                        # Guardar en la Cola de Revisión
                        row_cola = [
                            id_encuesta,
                            fecha_carga,
                            tipo_form,
                            municipio,
                            ie_ia,
                            rol,
                            json.dumps(resultado),
                            b64_p1,
                            b64_p2,
                            "Pendiente"
                        ]
                        
                        success, row_num = append_row_and_verify(ws_cola, row_cola, id_col_index=1)
                        if success:
                            st.success(f"🎉 ¡Encuesta leída con éxito! Guardada en Cola de Revisión con ID: {id_encuesta}")
                            st.balloons()
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"Error al guardar en la nube (Google Sheets): {row_num}")

# ------------------------------------------
# TAB 2: BANCO DE VERIFICACIÓN (HUMAN-IN-THE-LOOP)
# ------------------------------------------
with tab2:
    st.markdown("<h3 class='section-subheader'>Verificación Humana de Caligrafía Manuscríta</h3>", unsafe_allow_html=True)
    st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
    
    # Cargar datos pendientes
    df_cola = cargar_cola_revision()
    
    if df_cola.empty or "Estado" not in df_cola.columns:
        st.info("📬 No hay encuestas pendientes de verificación en la nube en este momento.")
    else:
        # Filtrar pendientes
        pendientes = df_cola[df_cola['Estado'] == 'Pendiente']
        
        if pendientes.empty:
            st.info("📬 No hay encuestas pendientes de verificación en la nube en este momento.")
        else:
            st.write(f"Tienes **{len(pendientes)}** encuestas en cola esperando revisión.")
            
            # Seleccionar el primer registro pendiente de la cola
            registro = pendientes.iloc[0]
            id_actual = registro['ID_Encuesta']
            
            # Buscar el número de fila real en Google Sheets
            fila_real = df_cola[df_cola['ID_Encuesta'] == id_actual].index[0] + 2 # +2 por cabecera e índice de pandas
            
            # Decodificar imágenes para mostrar lado a lado
            try:
                img_p1_bytes = base64.b64decode(registro['Foto_Pag1_Base64'])
                img_p2_bytes = base64.b64decode(registro['Foto_Pag2_Base64'])
                img_p1 = Image.open(BytesIO(img_p1_bytes))
                img_p2 = Image.open(BytesIO(img_p2_bytes))
            except Exception as e:
                st.error(f"Error al cargar las imágenes de respaldo: {e}")
                img_p1, img_p2 = None, None
                
            # Mostrar panel comparativo visual
            col_preview1, col_preview2 = st.columns(2)
            with col_preview1:
                st.subheader("Página 1 (Original)")
                if img_p1:
                    st.image(img_p1, use_container_width=True)
            with col_preview2:
                st.subheader("Página 2 (Original)")
                if img_p2:
                    st.image(img_p2, use_container_width=True)
                    
            # Parsear el JSON de respuestas extraídas
            try:
                datos_ia = json.loads(registro['JSON_Respuestas'])
            except Exception:
                datos_ia = {}
                
            enc_ia = datos_ia.get("encabezado", {})
            
            # Formulario de corrección humana
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.subheader("✏️ Validación de Metadatos del Taller")
            
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                municipio_verificado = st.text_input("Municipio:", value=registro['Municipio'] or enc_ia.get("municipio", ""))
            with col_f2:
                # El campo más crítico: Caligrafía manuscrita de la Institución Educativa
                ie_verificada = st.text_input("🏫 Institución Educativa (Corregir caligrafía):", value=registro['Institucion_Educativa_IA'] or enc_ia.get("institucion_educativa", ""))
            with col_f3:
                rol_verificado = st.text_input("Rol del Participante:", value=registro['Rol'] or enc_ia.get("rol", ""))
                
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Botones de Acción
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True, type="primary"):
                    with st.spinner("Guardando y verificando consistencia en Google Sheets..."):
                        # Organizar fila definitiva para Respuestas_SRPA
                        conocimientos = datos_ia.get("respuestas_conocimiento", {})
                        satisfaccion = datos_ia.get("evaluacion_satisfaccion", {})
                        
                        row_respuestas = [
                            id_actual,
                            registro['Tipo_Formulario'],
                            enc_ia.get("fecha", ""),
                            municipio_verificado,
                            ie_verificada,
                            rol_verificado,
                            conocimientos.get("p1", ""), conocimientos.get("p2", ""), conocimientos.get("p3", ""), conocimientos.get("p4", ""),
                            conocimientos.get("p5", ""), conocimientos.get("p6", ""), conocimientos.get("p7", ""), conocimientos.get("p8", ""),
                            satisfaccion.get("sat_p1", ""), satisfaccion.get("sat_p2", ""), satisfaccion.get("sat_p3", ""), satisfaccion.get("sat_p4", ""),
                            satisfaccion.get("sat_p5", ""), satisfaccion.get("sat_p6", ""), satisfaccion.get("sat_p7", ""), satisfaccion.get("sat_p8", ""), satisfaccion.get("sat_p9", ""),
                            "Verificador Campo",
                            time.strftime('%Y-%m-%d %H:%M:%S')
                        ]
                        
                        # 1. Guardar y verificar en la hoja final
                        saved, res_info = append_row_and_verify(ws_respuestas, row_respuestas, id_col_index=1)
                        if saved:
                            # 2. Actualizar estado en la cola a "Aprobado" y verificar
                            status_updated, status_err = update_row_and_verify(ws_cola, fila_real, "Estado", "Aprobado")
                            if status_updated:
                                st.success("🎉 ¡Registro verificado e ingresado a la base de datos final con éxito!")
                                time.sleep(1.5)
                                st.rerun()
                            else:
                                st.error(f"Error al actualizar estado en la cola: {status_err}")
                        else:
                            st.error(f"Error guardando en la hoja final (Respuestas_SRPA): {res_info}")
                            
            with col_b2:
                if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                    with st.spinner("Eliminando entrada de la cola de revisión..."):
                        # Cambiar el estado a "Rechazado" en Google Sheets
                        rejected, reject_err = update_row_and_verify(ws_cola, fila_real, "Estado", "Rechazado")
                        if rejected:
                            st.warning("🗑️ Entrada rechazada y eliminada de la cola de visualización activa.")
                            time.sleep(1.5)
                            st.rerun()
                        else:
                            st.error(f"Error al rechazar la entrada en Google Sheets: {reject_err}")

# ------------------------------------------
# TAB 3: DASHBOARD DE IMPACTO (ESTADÍSTICAS REALES)
# ------------------------------------------
with tab3:
    st.markdown("<h3 class='section-subheader'>Dashboard y Analítica de Conocimiento (Datos Reales)</h3>", unsafe_allow_html=True)
    
    # Cargar base de datos definitiva
    df_respuestas = cargar_respuestas_validadas()
    
    if df_respuestas.empty:
        st.info("📊 La base de datos definitiva se encuentra limpia y vacía. Los gráficos aparecerán cuando apruebes la primera encuesta en la pestaña 'Banco de Verificación'.")
    else:
        st.write(f"Visualizando estadísticas de **{len(df_respuestas)}** registros consolidados de talleres en Bolívar.")
        
        # Filtros del Dashboard en la barra lateral o superior
        col_fil1, col_fil2 = st.columns(2)
        with col_fil1:
            municipios_disponibles = ["Todos"] + list(df_respuestas['Municipio'].unique())
            mun_selected = st.selectbox("Filtrar por Municipio:", municipios_disponibles)
        with col_fil2:
            ie_disponibles = ["Todas"] + list(df_respuestas['Institucion_Educativa_Verificada'].unique())
            ie_selected = st.selectbox("Filtrar por Institución Educativa:", ie_disponibles)
            
        # Filtrar DataFrame
        df_filtrado = df_respuestas.copy()
        if mun_selected != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Municipio'] == mun_selected]
        if ie_selected != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Institucion_Educativa_Verificada'] == ie_selected]
            
        if df_filtrado.empty:
            st.warning("No hay registros que coincidan con los filtros seleccionados.")
        else:
            # Gráficos Dinámicos
            import plotly.express as px
            
            col_graph1, col_graph2 = st.columns(2)
            with col_graph1:
                st.subheader("Distribución de Participantes por Rol")
                rol_counts = df_filtrado['Rol'].value_counts().reset_index()
                rol_counts.columns = ['Rol', 'Cantidad']
                fig_rol = px.pie(rol_counts, values='Cantidad', names='Rol', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_rol, use_container_width=True)
                
            with col_graph2:
                st.subheader("Tipos de Formularios Digitalizados")
                tipo_counts = df_filtrado['Tipo_Formulario'].value_counts().reset_index()
                tipo_counts.columns = ['Tipo', 'Cantidad']
                fig_tipo = px.bar(tipo_counts, x='Tipo', y='Cantidad', color='Tipo', color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig_tipo, use_container_width=True)
                
            # Evaluación de Conocimiento Pre vs Post (Si hay ambos tipos de datos)
            st.subheader("Evolución de Respuestas de Conocimiento")
            st.write("Muestra la distribución de las respuestas marcadas en los cuestionarios aprobados.")
            
            # Tomar P2 como ejemplo de conocimiento ("Finalidad del SRPA")
            if 'Conocimientos_P2' in df_filtrado.columns:
                p2_counts = df_filtrado.groupby(['Tipo_Formulario', 'Conocimientos_P2']).size().reset_index(name='Cantidad')
                fig_p2 = px.bar(p2_counts, x='Conocimientos_P2', y='Cantidad', color='Tipo_Formulario', barmode='group',
                                labels={'Conocimientos_P2': 'Respuestas a: ¿Cuál es la finalidad del SRPA?', 'Cantidad': 'Número de Participantes'},
                                color_discrete_sequence=px.colors.qualitative.Prism)
                st.plotly_chart(fig_p2, use_container_width=True)

                st.write(pd.DataFrame(sat_summary).fillna(0))
            else:
                st.info("Cargue cuestionarios de tipo POSTEST para visualizar la matriz de satisfacción.")

