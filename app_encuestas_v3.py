import streamlit as st
import pandas as pd
import json
import io
import datetime
import base64
from PIL import Image

# Try importing the new Google GenAI SDK, fallback to the legacy google-generativeai
try:
    from google import genai
    from google.genai import types
    USE_NEW_SDK = True
except ImportError:
    try:
        import google.generativeai as genai_legacy
        USE_NEW_SDK = False
    except ImportError:
        USE_NEW_SDK = False

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Control SRPA - Proyecto Construyendo Futuro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado para hacer la app amigable en móviles
st.markdown("""
<style>
    .main-header {
        color: #1e3a8a;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-weight: 700;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-header {
        color: #4b5563;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-size: 14px;
        text-align: center;
        margin-bottom: 25px;
    }
    .card-pendiente {
        background-color: #fef3c7;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #f59e0b;
        margin-bottom: 15px;
    }
    .card-aprobado {
        background-color: #d1fae5;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #10b981;
        margin-bottom: 15px;
    }
    .stButton>button {
        border-radius: 20px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONEXIÓN A GOOGLE SHEETS
# ==============================================================================
COLS_COLA = [
    "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
    "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Base64_Pag1", "Base64_Pag2", "Estado"
]

COLS_RESPUESTAS = [
    "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
    "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4", 
    "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
    "Verificado_Por", "Fecha_Aprobacion"
]

def conectar_google_sheets():
    """Conecta con Google Sheets usando los secretos de Streamlit Cloud."""
    if not HAS_GSPREAD:
        return None
    
    if "gcp_service_account" not in st.secrets:
        return None
    
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_info = dict(st.secrets["gcp_service_account"])
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        
        # Intentar abrir la hoja
        try:
            sh = client.open("Base_Encuestas_SRPA")
        except gspread.exceptions.SpreadsheetNotFound:
            # Crear si no existe
            sh = client.create("Base_Encuestas_SRPA")
            # Compartir con el correo de servicio
            sh.share(creds_info["client_email"], perm_type="user", role="writer")
            
        return sh
    except Exception as e:
        st.sidebar.error(f"Error de credenciales de Google Sheets: {e}")
        return None

def obtener_hojas_trabajo():
    """Obtiene o inicializa las pestañas de Cola_Revision y Respuestas_SRPA."""
    sh = conectar_google_sheets()
    if sh is None:
        return None, None
    
    # 1. Pestaña Cola de Revisión
    try:
        ws_cola = sh.worksheet("Cola_Revision")
    except gspread.exceptions.WorksheetNotFound:
        ws_cola = sh.add_worksheet("Cola_Revision", rows=1000, cols=10)
        ws_cola.append_row(COLS_COLA)
        
    # 2. Pestaña Respuestas Consolidadas
    try:
        ws_resp = sh.worksheet("Respuestas_SRPA")
    except gspread.exceptions.WorksheetNotFound:
        ws_resp = sh.add_worksheet("Respuestas_SRPA", rows=1000, cols=25)
        ws_resp.append_row(COLS_RESPUESTAS)
        
    return ws_cola, ws_resp

def cargar_cola_revision():
    """Carga los datos de la Cola de Revisión de forma segura."""
    ws_cola, _ = obtener_hojas_trabajo()
    if ws_cola is None:
        return pd.DataFrame(columns=COLS_COLA)
    
    try:
        data = ws_cola.get_all_values()
        if len(data) <= 1:
            return pd.DataFrame(columns=COLS_COLA)
        
        df = pd.DataFrame(data[1:], columns=data[0])
        # Asegurar que todas las columnas requeridas existan
        for col in COLS_COLA:
            if col not in df.columns:
                df[col] = ""
        return df[COLS_COLA]
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame(columns=COLS_COLA)

def cargar_respuestas_validadas():
    """Carga los datos consolidados finales."""
    _, ws_resp = obtener_hojas_trabajo()
    if ws_resp is None:
        return pd.DataFrame(columns=COLS_RESPUESTAS)
    
    try:
        data = ws_resp.get_all_values()
        if len(data) <= 1:
            return pd.DataFrame(columns=COLS_RESPUESTAS)
        
        df = pd.DataFrame(data[1:], columns=data[0])
        # Asegurar que todas las columnas requeridas existan
        for col in COLS_RESPUESTAS:
            if col not in df.columns:
                df[col] = ""
        return df[COLS_RESPUESTAS]
    except Exception as e:
        st.error(f"Error cargando Respuestas SRPA: {e}")
        return pd.DataFrame(columns=COLS_RESPUESTAS)

# ==============================================================================
# COMPRESIÓN DE IMÁGENES PARA GOOGLE SHEETS
# ==============================================================================
def compress_and_encode_image(image_bytes):
    """Comprime la imagen para reducir tamaño de celda y la codifica a base64."""
    if not image_bytes:
        return ""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        # Ajustar dimensiones máximas a 350px para mantenerlo ultra liviano
        img.thumbnail((350, 350))
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=40)
        compressed_bytes = buffer.getvalue()
        return base64.b64encode(compressed_bytes).decode('utf-8')
    except Exception:
        return ""

# ==============================================================================
# MOTOR OCR GEMINI 3.6-FLASH
# ==============================================================================
def analizar_encuesta_con_gemini(api_key, pag1_bytes, pag2_bytes):
    """Envía la página 1 y página 2 unificadas a Gemini 3.6-flash para extraer la info."""
    if not api_key:
        raise ValueError("Falta configurar tu GEMINI_API_KEY en los secretos.")
        
    prompt = """
    Eres un sistema OCR especializado en procesar encuestas físicas de dos páginas del "Proyecto Construyendo Futuro" (SRPA).
    Analiza detalladamente las dos imágenes adjuntas (Página 1 y Página 2 de la misma encuesta) y extrae la información de manera integrada.
    
    Reglas estrictas de extracción:
    1. Identifica el tipo de formulario ("PRETEST" o "POSTEST") revisando los títulos superiores.
    2. Extrae los datos del encabezado:
       - Fecha: de la cabecera (formato YYYY-MM-DD o deja vacío "" si no está claro).
       - Municipio: el texto manuscrito al lado de "Municipio:".
       - Institución Educativa: el texto escrito a mano alzada por el participante.
       - Rol del participante: identifica cuál de las casillas (Estudiante, Docente, Padre de Familia, Lider comunitario) está marcada con una 'X' o tachada.
    3. Lee las respuestas marcadas con una equis (X) en las preguntas de selección múltiple:
       - En PRETEST: Preguntas de conocimiento 1 a 8.
       - En POSTEST: Sección A (Preguntas de conocimiento 1 a 5) y Sección B (Evaluación de Satisfacción 1 a 9, con opciones: Excelente, Bueno, Regular, Deficiente).
    
    IMPORTANTE: Si una pregunta o celda está en blanco, no tiene marcas o es ilegible, no inventes información. Deja el valor como una cadena de texto vacía "". No auto-rellenes por defecto.
    
    Responde estrictamente en un formato JSON plano estructurado exactamente así:
    {
      "tipo_formulario": "PRETEST" o "POSTEST",
      "encabezado": {
        "fecha": "fecha encontrada",
        "municipio": "municipio encontrado",
        "institucion_educativa": "nombre de la escuela manuscrito",
        "rol": "Estudiante" o "Docente" o "Padre de Familia" o "Lider comunitario"
      },
      "respuestas_conocimiento": {
        "p1": "letra_marcada_o_vacio",
        "p2": "letra_marcada_o_vacio",
        ...
        "p8": "letra_marcada_o_vacio"
      },
      "evaluacion_satisfaccion": {  // Solo si es POSTEST, si es PRETEST deja el objeto vacío o nulo
        "s1": "Excelente" o "Bueno" o "Regular" o "Deficiente" o "",
        ...
        "s9": "Excelente" o "Bueno" o "Regular" o "Deficiente" o ""
      }
    }
    """
    
    try:
        if USE_NEW_SDK:
            # Nuevo SDK de Google GenAI
            client = genai.Client(api_key=api_key)
            contents = [
                types.Part.from_bytes(data=pag1_bytes, mime_type="image/jpeg"),
                types.Part.from_bytes(data=pag2_bytes, mime_type="image/jpeg"),
                prompt
            ]
            response = client.models.generate_content(
                model='gemini-3.6-flash',
                contents=contents,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            text_response = response.text
        else:
            # Legacy SDK google-generativeai
            genai_legacy.configure(api_key=api_key)
            model = genai_legacy.GenerativeModel('gemini-3.6-flash')
            
            # Formatear imágenes para el SDK antiguo
            p1_part = {"mime_type": "image/jpeg", "data": pag1_bytes}
            p2_part = {"mime_type": "image/jpeg", "data": pag2_bytes}
            
            response = model.generate_content(
                contents=[p1_part, p2_part, prompt],
                generation_config={"response_mime_type": "application/json"}
            )
            text_response = response.text
            
        # Intentar parsear el JSON recibido
        # Limpiar posibles bloques de código de markdown si la IA los puso
        if "```json" in text_response:
            text_response = text_response.split("```json")[1].split("```")[0].strip()
        elif "```" in text_response:
            text_response = text_response.split("```")[1].split("```")[0].strip()
            
        return json.loads(text_response.strip())
    except Exception as e:
        raise RuntimeError(f"Error procesando con Gemini (gemini-3.6-flash): {e}")

# ==============================================================================
# MENU LATERAL - CONFIGURACIÓN Y CREDENCIALES
# ==============================================================================
st.sidebar.image("https://www.bolivar.gov.co/images/Gobernacion/Logo-Gobernacion-Color.png", use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.header("🔑 Configuración de Servicios")

# Obtener clave API de Gemini
gemini_key = st.secrets.get("GEMINI_API_KEY", "")
if not gemini_key:
    gemini_key = st.sidebar.text_input("Ingrese su Google Gemini API Key:", type="password")
else:
    st.sidebar.success("✅ Clave Gemini API configurada")

# Estado de Google Sheets
ws_cola, ws_resp = obtener_hojas_trabajo()
if ws_cola is not None and ws_resp is not None:
    st.sidebar.success("✅ Conectado a Google Sheets")
    sheets_ready = True
else:
    st.sidebar.warning("⚠️ No Conectado a Google Sheets")
    sheets_ready = False

# Navegación Principal
st.sidebar.markdown("---")
st.sidebar.header("📂 Navegación")
opcion = st.sidebar.radio(
    "Seleccione una pestaña:",
    ["Cargar Encuesta", "Cola de Revisión", "Dashboard Estadístico"]
)

# Panel de instrucciones en caso de que no esté configurado el Google Sheet
def mostrar_ayuda_credenciales():
    st.info("### ℹ️ Cómo activar la sincronización con Google Sheets")
    st.markdown("""
    Para que las encuestas cargadas desde el celular se guarden colaborativamente en tiempo real, debes configurar los secretos en tu consola de Streamlit Cloud:
    
    1. **Crea un archivo de Google Sheet** en tu cuenta de Drive llamado exactamente **`Base_Encuestas_SRPA`**.
    2. Crea una **Cuenta de Servicio** en tu consola de Google Cloud Developer, activa la API de Google Sheets y Google Drive, y descarga el archivo de credenciales en formato `.json`.
    3. Comparte tu Google Sheet con el correo de la cuenta de servicio como **Editor**.
    4. Ve a la consola de **Streamlit Cloud** ➔ selecciona tu App ➔ clic en **Manage app** ➔ **Settings** ➔ **Secrets** y pega tus credenciales en este formato:
    
    ```toml
    GEMINI_API_KEY = "tu_clave_api_aquí"
    
    [gcp_service_account]
    type = "service_account"
    project_id = "tu-proyecto-id"
    private_key_id = "tu-private-key-id"
    private_key = "-----BEGIN PRIVATE KEY-----\\nTU_LLAVE_PRIVADA\\n-----END PRIVATE KEY-----\\n"
    client_email = "tu-cuenta-de-servicio@proyecto.iam.gserviceaccount.com"
    client_id = "tu-client-id"
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.google.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/tu-cuenta-de-servicio"
    ```
    """)

# ==============================================================================
# PESTAÑA 1: CARGAR ENCUESTA (CARGA MÓVIL Y ANÁLISIS)
# ==============================================================================
if opcion == "Cargar Encuesta":
    st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Carga masiva de encuestas de doble página mediante fotos desde el celular</p>", unsafe_allow_html=True)
    
    if not sheets_ready:
        mostrar_ayuda_credenciales()
    else:
        st.markdown("### 📷 Capturar o seleccionar imágenes")
        
        col_img1, col_img2 = st.columns(2)
        with col_img1:
            st.markdown("**1. Foto de la PÁGINA 1** (Cabecera, Municipio, Escuela, Preguntas 1-3)")
            pag1_file = st.file_uploader("Subir Página 1", type=["jpg", "jpeg", "png"], key="pag1_upload", label_visibility="collapsed")
            if pag1_file:
                st.image(pag1_file, use_container_width=True, caption="Vista previa Página 1")
                
        with col_img2:
            st.markdown("**2. Foto de la PÁGINA 2** (Preguntas de conocimiento restantes y Satisfacción)")
            pag2_file = st.file_uploader("Subir Página 2", type=["jpg", "jpeg", "png"], key="pag2_upload", label_visibility="collapsed")
            if pag2_file:
                st.image(pag2_file, use_container_width=True, caption="Vista previa Página 2")
                
        if pag1_file and pag2_file:
            st.markdown("---")
            st.markdown("### 🧠 Procesar Encuesta")
            municipio_por_defecto = st.text_input("Municipio por defecto (por si no está escrito):", "Cartagena")
            
            if st.button("🔍 Digitalizar con Gemini 3.6-flash", use_container_width=True):
                if not gemini_key:
                    st.error("Por favor ingrese su API Key de Gemini en el panel lateral.")
                else:
                    with st.spinner("La Inteligencia Artificial está analizando la caligrafía y marcas de ambas páginas..."):
                        try:
                            # Leer bytes
                            p1_bytes = pag1_file.getvalue()
                            p2_bytes = pag2_file.getvalue()
                            
                            # Analizar
                            res = analizar_encuesta_con_gemini(gemini_key, p1_bytes, p2_bytes)
                            
                            # Comprimir para guardar en cola
                            b64_p1 = compress_and_encode_image(p1_bytes)
                            b64_p2 = compress_and_encode_image(p2_bytes)
                            
                            # Generar ID único
                            id_encuesta = f"ENCR_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                            
                            # Completar datos si vienen vacíos
                            enc = res.get("encabezado", {})
                            mun = enc.get("municipio") or municipio_por_defecto
                            ie_ia = enc.get("institucion_educativa", "")
                            rol = enc.get("rol", "")
                            tipo = res.get("tipo_formulario", "PRETEST")
                            
                            # Registrar en Cola de Revisión
                            nuevo_registro = [
                                id_encuesta,
                                datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                tipo,
                                mun,
                                ie_ia,
                                rol,
                                json.dumps(res),
                                b64_p1,
                                b64_p2,
                                "Pendiente"
                            ]
                            
                            ws_cola.append_row(nuevo_registro)
                            st.success("🎉 ¡Encuesta digitalizada con éxito y enviada a la Cola de Revisión!")
                            st.balloons()
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"Error procesando la encuesta: {e}")

# ==============================================================================
# PESTAÑA 2: COLA DE REVISIÓN (VERIFICACIÓN HUMANA)
# ==============================================================================
elif opcion == "Cola de Revisión":
    st.markdown("<h1 class='main-header'>Banco de Verificación de Escritura a Mano</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente</p>", unsafe_allow_html=True)
    
    if not sheets_ready:
        mostrar_ayuda_credenciales()
    else:
        df_cola = cargar_cola_revision()
        df_pendientes = df_cola[df_cola["Estado"] == "Pendiente"]
        
        if df_pendientes.empty:
            st.info("🎉 ¡Felicidades! No hay encuestas pendientes de verificación en la cola. Todas las encuestas cargadas han sido aprobadas.")
        else:
            st.markdown(f"Hay **{len(df_pendientes)}** encuesta(s) pendiente(s) de revisión en campo.")
            
            # Selector de encuesta
            lista_ids = df_pendientes["ID_Encuesta"].tolist()
            encuesta_id_selected = st.selectbox("Seleccione la encuesta a revisar:", lista_ids)
            
            row = df_pendientes[df_pendientes["ID_Encuesta"] == encuesta_id_selected].iloc[0]
            
            # Intentar decodificar los datos JSON
            try:
                datos_ia = json.loads(row["JSON_Respuestas"])
            except Exception:
                datos_ia = {}
                
            st.markdown("---")
            
            # Mostrar imágenes lado a lado para verificación
            st.markdown("### 📸 Documentos Originales Digitalizados")
            col_ver1, col_ver2 = st.columns(2)
            
            with col_ver1:
                st.markdown("**Foto Página 1 (Cabecera):**")
                b64_p1 = row.get("Base64_Pag1", "")
                if b64_p1:
                    try:
                        img_data = base64.b64decode(b64_p1)
                        st.image(Image.open(io.BytesIO(img_data)), use_container_width=True)
                    except Exception:
                        st.warning("No se pudo decodificar la imagen de la Página 1.")
                else:
                    st.info("Imagen no disponible en esta entrada.")
                    
            with col_ver2:
                st.markdown("**Foto Página 2 (Satisfacción/Respuestas):**")
                b64_p2 = row.get("Base64_Pag2", "")
                if b64_p2:
                    try:
                        img_data = base64.b64decode(b64_p2)
                        st.image(Image.open(io.BytesIO(img_data)), use_container_width=True)
                    except Exception:
                        st.warning("No se pudo decodificar la imagen de la Página 2.")
                else:
                    st.info("Imagen no disponible en esta entrada.")
            
            st.markdown("---")
            st.markdown("### ✏️ Panel de Validación de Datos")
            
            # Formulario de corrección
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                tipo_form = st.selectbox("Tipo de Formulario:", ["PRETEST", "POSTEST"], index=0 if row["Tipo_Formulario"] == "PRETEST" else 1)
            with col_f2:
                fecha_form = st.text_input("Fecha:", datos_ia.get("encabezado", {}).get("fecha", datetime.date.today().strftime('%Y-%m-%d')))
            with col_f3:
                mun_form = st.text_input("Municipio:", row["Municipio"])
                
            col_f4, col_f5 = st.columns(2)
            with col_f4:
                # CAMPO CRÍTICO: Corrección de caligrafía de Institución Educativa
                escuela_detectada = row["Institucion_Educativa_IA"]
                escuela_verificada = st.text_input(
                    "🏫 NOMBRE DE LA INSTITUCIÓN EDUCATIVA (Verifica y edita si es necesario):",
                    value=escuela_detectada,
                    help="Corrige la interpretación que hizo la IA de la escritura manuscrita."
                )
            with col_f5:
                rol_form = st.selectbox(
                    "Rol del Participante:", 
                    ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], 
                    index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(row["Rol"]) if row["Rol"] in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0
                )
                
            # Acciones finales: Aprobar o Rechazar
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True, type="primary"):
                    with st.spinner("Consolidando registro..."):
                        try:
                            # 1. Recuperar respuestas de conocimiento y satisfacción
                            resp_con = datos_ia.get("respuestas_conocimiento", {})
                            resp_sat = datos_ia.get("evaluacion_satisfaccion", {}) if tipo_form == "POSTEST" else {}
                            
                            # Formatear la fila de Respuestas SRPA
                            con_cols = [resp_con.get(f"p{i}", "") for i in range(1, 9)]
                            sat_cols = [resp_sat.get(f"s{i}", "") if tipo_form == "POSTEST" else "" for i in range(1, 10)]
                            
                            nueva_fila = [
                                row["ID_Encuesta"],
                                tipo_form,
                                fecha_form,
                                mun_form,
                                escuela_verificada,
                                rol_form
                            ] + con_cols + sat_cols + [
                                "Coordinador Campo",
                                datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            ]
                            
                            # Escribir en base consolidada
                            ws_resp.append_row(nueva_fila)
                            
                            # Cambiar estado en Cola a "Aprobado"
                            celda = ws_cola.find(row["ID_Encuesta"])
                            ws_cola.update_cell(celda.row, 10, "Aprobado")
                            
                            st.success("✅ ¡Encuesta aprobada e integrada permanentemente a la base de datos central!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error consolidando aprobación: {e}")
                            
            with col_btn2:
                if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                    with st.spinner("Rechazando entrada..."):
                        try:
                            # Cambiar estado en Cola a "Rechazado"
                            celda = ws_cola.find(row["ID_Encuesta"])
                            ws_cola.update_cell(celda.row, 10, "Rechazado")
                            st.warning("❌ Entrada rechazada y borrada de la cola de pendientes.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error rechazando entrada: {e}")

# ==============================================================================
# PESTAÑA 3: DASHBOARD ESTADÍSTICO (Métricas en Tiempo Real)
# ==============================================================================
elif opcion == "Dashboard Estadístico":
    st.markdown("<h1 class='main-header'>Dashboard Estadístico del Proyecto</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Medición del impacto educativo de prevención del SRPA y nivel de satisfacción de las jornadas</p>", unsafe_allow_html=True)
    
    if not sheets_ready:
        mostrar_ayuda_credenciales()
    else:
        df_resp = cargar_respuestas_validadas()
        
        if df_resp.empty:
            st.info("ℹ️ No hay registros consolidados en la base de datos de Google Sheets actualmente. Comienza a cargar encuestas para ver las estadísticas.")
        else:
            # ------------------------------------------------------------------
            # FILTROS DE DASHBOARD
            # ------------------------------------------------------------------
            municipios_list = ["Todos"] + df_resp["Municipio"].dropna().unique().tolist()
            escuelas_list = ["Todas"] + df_resp["Institucion_Educativa_Verificada"].dropna().unique().tolist()
            
            col_fil1, col_fil2 = st.columns(2)
            with col_fil1:
                mun_sel = st.selectbox("Filtrar por Municipio:", municipios_list)
            with col_fil2:
                esc_sel = st.selectbox("Filtrar por Institución Educativa:", escuelas_list)
                
            # Aplicar filtros
            df_filtered = df_resp.copy()
            if mun_sel != "Todos":
                df_filtered = df_filtered[df_filtered["Municipio"] == mun_sel]
            if esc_sel != "Todas":
                df_filtered = df_filtered[df_filtered["Institucion_Educativa_Verificada"] == esc_sel]
                
            # ------------------------------------------------------------------
            # TARJETAS DE INDICADORES (KPIs)
            # ------------------------------------------------------------------
            total_reg = len(df_filtered)
            pretests_count = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
            postests_count = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
            
            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                st.metric("Total Encuestas Consolidadas", total_reg)
            with col_kpi2:
                st.metric("Total Pretests Procesados", pretests_count)
            with col_kpi3:
                st.metric("Total Postests Procesados", postests_count)
                
            st.markdown("---")
            
            # ------------------------------------------------------------------
            # GRÁFICOS: COMPARATIVO PRE vs POST (IMPACTO)
            # ------------------------------------------------------------------
            st.markdown("### 📈 Medición de Impacto de Conocimientos (Línea Base vs Postest)")
            st.markdown("El gráfico analiza el porcentaje de respuestas correctas sobre los 4 conceptos clave abordados en los talleres:")
            
            # Calcular porcentajes de aciertos
            pre_df = df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"]
            post_df = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]
            
            conceptos = ["Finalidad del SRPA", "Factores de Riesgo", "Factores Protectores", "Corresponsabilidad de Prevención"]
            pre_scores = []
            post_scores = []
            
            # Mapeo de respuestas correctas
            # Pretest: Q2=b, Q3=b, Q4=a, Q5=d
            # Postest: Q1=b, Q2=b, Q3=b, Q4=d
            
            if not pre_df.empty:
                pre_scores.append((pre_df["Conocimientos_P2"].str.lower() == "b").mean() * 100)
                pre_scores.append((pre_df["Conocimientos_P3"].str.lower() == "b").mean() * 100)
                pre_scores.append((pre_df["Conocimientos_P4"].str.lower() == "a").mean() * 100)
                pre_scores.append((pre_df["Conocimientos_P5"].str.lower() == "d").mean() * 100)
            else:
                pre_scores = [0, 0, 0, 0]
                
            if not post_df.empty:
                post_scores.append((post_df["Conocimientos_P1"].str.lower() == "b").mean() * 100)
                post_scores.append((post_df["Conocimientos_P2"].str.lower() == "b").mean() * 100)
                post_scores.append((post_df["Conocimientos_P3"].str.lower() == "b").mean() * 100)
                post_scores.append((post_df["Conocimientos_P4"].str.lower() == "d").mean() * 100)
            else:
                post_scores = [0, 0, 0, 0]
                
            fig_impacto = go.Figure(data=[
                go.Bar(name='Línea Base (Pretest)', x=conceptos, y=pre_scores, marker_color='#ef4444'),
                go.Bar(name='Resultado (Postest)', x=conceptos, y=post_scores, marker_color='#10b981')
            ])
            fig_impacto.update_layout(
                barmode='group',
                yaxis_title='% de Respuestas Correctas',
                yaxis_range=[0, 100],
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_impacto, use_container_width=True)
            
            # ------------------------------------------------------------------
            # GRÁFICOS: SATISFACCIÓN Y ROLES
            # ------------------------------------------------------------------
            st.markdown("---")
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                st.markdown("### 👥 Distribución por Roles de Asistentes")
                roles_dist = df_filtered["Rol"].value_counts().reset_index()
                roles_dist.columns = ["Rol", "Cantidad"]
                fig_roles = px.pie(roles_dist, values='Cantidad', names='Rol', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_roles, use_container_width=True)
                
            with col_g2:
                st.markdown("### 🌟 Evaluación de Satisfacción de la Jornada (Postest)")
                if post_df.empty:
                    st.info("No hay datos de postest cargados para mostrar niveles de satisfacción.")
                else:
                    # Mapeo de valores de satisfacción a escala numérica
                    map_satisfaccion = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None}
                    
                    aspectos_nombres = [
                        "Claridad Info.", "Dominio Facilitadores", "Metodología", "Participación",
                        "Utilidad Temas", "Organización", "Materiales", "Fortaleció Conoc.", "Recomendaría"
                    ]
                    
                    promedios_sat = []
                    for i in range(1, 10):
                        col_name = f"Sat_P{i}"
                        if col_name in post_df.columns:
                            valores_numericos = post_df[col_name].map(map_satisfaccion).dropna()
                            promedios_sat.append(valores_numericos.mean() if not valores_numericos.empty else 0)
                        else:
                            promedios_sat.append(0)
                            
                    fig_sat = go.Figure(data=[
                        go.Bar(x=aspectos_nombres, y=promedios_sat, marker_color='#3b82f6', name="Nivel promedio")
                    ])
                    fig_sat.add_hline(y=3.0, line_dash="dash", line_color="red", annotation_text="Meta (Bueno)")
                    fig_sat.update_layout(
                        yaxis_title="Puntuación Promedio (1 a 4)",
                        yaxis_range=[1, 4]
                    )
                    st.plotly_chart(fig_sat, use_container_width=True)
