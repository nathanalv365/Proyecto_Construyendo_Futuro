import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import json
import os
from datetime import datetime
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Proyecto Construyendo Futuro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo de CSS personalizado para vista móvil y tarjetas pulidas
st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
    }
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        border-left: 5px solid #1e3a8a;
        margin-bottom: 15px;
    }
    .main-header {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #1e3a8a;
        font-weight: 700;
        text-align: center;
        margin-bottom: 10px;
    }
    .sub-header {
        color: #475569;
        text-align: center;
        margin-bottom: 30px;
        font-size: 1.1em;
    }
    .review-image-container {
        border: 2px solid #e2e8f0;
        border-radius: 8px;
        padding: 5px;
        background-color: white;
        text-align: center;
    }
    .status-badge {
        padding: 4px 8px;
        border-radius: 12px;
        font-size: 0.85em;
        font-weight: 600;
    }
    .status-pending {
        background-color: #fef3c7;
        color: #d97706;
    }
    .status-approved {
        background-color: #dcfce7;
        color: #15803d;
    }
</style>
""", unsafe_allow_html=True)

# Inicializar estados de sesión para imágenes y cola
if "temp_images" not in st.session_state:
    st.session_state.temp_images = {}

# Columnas oficiales para las bases de datos
COLUMNS_RESPUESTAS = [
    "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
    "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4", 
    "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
    "Verificado_Por", "Fecha_Aprobacion"
]

COLUMNS_COLA = [
    "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol",
    "JSON_Respuestas", "Estado"
]

# --- CONEXIÓN GOOGLE SHEETS ---
def conectar_google_sheets():
    """Establece conexión con Google Sheets usando Secrets de Streamlit."""
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error autenticando con Google Cloud: {e}")
        return None

def obtener_libro_central(client):
    """Obtiene o crea el archivo central Base_Encuestas_SRPA."""
    try:
        return client.open("Base_Encuestas_SRPA")
    except gspread.SpreadsheetNotFound:
        # Crear libro si no existe
        sh = client.create("Base_Encuestas_SRPA")
        # Inicializar pestañas
        ws_cola = sh.add_worksheet(title="Cola_Revision", rows="100", cols="20")
        ws_resp = sh.add_worksheet(title="Respuestas_SRPA", rows="100", cols="30")
        
        # Eliminar pestaña default
        try:
            default_sheet = sh.worksheet("Sheet1")
            sh.del_worksheet(default_sheet)
        except:
            pass
            
        # Escribir encabezados
        ws_cola.append_row(COLUMNS_COLA)
        ws_resp.append_row(COLUMNS_RESPUESTAS)
        return sh

def cargar_cola_revision():
    """Carga los registros pendientes de revisión de Google Sheets."""
    client = conectar_google_sheets()
    if not client:
        return pd.DataFrame(columns=COLUMNS_COLA)
    
    sh = obtener_libro_central(client)
    ws = sh.worksheet("Cola_Revision")
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=COLUMNS_COLA)
    
    df = pd.DataFrame(records)
    # Validar que tenga las columnas correctas
    for col in COLUMNS_COLA:
        if col not in df.columns:
            df[col] = ""
    return df[df['Estado'] == 'Pendiente']

def cargar_respuestas_validadas():
    """Carga las encuestas consolidadas finales de Google Sheets."""
    client = conectar_google_sheets()
    if not client:
        return pd.DataFrame(columns=COLUMNS_RESPUESTAS)
    
    sh = obtener_libro_central(client)
    ws = sh.worksheet("Respuestas_SRPA")
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=COLUMNS_RESPUESTAS)
    
    df = pd.DataFrame(records)
    for col in COLUMNS_RESPUESTAS:
        if col not in df.columns:
            df[col] = ""
    return df

# --- PROCESAMIENTO OCR CON GEMINI 3.6 FLASH ---
def procesar_encuesta_vision(img1, img2):
    """Envía las imágenes de la página 1 y 2 a Gemini 3.6 Flash para OCR."""
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("Falta la clave GEMINI_API_KEY en los secretos.")
        return None

    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    
    # Redimensionar y optimizar imágenes para reducir latencia y costo de red
    def optimizar_imagen(img):
        img_copy = img.convert("RGB")
        img_copy.thumbnail((1200, 1600))
        buf = io.BytesIO()
        img_copy.save(buf, format="JPEG", quality=85)
        return buf.getvalue()

    p1_bytes = optimizar_imagen(img1)
    p2_bytes = optimizar_imagen(img2)

    prompt = """
    Eres un transcriptor de encuestas del Sistema de Responsabilidad Penal para Adolescentes (SRPA) del Proyecto 'Construyendo Futuro'.
    Analiza las dos imágenes proporcionadas (Página 1 y Página 2 de una sola encuesta física).
    
    IMPORTANTE: Los nombres de los participantes NO se recolectan. Debes enfocarte exclusivamente en extraer el nombre de la 'Institución Educativa' que está escrito a mano alzada en la parte superior.

    Instrucciones estrictas de extracción:
    1. Identifica el Tipo de Formulario: 'PRETEST' o 'POSTEST' según el encabezado.
    2. Extrae la fecha (formato YYYY-MM-DD si es legible, o tal cual esté escrita).
    3. Municipio: Extrae el municipio de la cabecera.
    4. Institución Educativa: Lee con cuidado la caligrafía manuscrita y extrae el nombre de la escuela.
    5. Rol: Identifica el rol del participante marcado con una 'X' (Estudiante, Docente, Padre de Familia, Lider comunitario).
    6. Respuestas de conocimiento (Sección A / Conocimientos):
       - Identifica la opción marcada con una 'X' o marca equivalente en cada pregunta.
       - Si una pregunta está en blanco (el participante no marcó ninguna opción), pon un string vacío (""). NO inventes ni supongas respuestas.
       - En PRETEST hay preguntas 1 a 8.
       - En POSTEST hay preguntas de conocimiento 1 a 5.
    7. Evaluación de satisfacción (Solo en POSTEST - Sección B):
       - Transcribe las marcas para las preguntas 1 a 9. Las opciones posibles son: 'Excelente', 'Bueno', 'Regular', 'Deficiente'.
       - Si la casilla está vacía, pon un string vacío ("").

    Retorna los datos únicamente como un objeto JSON estructurado que cumpla exactamente este formato:
    {
       "tipo_formulario": "PRETEST" o "POSTEST",
       "fecha": "fecha_detectada",
       "municipio": "municipio_detectado",
       "institucion_educativa": "escuela_manuscrita_detectada",
       "rol": "Rol",
       "conocimientos": {
          "p1": "opcion_marcada_o_vacia",
          "p2": "opcion_marcada_o_vacia",
          "p3": "opcion_marcada_o_vacia",
          "p4": "opcion_marcada_o_vacia",
          "p5": "opcion_marcada_o_vacia",
          "p6": "opcion_marcada_o_vacia",
          "p7": "opcion_marcada_o_vacia",
          "p8": "opcion_marcada_o_vacia"
       },
       "satisfaccion": {
          "sat_p1": "valor_o_vacio",
          "sat_p2": "valor_o_vacio",
          "sat_p3": "valor_o_vacio",
          "sat_p4": "valor_o_vacio",
          "sat_p5": "valor_o_vacio",
          "sat_p6": "valor_o_vacio",
          "sat_p7": "valor_o_vacio",
          "sat_p8": "valor_o_vacio",
          "sat_p9": "valor_o_vacio"
       }
    }
    """
    
    try:
        model = genai.GenerativeModel("gemini-3.6-flash")
        response = model.generate_content([
            {"mime_type": "image/jpeg", "data": p1_bytes},
            {"mime_type": "image/jpeg", "data": p2_bytes},
            prompt
        ])
        
        # Extraer JSON de la respuesta
        text = response.text.strip()
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()
            
        return json.loads(text)
    except Exception as e:
        st.error(f"Error procesando con Gemini (gemini-3.6-flash): {e}")
        return None

# --- VISTA DE LA APLICACIÓN ---
st.markdown("<h1 class='main-header'>Proyecto 'Construyendo Futuro'</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Evaluación de Conocimientos y Satisfacción en el SRPA</p>", unsafe_allow_html=True)

# Comprobar secretos de conexión
sheets_active = "gcp_service_account" in st.secrets

if not sheets_active:
    st.warning("⚠️ Google Sheets no está conectado actualmente.")
    with st.expander("🛠️ Cómo conectar tu Google Sheet central en 5 minutos"):
        st.write("""
        1. Crea un proyecto en la consola de Google Cloud, habilita las APIs de **Google Drive** y **Google Sheets**.
        2. Crea una cuenta de servicio, descarga la clave en formato **JSON** y comparte tu Google Sheet con el correo de esa cuenta de servicio en rol de **Editor**.
        3. En la consola de administración de Streamlit Cloud, ve a los secretos (**Secrets**) de tu aplicación y pega la información de la clave en este formato:
        
        ```toml
        GEMINI_API_KEY = "tu_clave_api_aquí"
        
        [gcp_service_account]
        type = "service_account"
        project_id = "tu-proyecto-id"
        private_key_id = "tu-private-key"
        private_key = "-----BEGIN PRIVATE KEY-----\nTU_LLAVE_PRIVADA\n-----END PRIVATE KEY-----\n"
        client_email = "tu-correo-de-cuenta-de-servicio@..."
        client_id = "tu-id"
        auth_uri = "https://accounts.google.com/o/oauth2/auth"
        token_uri = "https://oauth2.google.com/token"
        auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
        client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
        ```
        """)
else:
    # Definir pestañas de la aplicación
    tab_carga, tab_cola, tab_dashboard = st.tabs([
        "📤 Cargar Nuevas Encuestas", 
        "✍️ Cola de Revisión de Caligrafía", 
        "📈 Dashboard Estadístico"
    ])
    
    # -------------------------------------------------------------
    # TAB 1: CARGA DE NUEVAS ENCUESTAS
    # -------------------------------------------------------------
    with tab_carga:
        st.subheader("Subir Encuestas de Doble Página")
        st.write("Cada encuesta física consta de dos páginas. Súbelas juntas para consolidar sus datos con Inteligencia Artificial.")
        
        col1, col2 = st.columns(2)
        with col1:
            file_p1 = st.file_uploader("Subir foto de la Página 1 (Cabecera y Preguntas Iniciales)", type=["jpg", "jpeg", "png"], key="uploader_p1")
        with col2:
            file_p2 = st.file_uploader("Subir foto de la Página 2 (Preguntas Finales y Satisfacción)", type=["jpg", "jpeg", "png"], key="uploader_p2")
            
        if file_p1 and file_p2:
            img1 = Image.open(file_p1)
            img2 = Image.open(file_p2)
            
            # Mostrar vistas previas
            with st.expander("🔍 Ver fotos cargadas"):
                prev_c1, prev_c2 = st.columns(2)
                with prev_c1:
                    st.image(img1, caption="Página 1", use_container_width=True)
                with prev_c2:
                    st.image(img2, caption="Página 2", use_container_width=True)
            
            if st.button("🚀 Procesar Encuesta con Gemini 3.6 Flash", use_container_width=True, type="primary"):
                with st.spinner("La Inteligencia Artificial está analizando las marcas de las respuestas y la caligrafía manuscrita..."):
                    res = procesar_encuesta_vision(img1, img2)
                    
                    if res:
                        id_encuesta = f"ENC_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        
                        # Almacenar de forma persistente las imágenes en st.session_state
                        st.session_state.temp_images[id_encuesta] = {
                            "p1": img1,
                            "p2": img2
                        }
                        
                        # Preparar la fila para guardar en Cola_Revision
                        nueva_fila_cola = [
                            id_encuesta,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            res.get("tipo_formulario", "PRETEST"),
                            res.get("municipio", ""),
                            res.get("institucion_educativa", ""),
                            res.get("rol", "Estudiante"),
                            json.dumps(res),
                            "Pendiente"
                        ]
                        
                        try:
                            client = conectar_google_sheets()
                            sh = obtener_libro_central(client)
                            ws = sh.worksheet("Cola_Revision")
                            ws.append_row(nueva_fila_cola)
                            st.success(f"✅ ¡Encuesta procesada con éxito! Se ha agregado a la cola de revisión con el ID: {id_encuesta}")
                            st.info("💡 Dirígete a la pestaña 'Cola de Revisión de Caligrafía' para confirmar el nombre de la institución y validar los resultados.")
                        except Exception as e:
                            st.error(f"Error guardando en la cola de Google Sheets: {e}")
                            
    # -------------------------------------------------------------
    # TAB 2: COLA DE REVISIÓN (HUMAN-IN-THE-LOOP CON VISUALIZACIÓN DE IMAGEN)
    # -------------------------------------------------------------
    with tab_cola:
        st.subheader("Banco de Verificación de Escritura a Mano")
        st.write("Verifica la caligrafía manuscrita de la Institución Educativa y confirma los resultados antes de guardarlos permanentemente.")
        
        # Cargar registros pendientes en tiempo real
        df_cola = cargar_cola_revision()
        
        if df_cola.empty:
            st.success("🎉 ¡Excelente! No hay encuestas pendientes de revisar en la cola.")
        else:
            st.warning(f"Hay {len(df_cola)} encuestas pendientes de aprobación.")
            
            # Selector de encuestas a revisar
            lista_opciones = []
            for idx, r in df_cola.iterrows():
                lista_opciones.append(f"{r['ID_Encuesta']} - {r['Tipo_Formulario']} ({r['Municipio']} - IA Detectó: {r['Institucion_Educativa_IA']})")
                
            opcion_seleccionada = st.selectbox("Selecciona la encuesta a validar:", lista_opciones)
            
            if opcion_seleccionada:
                id_selected = opcion_seleccionada.split(" - ")[0]
                registro = df_cola[df_cola['ID_Encuesta'] == id_selected].iloc[0]
                
                # Cargar el JSON de respuestas original extraído por la IA
                res_ia = json.loads(registro['JSON_Respuestas'])
                
                # --- AQUÍ MOSTRAMOS LAS IMÁGENES ORIGINALES PARA ASEGURAR QUE SEAN CORRECTAS ---
                st.markdown("### 📷 Imagen de la Encuesta Física original")
                if id_selected in st.session_state.temp_images:
                    img_dict = st.session_state.temp_images[id_selected]
                    img_col1, img_col2 = st.columns(2)
                    with img_col1:
                        st.markdown("<div class='review-image-container'><b>Página 1 (Cabecera y Preguntas Iniciales)</b>", unsafe_allow_html=True)
                        st.image(img_dict["p1"], use_container_width=True)
                        st.markdown("</div>", unsafe_allow_html=True)
                    with img_col2:
                        st.markdown("<div class='review-image-container'><b>Página 2 (Preguntas Finales y Satisfacción)</b>", unsafe_allow_html=True)
                        st.image(img_dict["p2"], use_container_width=True)
                        st.markdown("</div>", unsafe_allow_html=True)
                else:
                    st.info("ℹ️ Las imágenes originales están disponibles en la sesión actual mientras se procesan en lote. "
                            "Si actualizó la página, aún puede realizar la validación guiándose de los campos sugeridos abajo.")
                    
                    # Dar la opción de volver a cargar las imágenes para visualización si es necesario
                    re_uploaded_p1 = st.file_uploader("Adjuntar Página 1 para ayuda visual (Opcional):", type=["jpg", "jpeg", "png"], key="re_upload_p1")
                    re_uploaded_p2 = st.file_uploader("Adjuntar Página 2 para ayuda visual (Opcional):", type=["jpg", "jpeg", "png"], key="re_upload_p2")
                    if re_uploaded_p1 and re_uploaded_p2:
                        st.session_state.temp_images[id_selected] = {
                            "p1": Image.open(re_uploaded_p1),
                            "p2": Image.open(re_uploaded_p2)
                        }
                        st.rerun()

                st.markdown("---")
                st.markdown("### ✍️ Formulario de Corrección de Datos")
                
                # Formulario editable
                with st.form("form_revision"):
                    col_f1, col_f2, col_f3 = st.columns(3)
                    with col_f1:
                        tipo_form = st.text_input("Tipo de Formulario:", registro['Tipo_Formulario'])
                    with col_f2:
                        fecha_val = st.text_input("Fecha:", res_ia.get("fecha", ""))
                    with col_f3:
                        mun_val = st.text_input("Municipio:", registro['Municipio'])
                        
                    # CAMPO CLAVE MANUSCRITO
                    escuela_verificada = st.text_input("✏️ Institución Educativa (Corregir caligrafía manuscrita aquí):", registro['Institucion_Educativa_IA'])
                    rol_val = st.selectbox("Rol del Participante:", ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(registro['Rol']) if registro['Rol'] in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0)
                    
                    st.markdown("##### Respuestas de Conocimiento Registradas")
                    con_data = res_ia.get("conocimientos", {})
                    
                    # Mostrar preguntas de conocimientos de forma interactiva
                    col_q1, col_q2 = st.columns(2)
                    con_actualizadas = {}
                    with col_q1:
                        con_actualizadas["p1"] = st.text_input("Pregunta 1:", con_data.get("p1", ""))
                        con_actualizadas["p2"] = st.text_input("Pregunta 2:", con_data.get("p2", ""))
                        con_actualizadas["p3"] = st.text_input("Pregunta 3:", con_data.get("p3", ""))
                        con_actualizadas["p4"] = st.text_input("Pregunta 4:", con_data.get("p4", ""))
                    with col_q2:
                        con_actualizadas["p5"] = st.text_input("Pregunta 5:", con_data.get("p5", ""))
                        con_actualizadas["p6"] = st.text_input("Pregunta 6:", con_data.get("p6", ""))
                        con_actualizadas["p7"] = st.text_input("Pregunta 7:", con_data.get("p7", ""))
                        con_actualizadas["p8"] = st.text_input("Pregunta 8:", con_data.get("p8", ""))
                        
                    sat_actualizadas = {}
                    if tipo_form == "POSTEST":
                        st.markdown("##### Matriz de Evaluación de Satisfacción (Excelente / Bueno / Regular / Deficiente)")
                        sat_data = res_ia.get("satisfaccion", {})
                        col_s1, col_s2 = st.columns(2)
                        with col_s1:
                            sat_actualizadas["sat_p1"] = st.selectbox("1. Claridad de la información:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p1", "")) if sat_data.get("sat_p1", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p2"] = st.selectbox("2. Dominio de facilitadores:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p2", "")) if sat_data.get("sat_p2", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p3"] = st.selectbox("3. Metodología utilizada:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p3", "")) if sat_data.get("sat_p3", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p4"] = st.selectbox("4. Participación e integración:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p4", "")) if sat_data.get("sat_p4", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p5"] = st.selectbox("5. Utilidad de los temas:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p5", "")) if sat_data.get("sat_p5", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                        with col_s2:
                            sat_actualizadas["sat_p6"] = st.selectbox("6. Organización de la jornada:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p6", "")) if sat_data.get("sat_p6", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p7"] = st.selectbox("7. Materiales y recursos:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p7", "")) if sat_data.get("sat_p7", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p8"] = st.selectbox("8. Fortaleció conocimientos:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p8", "")) if sat_data.get("sat_p8", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)
                            sat_actualizadas["sat_p9"] = st.selectbox("9. Recomendaría la jornada:", ["", "Excelente", "Bueno", "Regular", "Deficiente"], index=["", "Excelente", "Bueno", "Regular", "Deficiente"].index(sat_data.get("sat_p9", "")) if sat_data.get("sat_p9", "") in ["", "Excelente", "Bueno", "Regular", "Deficiente"] else 0)

                    # Botones de Acción
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        btn_aprobar = st.form_submit_button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True)
                    with col_btn2:
                        btn_rechazar = st.form_submit_button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True)

                    if btn_aprobar:
                        # Estructurar fila final para Respuestas_SRPA
                        fila_final = [
                            id_selected,
                            tipo_form,
                            fecha_val,
                            mun_val,
                            escuela_verificada,
                            rol_val,
                            con_actualizadas.get("p1", ""),
                            con_actualizadas.get("p2", ""),
                            con_actualizadas.get("p3", ""),
                            con_actualizadas.get("p4", ""),
                            con_actualizadas.get("p5", ""),
                            con_actualizadas.get("p6", ""),
                            con_actualizadas.get("p7", ""),
                            con_actualizadas.get("p8", ""),
                            sat_actualizadas.get("sat_p1", ""),
                            sat_actualizadas.get("sat_p2", ""),
                            sat_actualizadas.get("sat_p3", ""),
                            sat_actualizadas.get("sat_p4", ""),
                            sat_actualizadas.get("sat_p5", ""),
                            sat_actualizadas.get("sat_p6", ""),
                            sat_actualizadas.get("sat_p7", ""),
                            sat_actualizadas.get("sat_p8", ""),
                            sat_actualizadas.get("sat_p9", ""),
                            "Verificador_Móvil",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ]
                        
                        try:
                            client = conectar_google_sheets()
                            sh = obtener_libro_central(client)
                            
                            # 1. Insertar en pestaña final
                            ws_resp = sh.worksheet("Respuestas_SRPA")
                            ws_resp.append_row(fila_final)
                            
                            # 2. Actualizar estado en la cola (eliminar fila o marcar como Aprobado)
                            ws_cola = sh.worksheet("Cola_Revision")
                            celda = ws_cola.find(id_selected)
                            if celda:
                                ws_cola.update_cell(celda.row, 8, "Aprobado") # Columna 8 es 'Estado'
                                
                            st.success(f"🎉 ¡Registro {id_selected} aprobado e incorporado con éxito!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error procesando la aprobación en Google Sheets: {e}")
                            
                    if btn_rechazar:
                        try:
                            client = conectar_google_sheets()
                            sh = obtener_libro_central(client)
                            ws_cola = sh.worksheet("Cola_Revision")
                            celda = ws_cola.find(id_selected)
                            if celda:
                                ws_cola.update_cell(celda.row, 8, "Rechazado")
                            st.info(f"🗑️ Registro {id_selected} rechazado y retirado de la cola de revisión.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error procesando el rechazo en Google Sheets: {e}")

    # -------------------------------------------------------------
    # TAB 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
    # -------------------------------------------------------------
    with tab_dashboard:
        st.subheader("Indicadores de Conocimiento y Satisfacción")
        
        # Cargar respuestas validadas desde Sheets
        df_respuestas = cargar_respuestas_validadas()
        
        if df_respuestas.empty:
            st.info("📊 El dashboard se activará automáticamente una vez que apruebes las primeras encuestas cargadas.")
        else:
            # Filtros dinámicos en barra lateral o parte superior
            st.write("Filtros Interactivos:")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filtro_municipio = st.multiselect("Filtrar por Municipio:", options=df_respuestas['Municipio'].unique(), default=df_respuestas['Municipio'].unique())
            with col_f2:
                filtro_rol = st.multiselect("Filtrar por Rol de Participante:", options=df_respuestas['Rol'].unique(), default=df_respuestas['Rol'].unique())
                
            # Aplicar filtros
            df_filtrado = df_respuestas[
                (df_respuestas['Municipio'].isin(filtro_municipio)) & 
                (df_respuestas['Rol'].isin(filtro_rol))
            ]
            
            if df_filtrado.empty:
                st.warning("⚠️ No hay registros que coincidan con los filtros seleccionados.")
            else:
                # Métricas principales
                total_encuestas = len(df_filtrado)
                total_pre = len(df_filtrado[df_filtrado['Tipo_Formulario'] == "PRETEST"])
                total_post = len(df_filtrado[df_filtrado['Tipo_Formulario'] == "POSTEST"])
                
                kpi1, kpi2, kpi3 = st.columns(3)
                with kpi1:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <h3 style='margin:0;color:#1e3a8a;'>{total_encuestas}</h3>
                        <p style='margin:0;color:#64748b;'>Total Encuestas Consolidadas</p>
                    </div>
                    """, unsafe_allow_html=True)
                with kpi2:
                    st.markdown(f"""
                    <div class='metric-card' style='border-left-color: #2563eb;'>
                        <h3 style='margin:0;color:#2563eb;'>{total_pre}</h3>
                        <p style='margin:0;color:#64748b;'>Pretests Diligenciados</p>
                    </div>
                    """, unsafe_allow_html=True)
                with kpi3:
                    st.markdown(f"""
                    <div class='metric-card' style='border-left-color: #10b981;'>
                        <h3 style='margin:0;color:#10b981;'>{total_post}</h3>
                        <p style='margin:0;color:#64748b;'>Postests Diligenciados</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                # --- COMPARATIVA PRETEST VS POSTEST (IMPACTO EDUCATIVO) ---
                st.markdown("### 📚 Comparación de Conocimientos: Pretest vs Postest")
                st.write("Mide el porcentaje de aciertos en los conceptos fundamentales evaluados antes y después de los talleres.")
                
                # Definir respuestas correctas para Pretest y Postest de acuerdo con el formato oficial
                # P2 Pretest == P1 Postest ("Promover la responsabilidad...")
                # P3 Pretest == P2 Postest ("Consumir sustancias...")
                # P4 Pretest == P3 Postest ("Participar en actividades deportivas...")
                # P5 Pretest == P4 Postest ("La familia, la escuela...")
                
                conceptos = [
                    "Finalidad del SRPA", 
                    "Factores de Riesgo", 
                    "Factores Protectores", 
                    "Responsabilidad de Prevención"
                ]
                
                aciertos_pre = [0, 0, 0, 0]
                aciertos_post = [0, 0, 0, 0]
                
                df_pre = df_filtrado[df_filtrado['Tipo_Formulario'] == "PRETEST"]
                df_post = df_filtrado[df_filtrado['Tipo_Formulario'] == "POSTEST"]
                
                # Calcular aciertos PRETEST
                if not df_pre.empty:
                    # Finalidad: opción b (Promover la responsabilidad...)
                    aciertos_pre[0] = (df_pre['Conocimientos_P2'].str.lower().str.contains("responsabilidad|b", na=False).sum() / len(df_pre)) * 100
                    # Riesgo: opción b (Consumir sustancias...)
                    aciertos_pre[1] = (df_pre['Conocimientos_P3'].str.lower().str.contains("sustancias|b", na=False).sum() / len(df_pre)) * 100
                    # Protector: opción a (Dialogar con la familia)
                    aciertos_pre[2] = (df_pre['Conocimientos_P4'].str.lower().str.contains("dialogar|a", na=False).sum() / len(df_pre)) * 100
                    # Responsabilidad: opción d (La familia, la escuela...)
                    aciertos_pre[3] = (df_pre['Conocimientos_P5'].str.lower().str.contains("familia, la escuela|d", na=False).sum() / len(df_pre)) * 100
                    
                # Calcular aciertos POSTEST
                if not df_post.empty:
                    # Finalidad: opción b (Promover la responsabilidad...)
                    aciertos_post[0] = (df_post['Conocimientos_P1'].str.lower().str.contains("responsabilidad|b", na=False).sum() / len(df_post)) * 100
                    # Riesgo: opción b (Consumir sustancias...)
                    aciertos_post[1] = (df_post['Conocimientos_P2'].str.lower().str.contains("sustancias|b", na=False).sum() / len(df_post)) * 100
                    # Protector: opción b (Participar en actividades deportivas, culturales...)
                    aciertos_post[2] = (df_post['Conocimientos_P3'].str.lower().str.contains("deportivas|b", na=False).sum() / len(df_post)) * 100
                    # Responsabilidad: opción d (La familia, la escuela...)
                    aciertos_post[3] = (df_post['Conocimientos_P4'].str.lower().str.contains("familia, la escuela|d", na=False).sum() / len(df_post)) * 100
                    
                fig_impacto = go.Figure(data=[
                    go.Bar(name='Antes (Pretest)', x=conceptos, y=aciertos_pre, marker_color='#2563eb'),
                    go.Bar(name='Después (Postest)', x=conceptos, y=aciertos_post, marker_color='#10b981')
                ])
                fig_impacto.update_layout(
                    barmode='group',
                    yaxis_title='Porcentaje de Respuestas Correctas (%)',
                    yaxis_range=[0, 100],
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_impacto, use_container_width=True)
                
                # --- GRÁFICOS SECUNDARIOS (ROLES Y SATISFACCIÓN) ---
                g_col1, g_col2 = st.columns(2)
                
                with g_col1:
                    st.markdown("#### 👥 Participación por Rol")
                    df_rol_count = df_filtrado['Rol'].value_counts().reset_index()
                    df_rol_count.columns = ['Rol', 'Cantidad']
                    fig_rol = px.pie(df_rol_count, values='Cantidad', names='Rol', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig_rol, use_container_width=True)
                    
                with g_col2:
                    st.markdown("#### 😊 Nivel de Satisfacción General")
                    if df_post.empty:
                        st.info("Los datos de satisfacción se mostrarán cuando existan Postests aprobados.")
                    else:
                        # Mapeo numérico para promediar satisfacción
                        map_sat = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None}
                        
                        aspectos = [
                            "Claridad de información", "Dominio de facilitadores", "Metodología",
                            "Participación", "Utilidad de temas", "Organización",
                            "Materiales", "Fortaleció conocimientos", "Recomendaría jornada"
                        ]
                        
                        valores_sat = []
                        for i in range(1, 10):
                            col_sat = f"Sat_P{i}"
                            # Calcular promedio numérico
                            prom = df_post[col_sat].map(map_sat).mean()
                            valores_sat.append(prom if not pd.isna(prom) else 0)
                            
                        fig_sat = go.Figure(go.Bar(
                            x=valores_sat,
                            y=aspectos,
                            orientation='h',
                            marker_color='#f59e0b'
                        ))
                        fig_sat.update_layout(
                            xaxis_title="Puntuación Promedio (Escala 1 a 4)",
                            xaxis_range=[1, 4],
                            yaxis=dict(autorange="reversed")
                        )
                        st.plotly_chart(fig_sat, use_container_width=True)
                        
                # Botón de Descarga de Base de Datos Unificada
                st.markdown("---")
                st.subheader("📥 Descargar Base de Datos")
                st.write("Descarga los registros validados directamente en formato Excel compatible con tu computadora.")
                
                # Crear buffer de Excel en memoria
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_respuestas.to_excel(writer, sheet_name="Respuestas_SRPA", index=False)
                excel_buffer.seek(0)
                
                st.download_button(
                    label="💾 Descargar respuestas validadas (Excel)",
                    data=excel_buffer,
                    file_name="Base_Respuestas_SRPA_Consolidada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

