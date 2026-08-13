import streamlit as st
import pandas as pd
import json
import io
import datetime
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go

# Configuración de compatibilidad para Google Gemini SDK
try:
    from google import genai
    from google.genai import types
    USE_NEW_SDK = True
except ImportError:
    try:
        import google.generativeai as genai
        USE_NEW_SDK = False
    except ImportError:
        st.error("No se encontraron las librerías de Google Gemini. Por favor instala google-generativeai en tu entorno.")

# Configuración de Google Sheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CONFIGURACIÓN DE ESTILOS CSS PARA DISPOSITIVOS MÓVILES ---
st.markdown("""
<style>
    /* Estilos móviles optimizados */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 12px;
        font-weight: 600;
        padding: 0.6rem 1rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f1f5f9;
        border-radius: 8px;
        padding: 10px 16px;
        font-size: 14px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563eb !important;
        color: white !important;
    }
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 1.2rem;
        border-radius: 16px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05);
    }
</style>
""", unsafe_allowed_value=True)

# --- PROMPT OCR CONFIGURATION ---
SYSTEM_PROMPT = """
Eres un analizador OCR experto para el "Proyecto Construyendo Futuro" (SRPA).
Analiza las dos imágenes proporcionadas correspondientes a la Página 1 y Página 2 de un mismo cuestionario.
Extrae la información en formato JSON estricto sin preámbulos ni explicaciones.

Reglas importantes de extracción:
1. Identifica si el formulario es "PRETEST" o "POSTEST" buscando el título superior en la Página 1.
2. Lee los datos manuscritos del encabezado (Página 1):
   - Fecha (Intenta estandarizar a YYYY-MM-DD, p.ej. "30/07/2026" -> "2026-07-30")
   - Municipio (p.ej. "Cartagena")
   - Institución Educativa (Busca el texto escrito en la línea, p.ej. "Promesa de Dios")
   - Rol del participante (Marca con una X la opción correspondiente: Estudiante, Docente, Padre de Familia, Lider comunitario)
3. Lee las respuestas de las preguntas de conocimientos marcadas con (X) en ambas páginas:
   - Para PRETEST: Extrae respuestas para las preguntas de conocimientos 1 a 8.
   - Para POSTEST: Sección A de conocimientos contiene preguntas 1 a 5.
   - Si una pregunta de conocimientos no tiene respuesta o está en blanco, DEVUELVE la cadena vacía "". No inventes respuestas.
4. Para POSTEST: Lee la SECCIÓN B (Evaluación de satisfacción, preguntas 1 a 9).
   - Registra el valor de la opción marcada con (X) ("Excelente", "Bueno", "Regular", "Deficiente").
   - Si está en blanco o sin marcar, devuelve "". No inventes respuestas.

Esquema JSON obligatorio de salida:
{
  "tipo_formulario": "PRETEST" o "POSTEST",
  "encabezado": {
    "fecha": "YYYY-MM-DD",
    "municipio": "Nombre Municipio",
    "institucion_educativa_ia": "Nombre manuscrito detectado de la Institución",
    "rol": "Estudiante" | "Docente" | "Padre de Familia" | "Lider comunitario"
  },
  "conocimientos": {
    "p1": "a"|"b"|"c"|"",
    "p2": "a"|"b"|"c"|"d"|"",
    ...
    "p8": "a"|"b"|"c"|"d"|"" (solo si es PRETEST)
  },
  "satisfaccion": { (Solo si es POSTEST, si es PRETEST este bloque debe estar vacío o ser null)
    "sat_1": "Excelente"|"Bueno"|"Regular"|"Deficiente"|"",
    ...
    "sat_9": "Excelente"|"Bueno"|"Regular"|"Deficiente"|""
  }
}
"""

# --- CONEXIÓN A GOOGLE SHEETS ---
def conectar_google_sheets():
    """Establece conexión con Google Sheets usando secretos de Streamlit."""
    if not GSPREAD_AVAILABLE:
        st.error("Librerías de Google Sheets no disponibles en el entorno.")
        return None
    
    if "gcp_service_account" not in st.secrets:
        # Mostrar panel instructivo interactivo si no está configurado
        st.info("ℹ️ Conexión con Google Sheets no configurada.")
        with st.expander("🛠️ Ver Instrucciones de Configuración", expanded=True):
            st.markdown("""
            ### Cómo conectar tu Google Sheet central en 3 pasos:
            1. **Crea un proyecto en Google Cloud Console** y activa las APIs de **Google Sheets** y **Google Drive**.
            2. Crea una **Cuenta de Servicio**, descarga la llave en formato **JSON** y compártela con tu hoja de cálculo dándole permisos de **Editor** al correo de la cuenta de servicio.
            3. Ve al panel de control de **Streamlit Cloud** -> **Settings** -> **Secrets** y pega tus credenciales así:
            ```toml
            GEMINI_API_KEY = "tu-clave-api"
            
            [gcp_service_account]
            type = "service_account"
            project_id = "tu-proyecto-id"
            private_key = "-----BEGIN PRIVATE KEY-----\\nTU_LLAVE_PRIVADA\\n-----END PRIVATE KEY-----\\n"
            client_email = "tu-cuenta-de-servicio@correo.com"
            ... (el contenido completo de tu JSON)
            ```
            """)
        return None

    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de autenticación con Google Cloud: {e}")
        return None

def inicializar_hoja_y_pestañas():
    """Busca o crea la hoja de cálculo y asegura que tenga las pestañas correctas."""
    conn = conectar_google_sheets()
    if conn is None:
        return None
    try:
        try:
            sheet = conn.open("Base_Encuestas_SRPA")
        except gspread.SpreadsheetNotFound:
            sheet = conn.create("Base_Encuestas_SRPA")
            # Compartir con el propietario real si es necesario (el correo de la cuenta lo tiene por defecto)
        
        # Pestaña 1: Cola_Revision
        try:
            sheet.worksheet("Cola_Revision")
        except gspread.WorksheetNotFound:
            worksheet = sheet.add_worksheet(title="Cola_Revision", rows="1000", cols="10")
            headers = ['ID_Encuesta', 'Fecha_Carga', 'Tipo_Formulario', 'Municipio', 'Institucion_Educativa_IA', 'Rol', 'JSON_Respuestas', 'Estado']
            worksheet.append_row(headers)
            
        # Pestaña 2: Respuestas_SRPA
        try:
            sheet.worksheet("Respuestas_SRPA")
        except gspread.WorksheetNotFound:
            worksheet = sheet.add_worksheet(title="Respuestas_SRPA", rows="5000", cols="30")
            headers = [
                'ID_Encuesta', 'Tipo_Formulario', 'Fecha', 'Municipio', 
                'Institucion_Educativa_Verificada', 'Rol', 
                'Conocimientos_P1', 'Conocimientos_P2', 'Conocimientos_P3', 'Conocimientos_P4', 
                'Conocimientos_P5', 'Conocimientos_P6', 'Conocimientos_P7', 'Conocimientos_P8', 
                'Sat_P1', 'Sat_P2', 'Sat_P3', 'Sat_P4', 'Sat_P5', 'Sat_P6', 'Sat_P7', 'Sat_P8', 'Sat_P9', 
                'Verificado_Por', 'Fecha_Aprobacion'
            ]
            worksheet.append_row(headers)
            
        return sheet
    except Exception as e:
        st.error(f"Error inicializando Google Sheet: {e}")
        return None

# --- CARGAR COLA DE REVISIÓN (DEFENSIVA PARA EVITAR KEYERRORS) ---
def cargar_cola_revision():
    conn = conectar_google_sheets()
    if conn is None:
        return pd.DataFrame()
    try:
        sheet = conn.open("Base_Encuestas_SRPA")
        worksheet = sheet.worksheet("Cola_Revision")
        records = worksheet.get_all_records()
        
        columns_cola = ['ID_Encuesta', 'Fecha_Carga', 'Tipo_Formulario', 'Municipio', 'Institucion_Educativa_IA', 'Rol', 'JSON_Respuestas', 'Estado']
        if not records:
            df = pd.DataFrame(columns=columns_cola)
        else:
            df = pd.DataFrame(records)
            for col in columns_cola:
                if col not in df.columns:
                    df[col] = "" # Inicializar columna faltante
        return df
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame()

# --- CARGAR RESPUESTAS VALIDADAS (DEFENSIVA PARA EVITAR KEYERRORS) ---
def cargar_respuestas_validadas():
    conn = conectar_google_sheets()
    if conn is None:
        return pd.DataFrame()
    try:
        sheet = conn.open("Base_Encuestas_SRPA")
        worksheet = sheet.worksheet("Respuestas_SRPA")
        records = worksheet.get_all_records()
        
        columns_resp = [
            'ID_Encuesta', 'Tipo_Formulario', 'Fecha', 'Municipio', 
            'Institucion_Educativa_Verificada', 'Rol', 
            'Conocimientos_P1', 'Conocimientos_P2', 'Conocimientos_P3', 'Conocimientos_P4', 
            'Conocimientos_P5', 'Conocimientos_P6', 'Conocimientos_P7', 'Conocimientos_P8', 
            'Sat_P1', 'Sat_P2', 'Sat_P3', 'Sat_P4', 'Sat_P5', 'Sat_P6', 'Sat_P7', 'Sat_P8', 'Sat_P9', 
            'Verificado_Por', 'Fecha_Aprobacion'
        ]
        if not records:
            df = pd.DataFrame(columns=columns_resp)
        else:
            df = pd.DataFrame(records)
            for col in columns_resp:
                if col not in df.columns:
                    df[col] = ""
        return df
    except Exception as e:
        st.error(f"Error cargando Respuestas SRPA: {e}")
        return pd.DataFrame()

# --- LLAMADA COMPATIBLE CON GEMINI (MÚLTIPLES VERSIONES DEL SDK) ---
def analizar_imagenes_con_gemini(api_key, img1_bytes, img2_bytes, prompt):
    """Envía de forma robusta dos imágenes y un prompt a Gemini."""
    if USE_NEW_SDK:
        try:
            client = genai.Client(api_key=api_key)
            img1_part = types.Part.from_bytes(data=img1_bytes, mime_type="image/jpeg")
            img2_part = types.Part.from_bytes(data=img2_bytes, mime_type="image/jpeg")
            
            response = client.models.generate_content(
                model='gemini-1.5-flash',
                contents=[img1_part, img2_part, prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            return response.text
        except Exception as e:
            # Fallback en caso de error del SDK nuevo
            pass
            
    # Configuración por SDK heredado
    import google.generativeai as legacy_genai
    legacy_genai.configure(api_key=api_key)
    model = legacy_genai.GenerativeModel('gemini-1.5-flash')
    
    img1 = Image.open(io.BytesIO(img1_bytes))
    img2 = Image.open(io.BytesIO(img2_bytes))
    
    response = model.generate_content(
        [img1, img2, prompt],
        generation_config=legacy_genai.types.GenerationConfig(
            response_mime_type="application/json",
            temperature=0.1
        )
    )
    return response.text

# --- PROGRAMA PRINCIPAL ---
st.title("Proyecto \"Construyendo Futuro\"")
st.caption("Sistematización Colaborativa con IA para Encuestas SRPA — Gobernación de Bolívar")

# Asegurar la creación e inicialización de la base de datos en Sheets
if GSPREAD_AVAILABLE and "gcp_service_account" in st.secrets:
    inicializar_hoja_y_pestañas()

tab_carga, tab_revision, tab_dashboard = st.tabs([
    "📸 Carga de Encuestas", 
    "✍️ Revisión de Caligrafía", 
    "📊 Dashboard Estadístico"
])

# ================= TAB: CARGA DE ENCUESTAS =================
with tab_carga:
    st.header("Sistematizar Cuestionario")
    st.write("Toma o sube las fotos de ambas páginas de la encuesta física para que la IA extraiga el contenido.")
    
    # Obtener clave API de los secretos
    api_key_disponible = "GEMINI_API_KEY" in st.secrets
    gemini_key = st.secrets.get("GEMINI_API_KEY", "") if api_key_disponible else ""

    if not api_key_disponible:
        gemini_key = st.text_input("Ingresa tu Gemini API Key (Capa gratuita en Google AI Studio):", type="password")
        
    col1, col2 = st.columns(2)
    with col1:
        img_pag1 = st.file_uploader("Subir foto: PÁGINA 1 (Cabecera e inicio)", type=["jpg", "jpeg", "png"], key="p1_loader")
    with col2:
        img_pag2 = st.file_uploader("Subir foto: PÁGINA 2 (Final de cuestionario)", type=["jpg", "jpeg", "png"], key="p2_loader")
        
    if img_pag1 and img_pag2:
        if st.button("🚀 Procesar Encuesta con IA"):
            if not gemini_key:
                st.error("Por favor ingresa o configura tu Gemini API Key.")
            else:
                with st.spinner("Procesando imágenes con Gemini 1.5 Flash..."):
                    try:
                        # Leer bytes de imágenes
                        bytes_p1 = img_pag1.read()
                        bytes_p2 = img_pag2.read()
                        
                        # Llamar a Gemini
                        json_result_text = analizar_imagenes_con_gemini(
                            api_key=gemini_key,
                            img1_bytes=bytes_p1,
                            img2_bytes=bytes_p2,
                            prompt=SYSTEM_PROMPT
                        )
                        
                        # Limpiar JSON por si el modelo devuelve markdown ticks
                        if "```json" in json_result_text:
                            json_result_text = json_result_text.split("```json")[1].split("```")[0].strip()
                        elif "```" in json_result_text:
                            json_result_text = json_result_text.split("```")[1].split("```")[0].strip()
                            
                        data = json.loads(json_result_text)
                        
                        # Guardar en Google Sheets (Cola_Revision)
                        conn = conectar_google_sheets()
                        if conn:
                            sheet = conn.open("Base_Encuestas_SRPA")
                            worksheet = sheet.worksheet("Cola_Revision")
                            
                            id_enc = f"SRPA_{int(datetime.datetime.now().timestamp())}"
                            fecha_carga = datetime.date.today().strftime("%Y-%m-%d")
                            enc = data.get("encabezado", {})
                            
                            row = [
                                id_enc,
                                fecha_carga,
                                data.get("tipo_formulario", "PRETEST"),
                                enc.get("municipio", ""),
                                enc.get("institucion_educativa_ia", ""),
                                enc.get("rol", ""),
                                json.dumps(data),
                                "Pendiente"
                            ]
                            worksheet.append_row(row)
                            st.success(f"¡Cuestionario procesado con éxito! Se guardó provisionalmente con el ID: **{id_enc}**. Pasa a la pestaña de Revisión para verificar la Institución.")
                            st.rerun()
                        else:
                            st.error("No se pudo conectar con Google Sheets para almacenar los resultados.")
                    except Exception as e:
                        st.error(f"Error procesando la encuesta: {e}")

# ================= TAB: REVISIÓN DE CALIGRAFÍA =================
with tab_revision:
    st.header("Banco de Verificación de Escritura a Mano")
    st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
    
    df_cola = cargar_cola_revision()
    
    if df_cola.empty:
        st.info("No hay encuestas pendientes de revisar en este momento.")
    else:
        df_pendiente = df_cola[df_cola['Estado'] == 'Pendiente']
        
        if df_pendiente.empty:
            st.success("🎉 ¡Felicidades! Todas las encuestas pendientes han sido revisadas y validadas.")
        else:
            st.warning(f"Tienes **{len(df_pendiente)}** encuestas pendientes por revisar.")
            
            # Selector de encuesta pendiente
            opciones_pendientes = df_pendiente.apply(
                lambda r: f"{r['ID_Encuesta']} - {r['Tipo_Formulario']} ({r['Municipio']} - {r['Fecha_Carga']})", axis=1
            ).tolist()
            
            seleccion = st.selectbox("Selecciona la encuesta para validar:", opciones_pendientes)
            
            if seleccion:
                id_seleccionado = seleccion.split(" - ")[0]
                registro = df_pendiente[df_pendiente['ID_Encuesta'] == id_seleccionado].iloc[0]
                
                # Cargar el JSON extraído por la IA
                try:
                    datos_ia = json.loads(registro['JSON_Respuestas'])
                except Exception:
                    datos_ia = {}
                    
                enc = datos_ia.get("encabezado", {})
                
                st.markdown("### Datos Extraídos preliminarmente")
                
                # Cuadro de edición para la caligrafía manuscrita de la escuela
                escuela_ia = enc.get("institucion_educativa_ia", "")
                
                st.markdown("#### 📝 Corrección de Institución Educativa:")
                institucion_corregida = st.text_input(
                    "Nombre de la Institución Educativa (Corrígelo de ser necesario):", 
                    value=escuela_ia
                )
                
                # Mostrar el resto de metadatos
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.text_input("Municipio:", value=enc.get("municipio", ""), disabled=True)
                with col2:
                    st.text_input("Rol de Participante:", value=enc.get("rol", ""), disabled=True)
                with col3:
                    st.text_input("Tipo de Encuesta:", value=datos_ia.get("tipo_formulario", ""), disabled=True)
                
                col_btn1, col_btn2 = st.columns(2)
                
                # ACCIÓN 1: APROBAR E INGRESAR
                with col_btn1:
                    if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True):
                        conn = conectar_google_sheets()
                        if conn:
                            try:
                                sheet = conn.open("Base_Encuestas_SRPA")
                                
                                # 1. Añadir fila consolidada a Respuestas_SRPA
                                worksheet_respuestas = sheet.worksheet("Respuestas_SRPA")
                                
                                # Mapear respuestas de conocimiento
                                resp_con = datos_ia.get("conocimientos", {})
                                p1 = resp_con.get("p1", "")
                                p2 = resp_con.get("p2", "")
                                p3 = resp_con.get("p3", "")
                                p4 = resp_con.get("p4", "")
                                p5 = resp_con.get("p5", "")
                                p6 = resp_con.get("p6", "")
                                p7 = resp_con.get("p7", "")
                                p8 = resp_con.get("p8", "") if datos_ia.get("tipo_formulario") == "PRETEST" else ""
                                
                                # Mapear respuestas de satisfacción
                                resp_sat = datos_ia.get("satisfaccion", {}) or {}
                                s1 = resp_sat.get("sat_1", "")
                                s2 = resp_sat.get("sat_2", "")
                                s3 = resp_sat.get("sat_3", "")
                                s4 = resp_sat.get("sat_4", "")
                                s5 = resp_sat.get("sat_5", "")
                                s6 = resp_sat.get("sat_6", "")
                                s7 = resp_sat.get("sat_7", "")
                                s8 = resp_sat.get("sat_8", "")
                                s9 = resp_sat.get("sat_9", "")
                                
                                # Armar fila final
                                row_final = [
                                    id_seleccionado,
                                    datos_ia.get("tipo_formulario", "PRETEST"),
                                    enc.get("fecha", ""),
                                    enc.get("municipio", ""),
                                    institucion_corregida, # Escuela corregida por el usuario
                                    enc.get("rol", ""),
                                    p1, p2, p3, p4, p5, p6, p7, p8,
                                    s1, s2, s3, s4, s5, s6, s7, s8, s9,
                                    "Facilitador Campo",
                                    datetime.date.today().strftime("%Y-%m-%d")
                                ]
                                worksheet_respuestas.append_row(row_final)
                                
                                # 2. Cambiar estado en Cola_Revision a Aprobado
                                worksheet_cola = sheet.worksheet("Cola_Revision")
                                cell = worksheet_cola.find(id_seleccionado)
                                if cell:
                                    # La columna Estado es la columna 8
                                    worksheet_cola.update_cell(cell.row, 8, "Aprobado")
                                    
                                st.success("¡Registro consolidado de forma definitiva!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error consolidando el registro: {e}")
                
                # ACCIÓN 2: RECHAZAR ENTRADA
                with col_btn2:
                    if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                        conn = conectar_google_sheets()
                        if conn:
                            try:
                                sheet = conn.open("Base_Encuestas_SRPA")
                                worksheet_cola = sheet.worksheet("Cola_Revision")
                                cell = worksheet_cola.find(id_seleccionado)
                                if cell:
                                    # La columna Estado es la columna 8. Lo marcamos como "Rechazado" para no borrar
                                    # físicamente el log, o podemos eliminar la fila si se desea. Marcamos "Rechazado"
                                    worksheet_cola.update_cell(cell.row, 8, "Rechazado")
                                st.warning("Registro descartado y eliminado de la cola de pendientes.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error rechazando el registro: {e}")

# ================= TAB: DASHBOARD ESTADÍSTICO =================
with tab_dashboard:
    st.header("Resultados de Impacto y Satisfacción en Tiempo Real")
    st.write("Estadísticas consolidadas directamente de las encuestas validadas.")
    
    df_validadas = cargar_respuestas_validadas()
    
    if df_validadas.empty:
        st.info("No hay datos consolidados en la base de datos de Google Sheets para generar estadísticas aún. Empieza a validar encuestas para ver los gráficos interactivos.")
    else:
        # ---- FILTROS DEL DASHBOARD ----
        st.markdown("### 🔍 Filtrar Visualizaciones")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            municipios_disponibles = ["Todos"] + df_validadas["Municipio"].dropna().unique().tolist()
            filtro_muni = st.selectbox("Municipio:", municipios_disponibles)
        with col_f2:
            roles_disponibles = ["Todos"] + df_validadas["Rol"].dropna().unique().tolist()
            filtro_rol = st.selectbox("Rol del Participante:", roles_disponibles)
            
        # Filtrar DataFrame
        df_filtrado = df_validadas.copy()
        if filtro_muni != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Municipio"] == filtro_muni]
        if filtro_rol != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Rol"] == filtro_rol]
            
        # ---- MÓDULO DE KPIs ----
        col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
        
        with col_kpi1:
            st.markdown(f"""
            <div class="kpi-card">
                <span style="font-size: 28px; font-weight: bold; color: #1e3a8a;">{len(df_filtrado)}</span><br>
                <span style="font-size: 14px; color: #475569; font-weight: 500;">Total Consolidadas</span>
            </div>
            """, unsafe_allowed_value=True)
            
        with col_kpi2:
            df_cola_activa = cargar_cola_revision()
            total_cola = len(df_cola_activa[df_cola_activa["Estado"] == "Pendiente"]) if not df_cola_activa.empty else 0
            st.markdown(f"""
            <div class="kpi-card">
                <span style="font-size: 28px; font-weight: bold; color: #b45309;">{total_cola}</span><br>
                <span style="font-size: 14px; color: #475569; font-weight: 500;">En Cola de Revisión</span>
            </div>
            """, unsafe_allowed_value=True)
            
        with col_kpi3:
            total_pre = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"])
            st.markdown(f"""
            <div class="kpi-card">
                <span style="font-size: 28px; font-weight: bold; color: #047857;">{total_pre}</span><br>
                <span style="font-size: 14px; color: #475569; font-weight: 500;">Pretests Registrados</span>
            </div>
            """, unsafe_allowed_value=True)
            
        with col_kpi4:
            total_post = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"])
            st.markdown(f"""
            <div class="kpi-card">
                <span style="font-size: 28px; font-weight: bold; color: #0284c7;">{total_post}</span><br>
                <span style="font-size: 14px; color: #475569; font-weight: 500;">Postests Registrados</span>
            </div>
            """, unsafe_allowed_value=True)
            
        st.markdown("---")
        
        # Gráficos si hay datos filtrados
        if df_filtrado.empty:
            st.warning("No hay registros que coincidan con los filtros seleccionados.")
        else:
            col_chart1, col_chart2 = st.columns(2)
            
            # 1. PARTICIPACIÓN POR ROL
            with col_chart1:
                st.subheader("👥 Distribución de Participantes")
                conteo_roles = df_filtrado["Rol"].value_counts().reset_index()
                conteo_roles.columns = ["Rol", "Cantidad"]
                fig_roles = px.pie(
                    conteo_roles, 
                    values="Cantidad", 
                    names="Rol", 
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Safe
                )
                fig_roles.update_layout(margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_roles, use_container_width=True)
                
            # 2. MEDICIÓN ANTES Y DESPUÉS (RESPUESTAS CORRECTAS CLAVE)
            with col_chart2:
                st.subheader("📚 Comparativa de Conocimiento (Antes vs Después)")
                
                # Filtrar Pretest y Postest para comparar
                df_pre = df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"]
                df_post = df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"]
                
                # Definimos conceptos clave y mapeamos respuestas correctas
                # P2 Pretest == P1 Postest (Finalidad del SRPA) -> Correcta: 'b'
                # P3 Pretest == P2 Postest (Factor de Riesgo) -> Correcta: 'b'
                # P4 Pretest == P3 Postest (Factor Protector) -> Correcta: 'a' en Pre, 'b' en Post (según cuestionarios oficiales)
                # P5 Pretest == P4 Postest (Responsabilidad de Prevención) -> Correcta: 'd'
                
                conceptos = ["Finalidad SRPA", "Factor Riesgo", "Factor Protector", "Responsabilidad"]
                correctas_pre = [0, 0, 0, 0]
                correctas_post = [0, 0, 0, 0]
                
                # Calcular porcentajes si hay registros
                if len(df_pre) > 0:
                    correctas_pre[0] = (df_pre["Conocimientos_P2"] == "b").mean() * 100
                    correctas_pre[1] = (df_pre["Conocimientos_P3"] == "b").mean() * 100
                    correctas_pre[2] = (df_pre["Conocimientos_P4"] == "a").mean() * 100
                    correctas_pre[3] = (df_pre["Conocimientos_P5"] == "d").mean() * 100
                    
                if len(df_post) > 0:
                    correctas_post[0] = (df_post["Conocimientos_P1"] == "b").mean() * 100
                    correctas_post[1] = (df_post["Conocimientos_P2"] == "b").mean() * 100
                    correctas_post[2] = (df_post["Conocimientos_P3"] == "b").mean() * 100
                    correctas_post[3] = (df_post["Conocimientos_P4"] == "d").mean() * 100
                
                fig_comp = go.Figure(data=[
                    go.Bar(name='Antes (Pretest)', x=conceptos, y=correctas_pre, marker_color='#ef4444'),
                    go.Bar(name='Después (Postest)', x=conceptos, y=correctas_post, marker_color='#22c55e')
                ])
                fig_comp.update_layout(
                    barmode='group',
                    yaxis_title='Respuestas Correctas (%)',
                    yaxis_range=[0, 100],
                    margin=dict(l=20, r=20, t=20, b=20)
                )
                st.plotly_chart(fig_comp, use_container_width=True)

            # 3. EVALUACIÓN DE SATISFACCIÓN (SÓLO POSTEST)
            df_post_satisfaccion = df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"]
            if not df_post_satisfaccion.empty:
                st.subheader("⭐ Satisfacción de las Jornadas de Prevención")
                
                sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
                aspectos = [
                    "1. Claridad de la información",
                    "2. Dominio de facilitadores",
                    "3. Metodología utilizada",
                    "4. Participación e integración",
                    "5. Utilidad de los temas",
                    "6. Organización de la jornada",
                    "7. Materiales y recursos",
                    "8. Fortaleció conocimientos",
                    "9. Recomendaría la jornada"
                ]
                
                # Mapear opiniones de satisfacción a puntajes numéricos
                # Excelente = 4, Bueno = 3, Regular = 2, Deficiente = 1
                map_sat = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None}
                
                puntajes_aspectos = []
                for col in sat_cols:
                    if col in df_post_satisfaccion.columns:
                        puntos = df_post_satisfaccion[col].map(map_sat).dropna()
                        promedio = puntos.mean() if not puntos.empty else 0
                        puntajes_aspectos.append(promedio)
                    else:
                        puntajes_aspectos.append(0)
                
                fig_sat = px.bar(
                    x=puntajes_aspectos,
                    y=aspectos,
                    orientation='h',
                    labels={'x': 'Puntuación Promedio (Máx 4.0)', 'y': 'Aspecto Evaluado'},
                    color=puntajes_aspectos,
                    color_continuous_scale='Greens',
                    range_x=[1.0, 4.0]
                )
                fig_sat.update_layout(coloraxis_showscale=False, margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_sat, use_container_width=True)
            else:
                st.info("ℹ️ Para ver las métricas de satisfacción y el desglose de calidad de las jornadas, aprueba al menos una encuesta de tipo **POSTEST**.")

