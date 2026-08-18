import streamlit as st
import pandas as pd
import json
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image
import io
import base64
import google.generativeai as genai
import os
import plotly.express as px

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo personalizado para móviles y marca institucional (Gobernación de Bolívar)
st.markdown("""
<style>
    .main-header {
        color: #0F2C59;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-weight: bold;
        text-align: center;
        margin-bottom: 5px;
    }
    .subheader {
        color: #DAC0A3;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        text-align: center;
        font-size: 1.1rem;
        margin-bottom: 25px;
    }
    .card {
        background-color: #F8F0E5;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #EADBC8;
        margin-bottom: 15px;
    }
    .stButton>button {
        background-color: #0F2C59;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        border: none;
        padding: 10px 20px;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #1A4D80;
        color: white;
    }
    .reject-button>button {
        background-color: #D9534F !important;
        color: white !important;
    }
    .reject-button>button:hover {
        background-color: #C9302C !important;
    }
    .approve-button>button {
        background-color: #5CB85C !important;
        color: white !important;
    }
    .approve-button>button:hover {
        background-color: #4CAE4C !important;
    }
</style>
""", unsafe_allow_html=True)

# --- CONEXIÓN DE GOOGLE SHEETS ---
def obtener_credenciales_gcp():
    """Obtiene y formatea las credenciales de la cuenta de servicio desde st.secrets."""
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    # Limpieza de saltos de línea en la llave privada para evitar fallos de formato
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return creds_dict

def conectar_google_sheets():
    """Conecta de forma segura con la hoja de cálculo en Drive."""
    creds_dict = obtener_credenciales_gcp()
    if creds_dict is None:
        return None
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de conexión con Google API: {e}")
        return None

def obtener_o_crear_spreadsheet():
    """Busca la hoja central 'Base_Encuestas_SRPA' o la crea e inicializa si no existe."""
    client = conectar_google_sheets()
    if client is None:
        return None
    
    nombre_doc = "Base_Encuestas_SRPA"
    try:
        sh = client.open(nombre_doc)
        return sh
    except gspread.exceptions.SpreadsheetNotFound:
        # Si no existe, la creamos desde cero
        try:
            sh = client.create(nombre_doc)
            
            # Inicializar pestaña 1: Cola_Revision
            ws_cola = sh.get_worksheet(0)
            ws_cola.update_title("Cola_Revision")
            headers_cola = [
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", 
                "Foto_P1_Base64", "Foto_P2_Base64", "Estado", "Usuario_Cargo"
            ]
            ws_cola.append_row(headers_cola)
            
            # Inicializar pestaña 2: Respuestas_SRPA
            ws_resp = sh.add_worksheet(title="Respuestas_SRPA", rows=1000, cols=30)
            headers_resp = [
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
                "Institucion_Educativa_Verificada", "Rol", 
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
                "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                "Verificado_Por", "Fecha_Aprobacion"
            ]
            ws_resp.append_row(headers_resp)
            return sh
        except Exception as e:
            st.error(f"No se pudo crear la hoja 'Base_Encuestas_SRPA' automáticamente: {e}")
            return None

def cargar_datos_hoja(worksheet_name, default_cols):
    """Carga los datos de una pestaña de forma segura mapeando dinámicamente las columnas."""
    sh = obtener_o_crear_spreadsheet()
    if sh is None:
        return pd.DataFrame(columns=default_cols)
    
    try:
        ws = sh.worksheet(worksheet_name)
        values = ws.get_all_values()
        
        if not values or len(values) < 1:
            # Si está completamente vacía, escribir cabeceras por defecto
            ws.append_row(default_cols)
            return pd.DataFrame(columns=default_cols)
            
        headers = [h.strip() for h in values[0]]
        rows = values[1:]
        
        # Mapeo y alineación defensiva de columnas para evitar KeyErrors
        cleaned_rows = []
        for r in rows:
            if len(r) < len(headers):
                r = r + [''] * (len(headers) - len(r))
            elif len(r) > len(headers):
                r = r[:len(headers)]
            cleaned_rows.append(r)
            
        df = pd.DataFrame(cleaned_rows, columns=headers)
        return df
    except Exception as e:
        st.error(f"Error al cargar la pestaña {worksheet_name}: {e}")
        return pd.DataFrame(columns=default_cols)

# Columnas oficiales por defecto
COLUMNAS_COLA = [
    "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
    "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", 
    "Foto_P1_Base64", "Foto_P2_Base64", "Estado", "Usuario_Cargo"
]

COLUMNAS_RESPUESTAS = [
    "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
    "Institucion_Educativa_Verificada", "Rol", 
    "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
    "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
    "Verificado_Por", "Fecha_Aprobacion"
]

def guardar_en_cola_sheets(row_dict):
    """Guarda un registro mapeándolo dinámicamente según las cabeceras reales de la hoja."""
    sh = obtener_o_crear_spreadsheet()
    if sh is None:
        st.error("No se pudo conectar a Google Sheets para guardar.")
        return False
    
    try:
        ws = sh.worksheet("Cola_Revision")
        values = ws.get_all_values()
        if not values:
            ws.append_row(COLUMNAS_COLA)
            headers = COLUMNAS_COLA
        else:
            headers = [h.strip() for h in values[0]]
            
        # Mapear diccionarios a los índices exactos de la cabecera real de la hoja
        row_to_append = []
        for h in headers:
            # Buscar coincidencia sin distinguir espacios ni mayúsculas/minúsculas
            matching_val = ""
            h_clean = h.lower().replace("_", "").replace(" ", "")
            for k, v in row_dict.items():
                k_clean = k.lower().replace("_", "").replace(" ", "")
                if k_clean == h_clean:
                    matching_val = v
                    break
            row_to_append.append(str(matching_val) if matching_val is not None else "")
            
        ws.append_row(row_to_append)
        return True
    except Exception as e:
        st.error(f"Error al escribir en la Cola de Revisión: {e}")
        return False

def guardar_aprobado_sheets(row_dict, id_a_eliminar):
    """Guarda en Respuestas_SRPA y elimina/actualiza en Cola_Revision con alineación dinámica."""
    sh = obtener_o_crear_spreadsheet()
    if sh is None:
        st.error("Error de conexión al aprobar.")
        return False
    
    try:
        # 1. Guardar en Respuestas_SRPA
        ws_resp = sh.worksheet("Respuestas_SRPA")
        headers_resp = [h.strip() for h in ws_resp.get_all_values()[0]]
        
        row_resp = []
        for h in headers_resp:
            matching_val = ""
            h_clean = h.lower().replace("_", "").replace(" ", "")
            for k, v in row_dict.items():
                k_clean = k.lower().replace("_", "").replace(" ", "")
                if k_clean == h_clean:
                    matching_val = v
                    break
            row_resp.append(str(matching_val) if matching_val is not None else "")
            
        ws_resp.append_row(row_resp)
        
        # 2. Actualizar estado a 'Aprobado' o eliminar en Cola_Revision
        ws_cola = sh.worksheet("Cola_Revision")
        values_cola = ws_cola.get_all_values()
        headers_cola = [h.strip() for h in values_cola[0]]
        
        # Buscar el ID de columna 'ID_Encuesta' y 'Estado'
        idx_id = -1
        idx_estado = -1
        for i, h in enumerate(headers_cola):
            h_clean = h.lower().replace("_", "").replace(" ", "")
            if h_clean == "idencuesta":
                idx_id = i
            elif h_clean == "estado":
                idx_estado = i
                
        if idx_id != -1 and idx_estado != -1:
            for row_num, r in enumerate(values_cola[1:], start=2):
                if len(r) > idx_id and r[idx_id] == id_a_eliminar:
                    # Cambiar estado a Aprobado en la hoja física
                    ws_cola.update_cell(row_num, idx_estado + 1, "Aprobado")
                    break
        return True
    except Exception as e:
        st.error(f"Error al aprobar el registro: {e}")
        return False

def rechazar_en_cola_sheets(id_a_eliminar):
    """Marca un registro como 'Rechazado' en la hoja de cálculo Cola_Revision."""
    sh = obtener_o_crear_spreadsheet()
    if sh is None:
        return False
    try:
        ws_cola = sh.worksheet("Cola_Revision")
        values_cola = ws_cola.get_all_values()
        headers_cola = [h.strip() for h in values_cola[0]]
        
        idx_id = -1
        idx_estado = -1
        for i, h in enumerate(headers_cola):
            h_clean = h.lower().replace("_", "").replace(" ", "")
            if h_clean == "idencuesta":
                idx_id = i
            elif h_clean == "estado":
                idx_estado = i
                
        if idx_id != -1 and idx_estado != -1:
            for row_num, r in enumerate(values_cola[1:], start=2):
                if len(r) > idx_id and r[idx_id] == id_a_eliminar:
                    ws_cola.update_cell(row_num, idx_estado + 1, "Rechazado")
                    return True
        return False
    except Exception as e:
        st.error(f"Error al rechazar registro: {e}")
        return False

# --- MOTOR DE INTELIGENCIA ARTIFICIAL GEMINI ---
def analizar_con_gemini(img1, img2, system_prompt):
    """Llama a la API de Google Gemini utilizando un sistema de redundancia ante deprecación."""
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        raise ValueError("GEMINI_API_KEY no encontrada. Configúrala en los Secrets de Streamlit.")
        
    genai.configure(api_key=api_key)
    
    # Lista de modelos ordenada de más nuevo a más antiguo para evitar errores 404
    modelos = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.0-flash"]
    
    # Preparación de imágenes en bytes
    def pil_to_bytes(img):
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        return buf.getvalue()
        
    contents = [
        {"mime_type": "image/jpeg", "data": pil_to_bytes(img1)},
        {"mime_type": "image/jpeg", "data": pil_to_bytes(img2)},
        system_prompt
    ]
    
    last_err = None
    for m in modelos:
        try:
            model = genai.GenerativeModel(m)
            response = model.generate_content(contents)
            return response.text
        except Exception as e:
            last_err = e
            continue
    raise last_err

SYSTEM_PROMPT_OCR = """
Eres un analizador OCR especializado en las encuestas de conocimientos del Proyecto "Construyendo Futuro" (SRPA) de la Gobernación de Bolívar.
Analiza las dos imágenes proporcionadas (Página 1 y Página 2 de un mismo cuestionario) y extrae de forma extremadamente estricta la información en un formato JSON plano estructurado.

Reglas de Oro de Extracción:
1. Identifica si el cuestionario es "PRETEST" o "POSTEST" (se indica en el título superior).
2. Extrae los metadatos escritos a mano en la cabecera de la Página 1:
   - "fecha": Extrae la fecha escrita a mano.
   - "municipio": Extrae el municipio (por ejemplo, Cartagena).
   - "institucion_educativa": Lee la caligrafía manuscrita del nombre de la escuela (ej. "Promesa de Dios"). Si está vacío, pon "".
   - "rol": Selecciona el rol marcado con (X): Estudiante, Docente, Padre de Familia, o Lider comunitario.
3. Lee las respuestas a las preguntas de conocimiento de selección única (marcadas con "X" u otra seña):
   - En el PRETEST son 8 preguntas.
   - En el POSTEST son 5 preguntas (Sección A).
   Asigna valores como "a", "b", "c", "d" de acuerdo con la opción elegida. Si la opción está completamente en blanco (el adolescente no marcó nada), debes dejar el campo estrictamente vacío como "". ¡No inventes ni auto-rellenes información!
4. Lee la matriz de satisfacción de la Sección B si es un POSTEST (9 filas con Excelente, Bueno, Regular, Deficiente). Si no hay marcas, pon "".

Formato de salida estricto esperado (JSON):
{
  "tipo_formulario": "PRETEST",
  "fecha": "30/07/2026",
  "municipio": "Cartagena",
  "institucion_educativa": "Promesa de Dios",
  "rol": "Estudiante",
  "conocimientos": {
    "p1": "b", "p2": "b", "p3": "b", "p4": "a", "p5": "d", "p6": "a", "p7": "a", "p8": "c"
  },
  "satisfaccion": {
    "sat_p1": "", "sat_p2": "", "sat_p3": "", "sat_p4": "", "sat_p5": "", "sat_p6": "", "sat_p7": "", "sat_p8": "", "sat_p9": ""
  }
}
"""

# --- COMPRESIÓN DE IMÁGENES ---
def comprimir_a_base64(img):
    """Comprime una imagen PIL y la convierte a string base64 para guardado seguro en la nube."""
    img_copy = img.copy()
    img_copy.thumbnail((350, 350)) # Resolución óptima para móvil
    buf = io.BytesIO()
    img_copy.save(buf, format="JPEG", quality=75)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def mostrar_imagen_base64(base64_str):
    """Decodifica un string base64 y lo muestra en Streamlit."""
    try:
        img_bytes = base64.b64decode(base64_str)
        img = Image.open(io.BytesIO(img_bytes))
        return img
    except Exception:
        return None

# --- INTERFAZ GRÁFICA DE USUARIO ---

# Encabezado institucional
st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
st.markdown("<p class='subheader'>Sistematización de Evaluaciones de Conocimientos SRPA - Bolívar</p>", unsafe_allow_html=True)

# Validación de Credenciales
credenciales_listas = obtener_credenciales_gcp() is not None

if not credenciales_listas:
    st.info("👋 ¡Bienvenido! Para comenzar, conecta tu base de datos de Google Sheets de la Gobernación.")
    with st.expander("🛠️ Guía Rápida de Conexión de Google Sheets (Secrets)"):
        st.markdown("""
        Para guardar colaborativamente los datos de campo, ingresa tu cuenta de servicio de Google Cloud en los Secretos de Streamlit Cloud:
        1. Crea un proyecto en Google Cloud Console, activa **Google Sheets API** y **Google Drive API**.
        2. Crea una **Cuenta de Servicio**, descarga el archivo de claves **JSON** y compártelo como editor con tu Google Sheet `Base_Encuestas_SRPA`.
        3. En Streamlit Cloud, ve a **Settings** -> **Secrets** y pega tus credenciales con esta estructura:
        ```toml
        GEMINI_API_KEY = "tu_api_key_aquí"

        [gcp_service_account]
        type = "service_account"
        project_id = "tu-proyecto"
        private_key_id = "tu-key-id"
        private_key = "-----BEGIN PRIVATE KEY-----\\nTU_LLAVE_AQUI\\n-----END PRIVATE KEY-----\\n"
        client_email = "tu-correo-servicio@gserviceaccount.com"
        client_id = "..."
        auth_uri = "https://accounts.google.com/o/oauth2/auth"
        token_uri = "https://oauth2.google.com/token"
        ```
        """)
    st.stop()

# Menú lateral
st.sidebar.image("https://www.bolivar.gov.co/images/Gobernacion/Logo-Gobernacion-Color.png", use_container_width=True)
st.sidebar.markdown("---")
modulo = st.sidebar.radio("Navegación", ["📷 Cargar Encuestas", "🔍 Banco de Verificación", "📊 Dashboard Estadístico"])

# --- MÓDULO 1: CARGAR ENCUESTAS ---
if modulo == "📷 Cargar Encuestas":
    st.subheader("Subir Cuestionario de Doble Página")
    st.write("Sube las dos fotos de la encuesta física (Página 1 y Página 2) del mismo participante. La IA las procesará conjuntamente de forma anónima.")
    
    col1, col2 = st.columns(2)
    with col1:
        p1_file = st.file_uploader("Subir Página 1 (Cabecera y Preguntas 1-3)", type=["jpg", "png", "jpeg"], key="p1")
    with col2:
        p2_file = st.file_uploader("Subir Página 2 (Preguntas de Conocimiento y Satisfacción)", type=["jpg", "png", "jpeg"], key="p2")
        
    if p1_file and p2_file:
        img1 = Image.open(p1_file)
        img2 = Image.open(p2_file)
        
        # Mostrar previsualización
        with col1:
            st.image(img1, caption="Página 1 Cargada", use_container_width=True)
        with col2:
            st.image(img2, caption="Página 2 Cargada", use_container_width=True)
            
        if st.button("🚀 Procesar Encuesta con IA (Gemini 2.0/3.6)"):
            with st.spinner("La Inteligencia Artificial está leyendo y extrayendo las marcas del cuestionario..."):
                try:
                    # Llamar al procesador inteligente
                    resultado_json_str = analizar_con_gemini(img1, img2, SYSTEM_PROMPT_OCR)
                    
                    # Limpieza por seguridad de marcas de código markdown ```json
                    if "```json" in resultado_json_str:
                        resultado_json_str = resultado_json_str.split("```json")[1].split("```")[0].strip()
                    elif "```" in resultado_json_str:
                        resultado_json_str = resultado_json_str.split("```")[1].split("```")[0].strip()
                        
                    datos_extraidos = json.loads(resultado_json_str.strip())
                    
                    # Generar ID único y comprimir fotos
                    id_encuesta = f"ENC_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    p1_b64 = comprimir_a_base64(img1)
                    p2_b64 = comprimir_a_base64(img2)
                    
                    # Construir registro exacto para Google Sheets
                    registro_cola = {
                        "ID_Encuesta": id_encuesta,
                        "Fecha_Carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Formulario": datos_extraidos.get("tipo_formulario", "PRETEST"),
                        "Municipio": datos_extraidos.get("municipio", ""),
                        "Institucion_Educativa_IA": datos_extraidos.get("institucion_educativa", ""),
                        "Rol": datos_extraidos.get("rol", ""),
                        "JSON_Respuestas": json.dumps(datos_extraidos),
                        "Foto_P1_Base64": p1_b64,
                        "Foto_P2_Base64": p2_b64,
                        "Estado": "Pendiente",
                        "Usuario_Cargo": "Facilitador Campo"
                    }
                    
                    # Guardar con mapeo dinámico de columnas
                    guardado = guardar_en_cola_sheets(registro_cola)
                    if guardado:
                        st.success("🎉 ¡Encuesta procesada con éxito! Se ha añadido al Banco de Verificación de Caligrafía.")
                        st.balloons()
                except Exception as e:
                    st.error(f"Error procesando la encuesta: {e}")

# --- MÓDULO 2: BANCO DE VERIFICACIÓN ---
elif modulo == "🔍 Banco de Verificación":
    st.subheader("Banco de Verificación de Escritura a Mano")
    st.write("Verifica y corrige la caligrafía manuscrita de la **Institución Educativa** utilizando el panel de comparación visual.")
    
    # Cargar datos de la cola
    df_cola = cargar_datos_hoja("Cola_Revision", COLUMNAS_COLA)
    
    # ⚠️ REGLA DE ORO DE ALINEACIÓN DE COLUMNAS:
    # Si la tabla tiene datos pero las columnas no coinciden con COLUMNAS_COLA, buscar si 'Estado' está en la tabla
    columnas_tabla = list(df_cola.columns)
    col_estado = ""
    col_id = ""
    
    for c in columnas_tabla:
        c_clean = c.lower().replace("_", "").replace(" ", "")
        if c_clean == "estado":
            col_estado = c
        elif c_clean == "idencuesta":
            col_id = c
            
    # Si encontramos la columna estado de forma dinámica, la usamos para el filtro
    if col_estado and col_id:
        pendientes = df_cola[df_cola[col_estado].astype(str).str.strip().str.lower() == "pendiente"]
    else:
        pendientes = pd.DataFrame()
        
    if len(pendientes) == 0:
        st.success("✨ ¡Todo al día! No hay encuestas pendientes de verificación de caligrafía en este momento.")
    else:
        st.warning(f"Tienes {len(pendientes)} encuestas por verificar.")
        
        # Tomar la primera encuesta en cola
        encuesta_actual = pendientes.iloc[0]
        id_actual = encuesta_actual[col_id]
        
        # Mapear dinámicamente el resto de las celdas para evitar KeyErrors si las columnas están corridas
        col_fecha = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "fechacarga"), "Fecha_Carga")
        col_tipo = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "tipoformulario"), "Tipo_Formulario")
        col_muni = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "municipio"), "Municipio")
        col_ie = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "institucioneducativaia"), "Institucion_Educativa_IA")
        col_rol = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "rol"), "Rol")
        col_json = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "jsonrespuestas"), "JSON_Respuestas")
        col_img1 = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "fotop1base64"), "Foto_P1_Base64")
        col_img2 = next((c for c in columnas_tabla if c.lower().replace("_","").replace(" ","") == "fotop2base64"), "Foto_P2_Base64")
        
        # Extraer valores de forma segura
        tipo_form = encuesta_actual.get(col_tipo, "PRETEST")
        municipio = encuesta_actual.get(col_muni, "")
        ie_ia = encuesta_actual.get(col_ie, "")
        rol = encuesta_actual.get(col_rol, "")
        respuestas_json_str = encuesta_actual.get(col_json, "{}")
        
        try:
            respuestas_dict = json.loads(respuestas_json_str)
        except Exception:
            respuestas_dict = {}
            
        # PANEL VISUAL LADO A LADO (PÁGINA 1 Y PÁGINA 2)
        st.markdown("<div class='card'><h4>🖼️ Panel de Comparación Visual</h4></div>", unsafe_allow_html=True)
        col_v1, col_v2 = st.columns(2)
        
        with col_v1:
            st.markdown("**Foto Página 1 (Cabecera)**")
            if col_img1 in encuesta_actual and encuesta_actual[col_img1]:
                img_p1 = mostrar_imagen_base64(encuesta_actual[col_img1])
                if img_p1:
                    st.image(img_p1, use_container_width=True)
                else:
                    st.info("No se puede renderizar la previsualización de la Página 1.")
            else:
                st.info("Imagen de Página 1 no disponible.")
                
        with col_v2:
            st.markdown("**Foto Página 2 (Firmas y Respuestas)**")
            if col_img2 in encuesta_actual and encuesta_actual[col_img2]:
                img_p2 = mostrar_imagen_base64(encuesta_actual[col_img2])
                if img_p2:
                    st.image(img_p2, use_container_width=True)
                else:
                    st.info("No se puede renderizar la previsualización de la Página 2.")
            else:
                st.info("Imagen de Página 2 no disponible.")
                
        # FORMULARIO DE EDICIÓN Y VALIDACIÓN
        st.markdown("---")
        st.markdown("### ✏️ Verificar Datos del Formulario")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            municipio_verificado = st.text_input("Municipio:", value=str(municipio))
            rol_verificado = st.text_input("Rol de Participante:", value=str(rol))
        with col_f2:
            # Entrada crítica del nombre de la Institución Educativa manuscrita
            ie_verificada = st.text_input("🎒 Institución Educativa (Nombre Manuscrito):", value=str(ie_ia))
            tipo_verificado = st.text_input("Tipo de Cuestionario:", value=str(tipo_form), disabled=True)
            
        # BOTONES DE ACCIÓN SIMULTÁNEOS (APROBAR / RECHAZAR)
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.markdown('<div class="approve-button">', unsafe_allow_html=True)
            if st.button("✅ Aprobar e Ingresar a Base de Datos"):
                # Construir el registro de respuestas definitivo
                conocimientos = respuestas_dict.get("conocimientos", {})
                satisfaccion = respuestas_dict.get("satisfaccion", {})
                
                registro_aprobado = {
                    "ID_Encuesta": str(id_actual),
                    "Tipo_Formulario": str(tipo_form),
                    "Fecha": respuestas_dict.get("fecha", datetime.now().strftime("%d/%m/%Y")),
                    "Municipio": str(municipio_verificado),
                    "Institucion_Educativa_Verificada": str(ie_verificada),
                    "Rol": str(rol_verificado),
                    "Conocimientos_P1": conocimientos.get("p1", ""),
                    "Conocimientos_P2": conocimientos.get("p2", ""),
                    "Conocimientos_P3": conocimientos.get("p3", ""),
                    "Conocimientos_P4": conocimientos.get("p4", ""),
                    "Conocimientos_P5": conocimientos.get("p5", ""),
                    "Conocimientos_P6": conocimientos.get("p6", ""),
                    "Conocimientos_P7": conocimientos.get("p7", ""),
                    "Conocimientos_P8": conocimientos.get("p8", ""),
                    "Sat_P1": satisfaccion.get("sat_p1", ""),
                    "Sat_P2": satisfaccion.get("sat_p2", ""),
                    "Sat_P3": satisfaccion.get("sat_p3", ""),
                    "Sat_P4": satisfaccion.get("sat_p4", ""),
                    "Sat_P5": satisfaccion.get("sat_p5", ""),
                    "Sat_P6": satisfaccion.get("sat_p6", ""),
                    "Sat_P7": satisfaccion.get("sat_p7", ""),
                    "Sat_P8": satisfaccion.get("sat_p8", ""),
                    "Sat_P9": satisfaccion.get("sat_p9", ""),
                    "Verificado_Por": "Supervisor Bolívar",
                    "Fecha_Aprobacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                exito = guardar_aprobado_sheets(registro_aprobado, id_actual)
                if exito:
                    st.success("✅ ¡Registro consolidado correctamente en la pestaña Respuestas_SRPA!")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            
        with col_btn2:
            st.markdown('<div class="reject-button">', unsafe_allow_html=True)
            if st.button("❌ Rechazar Entrada (Eliminar de la Cola)"):
                exito = rechazar_en_cola_sheets(id_actual)
                if exito:
                    st.warning("❌ Entrada rechazada y eliminada de la cola de revisión activa.")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

# --- MÓDULO 3: DASHBOARD ESTADÍSTICO ---
elif modulo == "📊 Dashboard Estadístico":
    st.subheader("Dashboard de Impacto Educativo en Tiempo Real")
    
    # Cargar datos consolidados
    df_resp = cargar_datos_hoja("Respuestas_SRPA", COLUMNAS_RESPUESTAS)
    
    if len(df_resp) == 0:
        st.info("📈 Cuando apruebes los primeros registros en el Banco de Verificación, aquí verás gráficos de impacto y satisfacción en tiempo real.")
        st.stop()
        
    # KPI Cards superiores
    c_tot, c_pre, c_post = st.columns(3)
    c_tot.metric("Total Registros Consolidados", len(df_resp))
    c_pre.metric("Evaluaciones Pretest", len(df_resp[df_resp['Tipo_Formulario'] == "PRETEST"]))
    c_post.metric("Evaluaciones Postest", len(df_resp[df_resp['Tipo_Formulario'] == "POSTEST"]))
    
    # Filtros interactivos de visualización
    st.markdown("---")
    st.markdown("### Filtros Dinámicos")
    col_fil1, col_fil2 = st.columns(2)
    with col_fil1:
        ie_list = ["Todas"] + list(df_resp["Institucion_Educativa_Verificada"].unique())
        ie_filter = st.selectbox("Filtrar por Institución Educativa:", ie_list)
    with col_fil2:
        muni_list = ["Todos"] + list(df_resp["Municipio"].unique())
        muni_filter = st.selectbox("Filtrar por Municipio:", muni_list)
        
    df_filtered = df_resp.copy()
    if ie_filter != "Todas":
        df_filtered = df_filtered[df_filtered["Institucion_Educativa_Verificada"] == ie_filter]
    if muni_filter != "Todos":
        df_filtered = df_filtered[df_filtered["Municipio"] == muni_filter]
        
    # Gráficos dinámicos
    if len(df_filtered) > 0:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("#### Participación por Rol")
            fig_rol = px.pie(df_filtered, names='Rol', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_rol, use_container_width=True)
            
        with col_g2:
            st.markdown("#### Histórico de Cargas por Tipo")
            fig_bar = px.histogram(df_filtered, x='Tipo_Formulario', color='Tipo_Formulario', color_discrete_sequence=["#0F2C59", "#DAC0A3"])
            st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("No hay datos coincidentes con los filtros seleccionados.")
