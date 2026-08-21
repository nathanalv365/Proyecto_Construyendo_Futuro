import streamlit as st
import pandas as pd
import json
import base64
import io
import datetime
import os
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA (Responsive y Moderna)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado para adaptar la interfaz a móviles y tablets
st.markdown("""
<style>
    .main-header {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #1E3A8A;
        font-size: 2.2rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        text-align: center;
        color: #4B5563;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    .card-pendiente {
        background-color: #FEF3C7;
        padding: 1.5rem;
        border-radius: 0.75rem;
        border-left: 5px solid #F59E0B;
        margin-bottom: 1.5rem;
    }
    .card-info {
        background-color: #EFF6FF;
        padding: 1.25rem;
        border-radius: 0.75rem;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1rem;
    }
    .stButton>button {
        border-radius: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# NORMALIZACIÓN DE DATOS DE GOOGLE SHEETS
# -----------------------------------------------------------------------------
def normalizar_texto(valor):
    """Limpia espacios normales e invisibles sin alterar el contenido útil."""
    if valor is None:
        return ""
    texto = str(valor)
    # Caracteres invisibles que pueden romper comparaciones exactas
    for caracter in ("\ufeff", "\u200b", "\u200c", "\u200d", "\u2060", "\xa0"):
        texto = texto.replace(caracter, " ")
    return " ".join(texto.strip().split())


def normalizar_encabezado(valor):
    """Normaliza encabezados para evitar desalineaciones por espacios/carácteres invisibles."""
    return normalizar_texto(valor).casefold()


def construir_dataframe_hoja(all_rows, expected_cols):
    """
    Convierte la hoja en DataFrame usando encabezados canónicos.
    Corrige:
      - espacios/carácteres invisibles en encabezados,
      - filas con menos/más columnas,
      - celdas con espacios extra.
    """
    if not all_rows:
        return pd.DataFrame(columns=expected_cols)

    headers_raw = list(all_rows[0])
    headers = [normalizar_texto(h) for h in headers_raw]

    # Solo recortar columnas vacías al final; no tocar columnas intermedias.
    while headers and headers[-1] == "":
        headers.pop()

    if not headers:
        return pd.DataFrame(columns=expected_cols)

    aliases = {normalizar_encabezado(col): col for col in expected_cols}

    # Convertir cualquier variante del encabezado a su nombre oficial.
    headers_canonicos = []
    usados = set()
    for header in headers:
        clave = normalizar_encabezado(header)
        canonico = aliases.get(clave, header)
        # Evitar duplicados de encabezados; conserva el primero.
        if canonico in usados and canonico in expected_cols:
            canonico = f"{canonico}__duplicado_{len(headers_canonicos)}"
        headers_canonicos.append(canonico)
        usados.add(canonico)

    registros = []
    for row in all_rows[1:]:
        row_values = [normalizar_texto(v) for v in list(row)]
        if len(row_values) < len(headers_canonicos):
            row_values.extend([""] * (len(headers_canonicos) - len(row_values)))
        elif len(row_values) > len(headers_canonicos):
            row_values = row_values[:len(headers_canonicos)]

        # Ignorar filas totalmente vacías.
        if not any(row_values):
            continue

        registros.append(dict(zip(headers_canonicos, row_values)))

    df = pd.DataFrame(registros)

    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    return df[expected_cols]

# -----------------------------------------------------------------------------
# CONEXIÓN INTEGRADA A GOOGLE SHEETS (Con Robustez Extrema)
# -----------------------------------------------------------------------------
@st.cache_resource
def conectar_google_sheets():
    """Establece la conexión segura con la API de Google Sheets."""
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # 1. Intentar cargar desde Secrets de Streamlit (Entorno Cloud de Producción)
    if "gcp_service_account" in st.secrets:
        try:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Error de autenticación con Google Secrets: {e}")
            
    # 2. Intentar cargar desde archivo local (Desarrollo local)
    elif os.path.exists("credentials.json"):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Error de autenticación con credentials.json: {e}")
            
    return None

def obtener_hoja_calculo():
    """Obtiene siempre el objeto Spreadsheet actual, sin cachear datos de hojas."""
    gc = conectar_google_sheets()
    if gc is None:
        return None

    nombre_db = "Base_Encuestas_SRPA"
    try:
        return gc.open(nombre_db)
    except gspread.exceptions.SpreadsheetNotFound:
        # Si no existe, la creamos desde cero con la estructura oficial limpia.
        sh = gc.create(nombre_db)

        # Pestaña 1: Cola_Revision
        ws_cola = sh.get_worksheet(0)
        ws_cola.update_title("Cola_Revision")
        headers_cola = [
            "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio",
            "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado",
            "Foto_P1_Base64", "Foto_P2_Base64", "Verificado_Por"
        ]
        ws_cola.append_row(headers_cola, value_input_option="RAW", insert_data_option="INSERT_ROWS")

        # Pestaña 2: Respuestas_SRPA
        headers_respuestas = [
            "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio",
            "Institucion_Educativa_Verificada", "Rol",
            "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
            "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
            "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
            "Verificado_Por", "Fecha_Aprobacion"
        ]
        ws_respuestas = sh.add_worksheet(title="Respuestas_SRPA", rows="1000", cols="30")
        ws_respuestas.append_row(headers_respuestas, value_input_option="RAW", insert_data_option="INSERT_ROWS")

        return sh

# -----------------------------------------------------------------------------
# CARGA DE DATOS ROBUSTA (Evitando GSpreadException y KeyError de forma absoluta)
# -----------------------------------------------------------------------------
def cargar_cola_revision():
    """Carga la cola y normaliza encabezados/celdas para evitar falsos vacíos."""
    sh = obtener_hoja_calculo()
    expected_cols = [
        "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio",
        "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado",
        "Foto_P1_Base64", "Foto_P2_Base64", "Verificado_Por"
    ]

    if sh is None:
        return pd.DataFrame(columns=expected_cols)

    try:
        ws = sh.worksheet("Cola_Revision")
        all_rows = ws.get_all_values()

        if not all_rows or len(all_rows) < 2:
            return pd.DataFrame(columns=expected_cols)

        return construir_dataframe_hoja(all_rows, expected_cols)

    except Exception as e:
        st.error(f"Error al cargar Cola_Revision de Google Sheets: {e}")
        return pd.DataFrame(columns=expected_cols)

def cargar_respuestas_validadas():
    """Carga de forma segura las respuestas definitivas aprobadas."""
    sh = obtener_hoja_calculo()
    expected_cols = [
        "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio",
        "Institucion_Educativa_Verificada", "Rol",
        "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
        "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
        "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
        "Verificado_Por", "Fecha_Aprobacion"
    ]

    if sh is None:
        return pd.DataFrame(columns=expected_cols)

    try:
        ws = sh.worksheet("Respuestas_SRPA")
        all_rows = ws.get_all_values()

        if not all_rows or len(all_rows) < 2:
            return pd.DataFrame(columns=expected_cols)

        return construir_dataframe_hoja(all_rows, expected_cols)

    except Exception as e:
        st.error(f"Error al cargar Respuestas_SRPA de Google Sheets: {e}")
        return pd.DataFrame(columns=expected_cols)

# -----------------------------------------------------------------------------
# GUARDADO Y CONFIRMACIÓN DE DOBLE VÍA (Inmune a fallos de retraso de Google)
# -----------------------------------------------------------------------------
def _buscar_fila_por_id(ws, id_encuesta):
    """Devuelve el número de fila real en Sheets para un ID, ignorando espacios invisibles."""
    id_normalizado = normalizar_texto(id_encuesta)
    all_rows = ws.get_all_values()

    if not all_rows:
        return None, None, None

    headers = [normalizar_texto(h) for h in all_rows[0]]
    headers_norm = [normalizar_encabezado(h) for h in headers]
    idx_id = headers_norm.index(normalizar_encabezado("ID_Encuesta")) if normalizar_encabezado("ID_Encuesta") in headers_norm else 0

    for sheet_row, row in enumerate(all_rows[1:], start=2):
        if idx_id < len(row) and normalizar_texto(row[idx_id]) == id_normalizado:
            return sheet_row, headers, row

    return None, headers, None


def _actualizar_estado_cola(ws_cola, id_encuesta, estado, usuario_revisor):
    """Actualiza Estado y Verificado_Por en una única operación."""
    sheet_row, headers, _ = _buscar_fila_por_id(ws_cola, id_encuesta)
    if sheet_row is None:
        return False

    headers_norm = [normalizar_encabezado(h) for h in headers]
    try:
        col_estado = headers_norm.index(normalizar_encabezado("Estado")) + 1
        col_verificador = headers_norm.index(normalizar_encabezado("Verificado_Por")) + 1
    except ValueError:
        return False

    # update_cells reduce llamadas a Google y evita que una actualización quede a medias.
    c1 = gspread.utils.rowcol_to_a1(sheet_row, col_estado)
    c2 = gspread.utils.rowcol_to_a1(sheet_row, col_verificador)
    ws_cola.update(f"{c1}:{c2}", [[estado, usuario_revisor]], value_input_option="RAW")

    # Confirmación de doble vía.
    _, _, row_verificada = _buscar_fila_por_id(ws_cola, id_encuesta)
    if row_verificada is None:
        return False

    estado_idx = col_estado - 1
    verificador_idx = col_verificador - 1
    return (
        estado_idx < len(row_verificada)
        and normalizar_encabezado(row_verificada[estado_idx]) == normalizar_encabezado(estado)
        and verificador_idx < len(row_verificada)
        and normalizar_texto(row_verificada[verificador_idx]) == normalizar_texto(usuario_revisor)
    )


def aprobar_e_ingresar_registro(id_encuesta, tipo_formulario, fecha, municipio, ie_verificada, rol, conocimientos, satisfaccion, usuario_revisor):
    """
    Guarda en Respuestas_SRPA y luego marca como Aprobado la fila exacta en Cola_Revision.
    Evita duplicados si el usuario vuelve a pulsar aprobar después de un error de red.
    """
    sh = obtener_hoja_calculo()
    if sh is None:
        st.error("No hay conexión con Google Sheets.")
        return False

    try:
        ws_respuestas = sh.worksheet("Respuestas_SRPA")
        ws_cola = sh.worksheet("Cola_Revision")

        id_encuesta = normalizar_texto(id_encuesta)

        # 1. Evitar duplicar el registro definitivo si ya fue ingresado.
        fila_existente, _, _ = _buscar_fila_por_id(ws_respuestas, id_encuesta)
        if fila_existente is None:
            fila_nueva = [
                id_encuesta,
                normalizar_texto(tipo_formulario).upper(),
                normalizar_texto(fecha),
                normalizar_texto(municipio),
                normalizar_texto(ie_verificada),
                normalizar_texto(rol),
                normalizar_texto(conocimientos.get("p1", "")),
                normalizar_texto(conocimientos.get("p2", "")),
                normalizar_texto(conocimientos.get("p3", "")),
                normalizar_texto(conocimientos.get("p4", "")),
                normalizar_texto(conocimientos.get("p5", "")),
                normalizar_texto(conocimientos.get("p6", "")),
                normalizar_texto(conocimientos.get("p7", "")),
                normalizar_texto(conocimientos.get("p8", "")),
                normalizar_texto(satisfaccion.get("sat_p1", "")),
                normalizar_texto(satisfaccion.get("sat_p2", "")),
                normalizar_texto(satisfaccion.get("sat_p3", "")),
                normalizar_texto(satisfaccion.get("sat_p4", "")),
                normalizar_texto(satisfaccion.get("sat_p5", "")),
                normalizar_texto(satisfaccion.get("sat_p6", "")),
                normalizar_texto(satisfaccion.get("sat_p7", "")),
                normalizar_texto(satisfaccion.get("sat_p8", "")),
                normalizar_texto(satisfaccion.get("sat_p9", "")),
                normalizar_texto(usuario_revisor),
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]

            res = ws_respuestas.append_row(
                fila_nueva,
                value_input_option="RAW",
                insert_data_option="INSERT_ROWS",
                include_values_in_response=True
            )

            if not res:
                st.error("Fallo al escribir en la hoja Respuestas_SRPA.")
                return False

            # Confirmar inmediatamente que Google realmente guardó el ID.
            fila_verificada, _, _ = _buscar_fila_por_id(ws_respuestas, id_encuesta)
            if fila_verificada is None:
                st.error("Google Sheets no confirmó la escritura en Respuestas_SRPA.")
                return False

        # 2. Actualizar la misma encuesta en la cola.
        if not _actualizar_estado_cola(
            ws_cola, id_encuesta, "Aprobado", normalizar_texto(usuario_revisor)
        ):
            st.error(
                "La respuesta definitiva quedó guardada, pero no se pudo actualizar "
                "la encuesta correspondiente en Cola_Revision."
            )
            return False

        return True

    except Exception as e:
        st.error(f"Fallo durante la aprobación del registro: {e}")
        return False


def rechazar_registro_cola(id_encuesta, usuario_revisor):
    """Marca un registro en la Cola_Revision como 'Rechazado'."""
    sh = obtener_hoja_calculo()
    if sh is None:
        return False

    try:
        ws_cola = sh.worksheet("Cola_Revision")
        return _actualizar_estado_cola(
            ws_cola,
            normalizar_texto(id_encuesta),
            "Rechazado",
            normalizar_texto(usuario_revisor)
        )
    except Exception as e:
        st.error(f"Error al rechazar el registro: {e}")
        return False

# -----------------------------------------------------------------------------
# COMPRESIÓN DE IMÁGENES ULTRA-COMPATIBLE (Soporta LANCZOS, ANTIALIAS o fallback)
# -----------------------------------------------------------------------------
def comprimir_imagen_base64(uploaded_file, max_width=450):
    """Redimensiona y comprime la imagen a base64 ligero optimizado para Google Sheets."""
    try:
        img = Image.open(uploaded_file)
        
        # Convertir a RGB si tiene canal Alpha
        if img.mode in ('RGBA', 'LA'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
            
        # Calcular proporciones
        w, h = img.size
        if w > max_width:
            ratio = max_width / float(w)
            new_size = (max_width, int(float(h) * ratio))
            
            # Selección de método de remuestreo dinámica e indestructible
            resample_method = None
            if hasattr(Image, "Resampling") and hasattr(Image.Resampling, "LANCZOS"):
                resample_method = Image.Resampling.LANCZOS
            elif hasattr(Image, "LANCZOS"):
                resample_method = Image.LANCZOS
            elif hasattr(Image, "ANTIALIAS"):
                resample_method = Image.ANTIALIAS
                
            if resample_method is not None:
                img = img.resize(new_size, resample_method)
            else:
                img = img.resize(new_size)
                
        # Guardar en memoria comprimida
        buffered = io.BytesIO()
        img.save(buffered, format="JPEG", quality=75)
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        st.error(f"Error comprimiendo imagen: {e}")
        return ""

# -----------------------------------------------------------------------------
# CONSULTAS MULTIMODALES INTELIGENTES (Gemini 2.5/3.6 con esquema de reintentos)
# -----------------------------------------------------------------------------
def analizar_encuesta_fisica(base64_p1, base64_p2=None):
    """Envía la/s foto/s a Google Gemini usando un pool de modelos y fallback de robustez."""
    api_key = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("Por favor, ingresa tu GEMINI_API_KEY en los Secrets de Streamlit.")
        return None
        
    system_prompt = """
    Analiza la/s imagen/es de esta encuesta del "Proyecto Construyendo Futuro" (SRPA).
    Extrae la información manuscrita del encabezado y las marcas hechas con (X) en las casillas.
    Sigue las siguientes reglas de extracción estrictas:
    1. Identifica el tipo de test: "PRETEST" o "POSTEST" según el título superior.
    2. En el encabezado lee la caligrafía escrita a mano:
       - Fecha (en formato YYYY-MM-DD o deja vacío "" si no está clara)
       - Municipio
       - Institución Educativa (presta suma atención al nombre manuscrito del colegio)
       - Rol del participante (debe ser Estudiante, Docente, Padre de Familia o Lider comunitario)
    3. Lee las respuestas:
       - Para conocimientos (Conocimientos_P1 a Conocimientos_P8): marca la opción elegida por el usuario (ej: "a", "b", "c"). Si no hay marca o está en blanco, pon "".
       - Para evaluación de satisfacción (sólo si es POSTEST): lee la matriz de satisfacción (Preguntas 1 a 9). Los valores válidos son "Excelente", "Bueno", "Regular", o "Deficiente". Si el participante dejó alguna vacía, pon "".
    
    Genera un JSON estrictamente estructurado en este formato:
    {
      "tipo_formulario": "PRETEST",
      "fecha": "2026-07-30",
      "municipio": "Cartagena",
      "institucion_educativa": "Promesa de Dios",
      "rol": "Estudiante",
      "conocimientos": {
         "p1": "b", "p2": "b", "p3": "b", "p4": "a",
         "p5": "d", "p6": "a", "p7": "a", "p8": "a"
      },
      "satisfaccion": {
         "sat_p1": "", "sat_p2": "", "sat_p3": "", "sat_p4": "",
         "sat_p5": "", "sat_p6": "", "sat_p7": "", "sat_p8": "", "sat_p9": ""
      }
    }
    """
    
    # Lista de modelos por orden de preferencia y vigencia de la API
    modelos_pool = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash", "gemini-1.5-flash"]
    
    # 1. Intentar conectarse usando el nuevo SDK de Google GenAI
    try:
        from google import genai
        from google.genai import types
        
        client = genai.Client(api_key=api_key)
        
        # Preparar partes multimedia
        partes = []
        partes.append(types.Part.from_bytes(data=base64.b64decode(base64_p1), mime_type="image/jpeg"))
        if base64_p2:
            partes.append(types.Part.from_bytes(data=base64.b64decode(base64_p2), mime_type="image/jpeg"))
        partes.append(types.Part.from_text(text=system_prompt))
        
        for modelo in modelos_pool:
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=partes,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                if response.text:
                    # Limpieza segura de marcas de markdown antes de parsear
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    return json.loads(clean_text)
            except Exception:
                continue # Probar siguiente modelo del pool
    except Exception:
        pass # Fallback al SDK tradicional
        
    # 2. Fallback usando el SDK tradicional google-generativeai
    try:
        import google.generativeai as genai_legacy
        genai_legacy.configure(api_key=api_key)
        
        partes_legacy = []
        partes_legacy.append({"mime_type": "image/jpeg", "data": base64_p1})
        if base64_p2:
            partes_legacy.append({"mime_type": "image/jpeg", "data": base64_p2})
        partes_legacy.append(system_prompt)
        
        for modelo in modelos_pool:
            try:
                model = genai_legacy.GenerativeModel(modelo)
                response = model.generate_content(partes_legacy)
                if response.text:
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    return json.loads(clean_text)
            except Exception:
                continue
    except Exception as e:
        st.error(f"Error de inicialización de la IA en el servidor: {e}")
        
    return None

# -----------------------------------------------------------------------------
# INTERFAZ GRÁFICA DE USUARIO (Streamlit Multi-Pestañas)
# -----------------------------------------------------------------------------
def main():
    # Banner de la Gobernación de Bolívar y Encabezado
    col_logo1, col_logo2 = st.columns([1, 4])
    with col_logo1:
        st.image("https://www.bolivar.gov.co/images/Gobernacion/Logo-Gobernacion-Color.png", use_container_width=True)
    with col_logo2:
        st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
        st.markdown("<p class='sub-header'>SISTEMA COLABORATIVO DE VERIFICACIÓN DE CONOCIMIENTOS SRPA</p>", unsafe_allow_html=True)
        
    # Comprobar estado de la conexión a Google Sheets
    sh = obtener_hoja_calculo()
    if sh is None:
        st.warning("⚠️ Sin conexión activa a Google Sheets.")
        st.info("""
        **Instrucciones para conectar Google Sheets:**
        1. Crea una cuenta de servicio en Google Cloud Console y descarga tus credenciales JSON.
        2. Ve al panel de control de Streamlit Cloud -> Manage App -> Settings -> Secrets.
        3. Pega tus credenciales JSON en una variable llamada `[gcp_service_account]`.
        """)
        return
        
    # Definición de Pestañas
    tab_dashboard, tab_carga, tab_revision = st.tabs([
        "📊 Dashboard Estadístico", 
        "📱 Carga en Campo (Cámara)", 
        "🔍 Banco de Verificación"
    ])
    
    # -------------------------------------------------------------------------
    # PESTAÑA 1: DASHBOARD ESTADÍSTICO (Cero Simulación)
    # -------------------------------------------------------------------------
    with tab_dashboard:
        df_respuestas = cargar_respuestas_validadas()
        
        if df_respuestas.empty:
            st.info("👋 **¡Bienvenido al sistema de verificación SRPA!**")
            st.write("La base de datos se encuentra limpia y lista para recibir registros en campo. Comienza cargando las encuestas en la pestaña **Carga en Campo**.")
        else:
            st.success(f"📊 Se han consolidado **{len(df_respuestas)}** encuestas reales del proyecto en la Gobernación de Bolívar.")
            
            # Filtros dinámicos responsive
            col_filtros1, col_filtros2 = st.columns(2)
            with col_filtros1:
                municipio_filtro = st.selectbox("Filtrar por Municipio", ["Todos"] + sorted(list(df_respuestas["Municipio"].unique())))
            with col_filtros2:
                rol_filtro = st.selectbox("Filtrar por Rol", ["Todos"] + sorted(list(df_respuestas["Rol"].unique())))
                
            df_filtrado = df_respuestas.copy()
            if municipio_filtro != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Municipio"] == municipio_filtro]
            if rol_filtro != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Rol"] == rol_filtro]
                
            # Métricas rápidas
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric("Total Registros Filtrados", len(df_filtrado))
            tipo_normalizado = df_filtrado["Tipo_Formulario"].map(normalizar_encabezado)
            pretest_count = int((tipo_normalizado == "pretest").sum())
            posttest_count = int((tipo_normalizado == "postest").sum())

            with col_m2:
                st.metric("Total Pretest", pretest_count)
            with col_m3:
                st.metric("Total Postest", posttest_count)
                
            # Gráficos de Satisfacción si existen Postest
            if posttest_count > 0:
                import plotly.express as px

                st.subheader("Evaluación de Satisfacción de las Jornadas (Postest)")

                columnas_sat = [
                    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5",
                    "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9"
                ]

                registros_sat = []
                for columna in columnas_sat:
                    if columna not in df_filtrado.columns:
                        continue
                    valores = (
                        df_filtrado[columna]
                        .map(normalizar_texto)
                        .replace("", pd.NA)
                        .dropna()
                    )
                    for valor in valores:
                        registros_sat.append({
                            "Pregunta": columna.replace("Sat_", "Pregunta "),
                            "Valoración": valor
                        })

                if registros_sat:
                    df_sat = pd.DataFrame(registros_sat)
                    grafico = (
                        df_sat.groupby(["Pregunta", "Valoración"])
                        .size()
                        .reset_index(name="Cantidad")
                    )
                    fig = px.bar(
                        grafico,
                        x="Pregunta",
                        y="Cantidad",
                        color="Valoración",
                        barmode="group",
                        text_auto=True
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Hay postest aprobados, pero todavía no hay valores de satisfacción registrados.")
                
    # -------------------------------------------------------------------------
    # PESTAÑA 2: CARGA EN CAMPO (Cámara móvil / Galería)
    # -------------------------------------------------------------------------
    with tab_carga:
        st.subheader("📱 Capturar Nueva Encuesta")
        st.write("Sube las imágenes de la encuesta. La IA de Gemini leerá el texto de forma autónoma.")
        
        tipo_form = st.selectbox("Tipo de Formulario a Subir", ["PRETEST", "POSTEST"])
        usuario = st.text_input("Nombre del Facilitador que carga la encuesta", "Equipo Bolívar")
        
        col_img1, col_img2 = st.columns(2)
        with col_img1:
            file_p1 = st.file_uploader("📷 Subir Foto Página 1 (Encabezado y Preguntas 1-3)", type=["jpg", "png", "jpeg"])
        with col_img2:
            file_p2 = None
            if tipo_form == "POSTEST":
                file_p2 = st.file_uploader("📷 Subir Foto Página 2 (Satisfacción)", type=["jpg", "png", "jpeg"])
                
        if file_p1:
            if st.button("🚀 Procesar Encuesta con IA (Gemini)", use_container_width=True):
                with st.spinner("Procesando y digitalizando encuesta con Gemini..."):
                    base64_p1 = comprimir_imagen_base64(file_p1)
                    base64_p2 = comprimir_imagen_base64(file_p2) if file_p2 else ""
                    
                    # Llamar a la IA
                    data_ia = analizar_encuesta_fisica(base64_p1, base64_p2)
                    
                    if data_ia:
                        # Crear ID único
                        id_encuesta = f"ENC_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        fecha_carga = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        # Escribir en Cola_Revision
                        ws_cola = sh.worksheet("Cola_Revision")
                        fila_cola = [
                            id_encuesta,
                            fecha_carga,
                            tipo_form,
                            data_ia.get("municipio", ""),
                            data_ia.get("institucion_educativa", ""),
                            data_ia.get("rol", ""),
                            json.dumps(data_ia),
                            "Pendiente",
                            base64_p1,
                            base64_p2,
                            usuario
                        ]
                        
                        try:
                            ws_cola.append_row(
                                fila_cola,
                                value_input_option="RAW",
                                insert_data_option="INSERT_ROWS",
                                include_values_in_response=True
                            )

                            # Confirmar que el registro se puede volver a localizar inmediatamente.
                            fila_confirmada, _, _ = _buscar_fila_por_id(ws_cola, id_encuesta)
                            if fila_confirmada is None:
                                st.error(
                                    "La encuesta fue enviada a Google Sheets, pero no se pudo "
                                    "confirmar su presencia en Cola_Revision. No se marcará como exitosa."
                                )
                            else:
                                st.success(
                                    f"✅ Encuesta subida y encolada exitosamente para revisión. ID: {id_encuesta}"
                                )
                                st.info(
                                    "La encuesta ya está en Cola_Revision. Ve a la pestaña "
                                    "**Banco de Verificación** para corregir el nombre y aprobarla."
                                )
                        except Exception as e:
                            st.error(f"Error al guardar la encuesta en Cola_Revision: {e}")
                    else:
                        st.error("Error al procesar la encuesta con Gemini. Por favor, asegúrate de tener una conexión estable y una GEMINI_API_KEY válida.")
                        
    # -------------------------------------------------------------------------
    # PESTAÑA 3: BANCO DE VERIFICACIÓN (Lado a Lado e Inmune a Desfase de Columnas)
    # -------------------------------------------------------------------------
    with tab_revision:
        st.subheader("🔍 Banco de Verificación de Escritura a Mano")
        st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
        
        col_refresh, _ = st.columns([1, 5])
        with col_refresh:
            if st.button("🔄 Actualizar cola", use_container_width=True):
                st.rerun()

        df_cola = cargar_cola_revision()
        
        # Filtrar registros en Estado "Pendiente" tolerando espacios/carácteres invisibles.
        if not df_cola.empty:
            df_cola["_Estado_Normalizado"] = df_cola["Estado"].map(normalizar_encabezado)
            df_cola["_ID_Normalizado"] = df_cola["ID_Encuesta"].map(normalizar_texto)
            df_pendientes = df_cola[
                (df_cola["_Estado_Normalizado"] == "pendiente")
                & (df_cola["_ID_Normalizado"] != "")
            ].copy()
        else:
            df_pendientes = df_cola.copy()
        
        if df_pendientes.empty:
            st.info("🎉 **¡Excelente trabajo! No hay encuestas pendientes de verificación en la cola de revisión.**")
        else:
            st.write(f"Hay **{len(df_pendientes)}** encuesta(s) pendiente(s) de aprobación:")
            
            # Tomar la primera encuesta en la cola
            row = df_pendientes.iloc[0]
            
            # Mostrar la tarjeta de revisión
            st.markdown(f"""
            <div class='card-pendiente'>
                <h3>Encuesta: {row['ID_Encuesta']}</h3>
                <p><strong>Fecha de Carga:</strong> {row['Fecha_Carga']} | <strong>Tipo:</strong> {row['Tipo_Formulario']}</p>
                <p><strong>Cargado Por:</strong> {row['Verificado_Por']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Diseño de Columnas para Visualización Lado a Lado
            col_preview, col_form = st.columns([1, 1])
            
            with col_preview:
                st.subheader("📷 Imágenes del Documento Físico")
                
                # Mostrar Página 1
                if row['Foto_P1_Base64']:
                    try:
                        p1_bytes = base64.b64decode(row['Foto_P1_Base64'])
                        st.image(p1_bytes, caption="Página 1 (Encabezado y Conocimientos)", use_container_width=True)
                    except Exception:
                        st.warning("No se pudo previsualizar la foto de la Página 1.")
                else:
                    st.info("Sin foto de Página 1.")
                    
                # Mostrar Página 2
                if row['Foto_P2_Base64']:
                    try:
                        p2_bytes = base64.b64decode(row['Foto_P2_Base64'])
                        st.image(p2_bytes, caption="Página 2 (Satisfacción)", use_container_width=True)
                    except Exception:
                        st.warning("No se pudo previsualizar la foto de la Página 2.")
                        
            with col_form:
                st.subheader("✏️ Campos Interpretados por la IA")
                
                # Parsear el JSON guardado de forma segura
                try:
                    data_json = json.loads(row['JSON_Respuestas'])
                except Exception:
                    data_json = {}
                    
                # Inputs editables para el Verificador Humano
                fecha_revisada = st.text_input("Fecha (YYYY-MM-DD)", value=row['Fecha_Carga'][:10])
                municipio_revisado = st.text_input("Municipio", value=row['Municipio'] if row['Municipio'] else data_json.get("municipio", ""))
                ie_revisada = st.text_input("Institución Educativa (Escrita a Mano en Foto)", value=row['Institucion_Educativa_IA'] if row['Institucion_Educativa_IA'] else data_json.get("institucion_educativa", ""))
                rol_revisado = st.selectbox("Rol del Participante", ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(row['Rol']) if row['Rol'] in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0)
                
                st.write("---")
                st.write("📋 **Respuestas leídas de la encuesta:**")
                st.write(data_json.get("conocimientos", {}))
                
                if normalizar_encabezado(row['Tipo_Formulario']) == "postest":
                    st.write("📊 **Respuestas de satisfacción:**")
                    st.write(data_json.get("satisfaccion", {}))
                    
                usuario_aprobador = st.text_input("Nombre de quien aprueba el registro", "Revisor Oficial")
                
                st.write("---")
                # Botones de Acción directos lado a lado
                col_btn_aprob, col_btn_rech = st.columns(2)
                
                with col_btn_aprob:
                    if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True, type="primary"):
                        with st.spinner("Guardando en la base de datos central..."):
                            success = aprobar_e_ingresar_registro(
                                id_encuesta=row['ID_Encuesta'],
                                tipo_formulario=row['Tipo_Formulario'],
                                fecha=fecha_revisada,
                                municipio=municipio_revisado,
                                ie_verificada=ie_revisada,
                                rol=rol_revisado,
                                conocimientos=data_json.get("conocimientos", {}),
                                satisfaccion=data_json.get("satisfaccion", {}),
                                usuario_revisor=usuario_aprobador
                            )
                            if success:
                                st.success("¡Encuesta verificada e ingresada correctamente!")
                                st.rerun()
                                
                with col_btn_rech:
                    if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                        with st.spinner("Rechazando encuesta..."):
                            success = rechazar_registro_cola(row['ID_Encuesta'], usuario_aprobador)
                            if success:
                                st.warning("El registro ha sido rechazado y removido de la cola.")
                                st.rerun()

# -----------------------------------------------------------------------------
# EJECUCIÓN DEL SCRIPT
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()
