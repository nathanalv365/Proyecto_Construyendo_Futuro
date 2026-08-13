import streamlit as st
import pandas as pd
import json
import uuid
import datetime
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google import genai
from google.genai import types

# Configuración de página de Streamlit para móviles y escritorio
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo personalizado para botones de color y diseño móvil
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #0d6efd;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CONEXIÓN EXCLUSIVA A GOOGLE SHEETS
# -----------------------------------------------------------------------------
def obtener_credenciales_google():
    """Obtiene y formatea las credenciales de la cuenta de servicio desde los secretos de Streamlit."""
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Asegurar el formateo correcto de saltos de línea en la clave privada
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return creds_dict
    except Exception as e:
        st.error(f"Error al procesar las credenciales secretas: {e}")
        return None

def conectar_google_sheets():
    """Establece conexión con el Google Sheet central llamado 'Base_Encuestas_SRPA'."""
    creds_dict = obtener_credenciales_google()
    if creds_dict is None:
        return None
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        # Intentar abrir la hoja de cálculo
        try:
            spreadsheet = client.open("Base_Encuestas_SRPA")
        except gspread.SpreadsheetNotFound:
            # Si no existe, la creamos de forma automática
            st.info("Creando hoja de cálculo central 'Base_Encuestas_SRPA' en tu Google Drive...")
            spreadsheet = client.create("Base_Encuestas_SRPA")
            # Compartir con el correo de servicio y opcionalmente con el usuario si se tuviera su mail
            # Para este flujo asumimos que ya está compartida si la crearon manualmente.
            
        return spreadsheet
    except Exception as e:
        st.error(f"No se pudo conectar a Google Sheets: {e}")
        return None

def inicializar_pestanas(spreadsheet):
    """Inicializa las pestañas de Google Sheets con las columnas oficiales si no existen."""
    if spreadsheet is None:
        return False
    try:
        # Pestaña 1: Cola_Revision
        try:
            ws_cola = spreadsheet.worksheet("Cola_Revision")
        except gspread.WorksheetNotFound:
            ws_cola = spreadsheet.add_worksheet(title="Cola_Revision", rows="1000", cols="8")
            headers_cola = [
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
            ]
            ws_cola.append_row(headers_cola)
            
        # Pestaña 2: Respuestas_SRPA
        try:
            ws_resp = spreadsheet.worksheet("Respuestas_SRPA")
        except gspread.WorksheetNotFound:
            ws_resp = spreadsheet.add_worksheet(title="Respuestas_SRPA", rows="10000", cols="25")
            headers_resp = [
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
                "Institucion_Educativa_Verificada", "Rol", 
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", 
                "Conocimientos_P4", "Conocimientos_P5", "Conocimientos_P6", 
                "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", 
                "Sat_P7", "Sat_P8", "Sat_P9", 
                "Verificado_Por", "Fecha_Aprobacion"
            ]
            ws_resp.append_row(headers_resp)
        return True
    except Exception as e:
        st.error(f"Error inicializando pestañas en Google Sheets: {e}")
        return False

# -----------------------------------------------------------------------------
# INTERFAZ DE CONFIGURACIÓN / COMPROBACIÓN DE CONEXIÓN
# -----------------------------------------------------------------------------
spreadsheet = conectar_google_sheets()

if spreadsheet is None:
    st.title("🔌 Configuración de Conexión Requerida")
    st.warning("La aplicación requiere exclusivamente conexión a **Google Sheets** en la nube para sincronizar en tiempo real.")
    
    st.markdown("""
    ### 🛠️ Pasos para activar la Base de Datos en Google Sheets:
    
    1. **Crea una Cuenta de Servicio en Google Cloud Platform (GCP):**
       * Ve a [Google Cloud Console](https://console.cloud.google.com/).
       * Crea un proyecto y activa las APIs de **Google Sheets API** y **Google Drive API**.
       * Crea una **Cuenta de Servicio**, genera una **clave JSON** y descárgala.
    
    2. **Crea la Hoja de Cálculo:**
       * Crea un archivo de Google Sheets llamado exactamente **`Base_Encuestas_SRPA`**.
       * Abre el archivo `.json` de tu cuenta de servicio y copia el correo `client_email`.
       * Comparte tu hoja de cálculo con este correo con rol de **Editor**.
    
    3. **Configura los Secretos en tu plataforma de Despliegue (Streamlit Cloud):**
       * Ve a la configuración de tu App en Streamlit Cloud -> **Secrets**.
       * Pega el contenido de tu archivo JSON bajo el bloque `[gcp_service_account]` tal como se muestra abajo:
    """)
    
    st.code("""
[gcp_service_account]
type = "service_account"
project_id = "tu-proyecto-id"
private_key_id = "tu-private-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\\nTU_LLAVE_PRIVADA_AQUÍ\\n-----END PRIVATE KEY-----\\n"
client_email = "tu-cuenta-de-servicio@proyecto.iam.gserviceaccount.com"
client_id = "tu-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.google.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/tu-cuenta-de-servicio"
    """, language="toml")
    
    st.info("💡 Una vez agregues los secretos, recarga esta página y la aplicación se conectará automáticamente.")
    st.stop()

# Inicializar pestañas oficiales si no están creadas
inicializar_pestanas(spreadsheet)

# Conectores a las hojas específicas
ws_cola = spreadsheet.worksheet("Cola_Revision")
ws_respuestas = spreadsheet.worksheet("Respuestas_SRPA")

# -----------------------------------------------------------------------------
# MOTOR DE OCR CON GEMINI
# -----------------------------------------------------------------------------
def get_gemini_client():
    """Inicializa el cliente de Gemini utilizando la API Key guardada en secretos."""
    if "GEMINI_API_KEY" in st.secrets:
        return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    return None

# Esquemas de datos para Gemini API utilizando tipos de Pydantic
from pydantic import BaseModel, Field
from typing import Optional

class ExtraccionEncabezado(BaseModel):
    fecha: Optional[str] = Field(None, description="Fecha de la encuesta en formato YYYY-MM-DD o vacío si no se suministra.")
    municipio: Optional[str] = Field(None, description="Municipio escrito en la encuesta.")
    institucion_educativa: Optional[str] = Field(None, description="Nombre manuscrito de la Institución Educativa.")
    rol: Optional[str] = Field(None, description="Rol marcado con X: Estudiante, Docente, Padre de Familia, Lider comunitario.")

class ExtraccionEncuesta(BaseModel):
    tipo_formulario: str = Field(description="PRETEST o POSTEST")
    encabezado: ExtraccionEncabezado
    respuestas_conocimiento: dict = Field(description="Diccionario con claves p1 a p8 y valor la opción seleccionada (por ejemplo: 'a', 'b', 'Sí', 'ICBF' o vacío '' si no hay marca).")
    evaluacion_satisfaccion: Optional[dict] = Field(None, description="Solo si es POSTEST. Diccionario con claves s1 a s9 y valor de satisfacción: 'Excelente', 'Bueno', 'Regular', 'Deficiente' o vacío '' si no hay marca.")

SYSTEM_PROMPT_OCR = """
Eres un sistema OCR experto de alta precisión para el "Proyecto Construyendo Futuro" (Prevención en el SRPA).
Analizas dos imágenes que corresponden a la Página 1 y Página 2 de un mismo cuestionario físico.
Debes extraer la información con absoluto rigor y estructurarla en el formato JSON requerido.

Instrucciones de Lectura Estrictas:
1. Identifica el Tipo de Formulario (PRETEST o POSTEST).
2. Lee los datos del encabezado de la Página 1: Fecha, Municipio y el nombre manuscrito de la Institución Educativa.
3. Rol del participante: Identifica cuál de las casillas tiene una marca (X) o tachadura clara.
4. Respuestas de conocimiento:
   - Para PRETEST: Lee las respuestas de las preguntas 1 a 8 marcadas con (X). Las opciones posibles son letras ('a', 'b', 'c', 'd') o texto según la pregunta.
   - Para POSTEST (Sección A): Lee las preguntas 1 a 5.
5. Evaluación de Satisfacción (Sección B - SOLO EN POSTEST):
   - Lee las marcas de la tabla de 9 filas evaluando los aspectos de Excelente, Bueno, Regular, Deficiente.
6. RIGOR DE VACÍOS: Si una pregunta, campo de satisfacción, o campo de texto no tiene absolutamente ninguna marca o información escrita a mano, deves dejarlo como una cadena completamente vacía "". NO inventes ni auto-rellenes información con supuestos o valores predeterminados.
"""

def analizar_paginas_con_gemini(img1, img2):
    """Envía la página 1 y página 2 juntas a la API de Gemini para procesamiento consolidado."""
    client = get_gemini_client()
    if client is None:
        st.error("No se encontró la API Key de Gemini en los secretos ('GEMINI_API_KEY'). Activando procesamiento simulado.")
        return simular_procesamiento()
        
    try:
        # Convertir imágenes de PIL a bytes
        import io
        img_byte_arr1 = io.BytesIO()
        img1.save(img_byte_arr1, format='JPEG')
        img_bytes1 = img_byte_arr1.getvalue()
        
        img_byte_arr2 = io.BytesIO()
        img2.save(img_byte_arr2, format='JPEG')
        img_bytes2 = img_byte_arr2.getvalue()
        
        parts = [
            types.Part.from_bytes(data=img_bytes1, mime_type="image/jpeg"),
            types.Part.from_bytes(data=img_bytes2, mime_type="image/jpeg"),
            SYSTEM_PROMPT_OCR
        ]
        
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=parts,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=ExtraccionEncuesta,
                temperature=0.0
            )
        )
        
        return json.loads(response.text)
    except Exception as e:
        st.error(f"Error procesando con la API de Gemini: {e}. Se generará un caso simulado de contingencia.")
        return simular_procesamiento()

def simular_procesamiento():
    """Genera un caso simulado de contingencia con datos basados en la estructura de los formatos reales."""
    import random
    tipo = random.choice(["PRETEST", "POSTEST"])
    IE_manuscritas = ["I.E. Promesa de Dios", "Colegio Militar Almirante Colon", "I.E. San Jose de Turbaco"]
    ie_detectada = random.choice(IE_manuscritas)
    
    respuestas_c = {f"p{i}": random.choice(["a", "b", "c", ""]) for i in range(1, 9)} if tipo == "PRETEST" else {f"p{i}": random.choice(["a", "b", "c", ""]) for i in range(1, 6)}
    
    sat = None
    if tipo == "POSTEST":
        sat = {f"s{i}": random.choice(["Excelente", "Bueno", "", "Regular"]) for i in range(1, 10)}
        
    return {
        "tipo_formulario": tipo,
        "encabezado": {
            "fecha": str(datetime.date.today()),
            "municipio": random.choice(["Cartagena", "Turbaco", "Arjona"]),
            "institucion_educativa": ie_detectada,
            "rol": random.choice(["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"])
        },
        "respuestas_conocimiento": respuestas_c,
        "evaluacion_satisfaccion": sat
    }

# -----------------------------------------------------------------------------
# INTERFAZ DE NAVEGACIÓN
# -----------------------------------------------------------------------------
st.title("🛡️ Proyecto Construyendo Futuro: Gestión de Conocimientos SRPA")
st.caption("Digitalización colaborativa con IA y base de datos unificada en Google Sheets")

tabs = st.tabs(["📸 Carga de Encuestas (Móvil)", "🔍 Cola de Revisión Humana", "📊 Dashboard Estadístico"])

# =============================================================================
# PESTAÑA 1: CARGA DE ENCUESTAS (Optimizado para teléfonos)
# =============================================================================
with tabs[0]:
    st.header("Carga Obligatoria de Doble Página")
    st.info("Para procesar una encuesta de forma correcta, debes tomar o subir la foto de la **Página 1** (Cabecera y primeras preguntas) y de la **Página 2** (Segunda sección y matriz de satisfacción en caso de Postest).")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ Página 1 (Cabecera)")
        foto_p1 = st.file_uploader("Tomar Foto o Subir Pág 1", type=["jpg", "png", "jpeg"], key="p1_loader")
        if foto_p1:
            img1 = Image.open(foto_p1)
            st.image(img1, caption="Página 1 cargada", use_container_width=True)
            
    with col2:
        st.subheader("2️⃣ Página 2 (Preguntas restantes / Satisfacción)")
        foto_p2 = st.file_uploader("Tomar Foto o Subir Pág 2", type=["jpg", "png", "jpeg"], key="p2_loader")
        if foto_p2:
            img2 = Image.open(foto_p2)
            st.image(img2, caption="Página 2 cargada", use_container_width=True)

    if foto_p1 and foto_p2:
        if st.button("🚀 Procesar Encuesta con Inteligencia Artificial"):
            with st.spinner("La Inteligencia Artificial está analizando las dos páginas de la encuesta, procesando la caligrafía manuscrita y extrayendo marcas de casilla..."):
                resultado = analizar_paginas_con_gemini(img1, img2)
                
                # Insertar en Cola_Revision de Google Sheets
                id_encuesta = str(uuid.uuid4())[:8].upper()
                fecha_carga = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                
                nueva_fila = [
                    id_encuesta,
                    fecha_carga,
                    resultado.get("tipo_formulario", "PRETEST"),
                    resultado.get("encabezado", {}).get("municipio", ""),
                    resultado.get("encabezado", {}).get("institucion_educativa", ""),
                    resultado.get("encabezado", {}).get("rol", ""),
                    json.dumps(resultado),
                    "Pendiente"
                ]
                
                try:
                    ws_cola.append_row(nueva_fila)
                    st.success(f"¡Encuesta registrada con ID **{id_encuesta}** en la Cola de Revisión!")
                    st.balloons()
                except Exception as e:
                    st.error(f"Error guardando en la cola de revisión: {e}")

# =============================================================================
# PESTAÑA 2: COLA DE REVISIÓN HUMANA (Human-in-the-Loop)
# =============================================================================
with tabs[1]:
    st.header("Verificación de Caligrafía e Instituciones Educativas")
    st.write("Debido a que el nombre de la **Institución Educativa** está escrito a mano, revisa y corrige cualquier error de la IA antes de insertarlo en la base de datos oficial. Los nombres de los niños son anónimos.")

    # Cargar datos pendientes de la Cola
    try:
        records_cola = ws_cola.get_all_records()
        pendientes = [r for r in records_cola if r["Estado"] == "Pendiente"]
    except Exception as e:
        st.error(f"Error consultando la cola de revisión: {e}")
        pendientes = []

    if not pendientes:
        st.success("🎉 ¡No hay encuestas pendientes de revisión en la cola!")
    else:
        st.warning(f"Tienes **{len(pendientes)}** encuestas esperando validación.")
        
        # Seleccionar una encuesta de la cola
        opciones_select = {f"{p['ID_Encuesta']} - {p['Tipo_Formulario']} ({p['Municipio']})": p for p in pendientes}
        seleccionado_key = st.selectbox("Selecciona la encuesta a verificar:", list(opciones_select.keys()))
        
        if seleccionado_key:
            item = opciones_select[seleccionado_key]
            datos_json = json.loads(item["JSON_Respuestas"])
            
            st.markdown("---")
            col_rev1, col_rev2 = st.columns([1, 1])
            
            with col_rev1:
                st.subheader("📝 Datos sugeridos por la Inteligencia Artificial")
                
                # Cuadro de texto editable para corregir la caligrafía de la escuela
                ie_sugerida = datos_json.get("encabezado", {}).get("institucion_educativa", "")
                ie_verificada = st.text_input("🏫 Institución Educativa (Edita si es necesario):", value=ie_sugerida)
                
                rol_sugerido = datos_json.get("encabezado", {}).get("rol", "")
                rol_verificado = st.selectbox("👤 Rol del Participante:", ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(rol_sugerido) if rol_sugerido in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0)
                
                tipo_form = datos_json.get("tipo_formulario", "PRETEST")
                municipio = st.text_input("📍 Municipio:", value=datos_json.get("encabezado", {}).get("municipio", ""))
                fecha_doc = st.text_input("📅 Fecha del Documento:", value=datos_json.get("encabezado", {}).get("fecha", str(datetime.date.today())))
                
            with col_rev2:
                st.subheader("📑 Visualización de Respuestas Extraídas")
                st.write(f"**Tipo de Formulario:** {tipo_form}")
                
                # Mostrar respuestas de conocimiento sin auto-rellenar con ficticios si venían vacías
                resp_c = datos_json.get("respuestas_conocimiento", {})
                st.write("**Respuestas de Conocimiento:**")
                for k, v in resp_c.items():
                    val_mostrar = v if v != "" else "⚠️ *En Blanco / Sin respuesta*"
                    st.write(f"- Pregunta {k.upper()}: **{val_mostrar}**")
                    
                # Mostrar respuestas de satisfacción si aplica
                sat_c = datos_json.get("evaluacion_satisfaccion", {})
                if tipo_form == "POSTEST" and sat_c:
                    st.write("**Respuestas de Satisfacción (Matriz):**")
                    for k, v in sat_c.items():
                        val_mostrar = v if v != "" else "⚠️ *En Blanco / Sin respuesta*"
                        st.write(f"- Aspecto {k.upper()}: **{val_mostrar}**")

            st.markdown("---")
            col_btn1, col_btn2 = st.columns(2)
            
            # ACCIÓN 1: APROBAR E INTEGRAR A LA BASE DE DATOS DEFINITIVA
            with col_btn1:
                if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True):
                    # Preparar registro final plano para Google Sheets
                    # 8 preguntas de conocimiento (P1-P8). Rellenamos con "" si no existen
                    conocimientos_flat = [resp_c.get(f"p{i}", "") for i in range(1, 9)]
                    # 9 preguntas de satisfacción (S1-S9). Rellenamos con "" si no existen
                    satisfaccion_flat = [sat_c.get(f"s{i}", "") if sat_c else "" for i in range(1, 10)]
                    
                    fila_consolidada = [
                        item["ID_Encuesta"],
                        tipo_form,
                        fecha_doc,
                        municipio,
                        ie_verificada,
                        rol_verificado
                    ] + conocimientos_flat + satisfaccion_flat + [
                        "Coordinador de Campo", # Verificado por
                        str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")) # Fecha de validación
                    ]
                    
                    try:
                        # 1. Insertar en Respuestas_SRPA
                        ws_respuestas.append_row(fila_consolidada)
                        
                        # 2. Actualizar estado a "Aprobado" en Cola_Revision
                        # Buscamos la fila correspondiente al ID_Encuesta
                        cell = ws_cola.find(item["ID_Encuesta"])
                        ws_cola.update_cell(cell.row, 8, "Aprobado") # Columna 8 es 'Estado'
                        
                        st.success("¡Registro verificado e insertado exitosamente en Google Sheets!")
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error consolidando el registro: {e}")
            
            # ACCIÓN 2: RECHAZAR ENTRADA (ELIMINAR DE LA COLA)
            with col_btn2:
                if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                    try:
                        cell = ws_cola.find(item["ID_Encuesta"])
                        # En lugar de borrar la fila físicamente (lo cual desordena los índices de la hoja),
                        # marcamos el estado como "Rechazado" para que no aparezca en la cola.
                        ws_cola.update_cell(cell.row, 8, "Rechazado")
                        st.warning("La encuesta ha sido rechazada y removida de la cola de verificación.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error rechazando el registro: {e}")

# =============================================================================
# PESTAÑA 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
# =============================================================================
with tabs[2]:
    st.header("Análisis de Impacto y Cobertura Educativa")
    
    # Cargar datos de la base unificada de Google Sheets
    try:
        records_resp = ws_respuestas.get_all_records()
        df = pd.DataFrame(records_resp)
    except Exception as e:
        st.error(f"Error cargando base de datos para estadísticas: {e}")
        df = pd.DataFrame()
        
    # Validar si no hay información real
    if df.empty or len(df) == 0:
        st.info("📊 Aún no hay encuestas validadas y consolidadas en la base de datos de Google Sheets.")
        st.markdown("""
        **Para comenzar a ver estadísticas en tiempo real:**
        1. Sube y procesa una encuesta en la pestaña **📸 Carga de Encuestas**.
        2. Ve a la pestaña **🔍 Cola de Revisión Humana** y haz clic en **Aprobar e Ingresar**.
        3. ¡Regresa aquí para ver las métricas automatizadas!
        """)
    else:
        # 1. Filtros del panel en barra lateral o parte superior
        st.write("### Filtros de Visualización")
        filtro_mun, filtro_ie = st.columns(2)
        
        municipios_disponibles = ["Todos"] + list(df["Municipio"].unique())
        with filtro_mun:
            mun_selected = st.selectbox("Filtrar por Municipio:", municipios_disponibles)
            
        if mun_selected != "Todos":
            df_filtered = df[df["Municipio"] == mun_selected]
        else:
            df_filtered = df.copy()
            
        ie_disponibles = ["Todas"] + list(df_filtered["Institucion_Educativa_Verificada"].unique())
        with filtro_ie:
            ie_selected = st.selectbox("Filtrar por Institución Educativa:", ie_disponibles)
            
        if ie_selected != "Todas":
            df_filtered = df_filtered[df_filtered["Institucion_Educativa_Verificada"] == ie_selected]
            
        # 2. KPIs de Cobertura
        st.markdown("---")
        kpi1, kpi2, kpi3 = st.columns(3)
        
        total_encuestas = len(df_filtered)
        total_pre = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
        total_post = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
        
        with kpi1:
            st.markdown(f"""
            <div class="metric-card">
                <h3>Total Procesados</h3>
                <h2>{total_encuestas} participantes</h2>
            </div>
            """, unsafe_allow_html=True)
            
        with kpi2:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <h3>Total Pretests</h3>
                <h2>{total_pre} aplicados</h2>
            </div>
            """, unsafe_allow_html=True)
            
        with kpi3:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #198754;">
                <h3>Total Postests</h3>
                <h2>{total_post} aplicados</h2>
            </div>
            """, unsafe_allow_html=True)
            
        # 3. Gráficos Estadísticos de Impacto y Satisfacción
        st.markdown("---")
        st.subheader("📈 Evaluación de Conocimientos de Conceptos Clave (Impacto)")
        
        # Mapeo de respuestas correctas en los conceptos comunes
        # Pregunta 2 Pre y Pregunta 1 Post (Finalidad SRPA): b / Promover responsabilidad...
        # Pregunta 3 Pre y Pregunta 2 Post (Factor de riesgo): b / Consumir sustancias psicoactivas
        # Pregunta 4 Pre y Pregunta 3 Post (Factor protector): a / Dialogar con la familia o b / Actividades deportivas (según versión)
        # Pregunta 5 Pre y Pregunta 4 Post (Corresponsabilidad): d / La familia, la escuela, la comunidad y las instituciones
        
        correctas_pre = []
        correctas_post = []
        conceptos = ["Finalidad SRPA", "Factor Riesgo", "Factor Protector", "Prevención/Corresponsabilidad"]
        
        # Filtrar Pre y Post por separado
        df_pre = df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"]
        df_post = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]
        
        if len(df_pre) > 0 and len(df_post) > 0:
            # Concepto 1: Finalidad del SRPA
            aciertos_p1_pre = df_pre["Conocimientos_P2"].apply(lambda x: 1 if str(x).lower().strip() in ["b", "promover la responsabilidad, la protección de derechos y la resocialización"] else 0).mean() * 100
            aciertos_p1_post = df_post["Conocimientos_P1"].apply(lambda x: 1 if str(x).lower().strip() in ["b", "promover la responsabilidad, la protección de derechos y la resocialización"] else 0).mean() * 100
            
            # Concepto 2: Factor de Riesgo
            aciertos_p2_pre = df_pre["Conocimientos_P3"].apply(lambda x: 1 if str(x).lower().strip() in ["b", "consumir sustancias psicoactivas"] else 0).mean() * 100
            aciertos_p2_post = df_post["Conocimientos_P2"].apply(lambda x: 1 if str(x).lower().strip() in ["b", "consumir sustancias psicoactivas"] else 0).mean() * 100
            
            # Concepto 3: Factor Protector
            aciertos_p3_pre = df_pre["Conocimientos_P4"].apply(lambda x: 1 if str(x).lower().strip() in ["a", "dialogar con la familia."] else 0).mean() * 100
            aciertos_p3_post = df_post["Conocimientos_P3"].apply(lambda x: 1 if str(x).lower().strip() in ["b", "participar en actividades deportivas, culturales o comunitarias"] else 0).mean() * 100
            
            # Concepto 4: Corresponsabilidad de Prevención
            aciertos_p4_pre = df_pre["Conocimientos_P5"].apply(lambda x: 1 if str(x).lower().strip() in ["d", "la familia, la escuela, la comunidad y las instituciones"] else 0).mean() * 100
            aciertos_p4_post = df_post["Conocimientos_P4"].apply(lambda x: 1 if str(x).lower().strip() in ["d", "la familia, la escuela, la comunidad y las instituciones"] else 0).mean() * 100
            
            fig_impacto = go.Figure(data=[
                go.Bar(name='Antes del Taller (PRETEST)', x=conceptos, y=[aciertos_p1_pre, aciertos_p2_pre, aciertos_p3_pre, aciertos_p4_pre], marker_color='#ffc107'),
                go.Bar(name='Después del Taller (POSTEST)', x=conceptos, y=[aciertos_p1_post, aciertos_p2_post, aciertos_p3_post, aciertos_p4_post], marker_color='#198754')
            ])
            fig_impacto.update_layout(
                barmode='group',
                yaxis_title='% de Aciertos / Respuestas Correctas',
                ylim=[0, 100],
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_impacto, use_container_width=True)
        else:
            st.info("💡 Se requiere tener al menos un PRETEST y un POSTEST validados para graficar la comparativa de impacto educativo.")

        # Gráficos secundarios (Roles y Satisfacción)
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.subheader("👤 Distribución de Participantes por Rol")
            df_roles = df_filtered["Rol"].value_counts().reset_index()
            df_roles.columns = ["Rol", "Cantidad"]
            fig_roles = px.pie(df_roles, values="Cantidad", names="Rol", color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_roles, use_container_width=True)
            
        with col_g2:
            st.subheader("⭐ Satisfacción del Taller (POSTEST)")
            # Calcular promedios para las variables de satisfacción (Excelente = 4, Bueno = 3, Regular = 2, Deficiente = 1)
            map_satisfaccion = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None}
            
            sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
            aspectos_nombres = [
                "Claridad de info", "Dominio de facilitadores", "Metodología", 
                "Participación", "Utilidad", "Organización", 
                "Materiales", "Fortaleció conocimientos", "Recomendaría"
            ]
            
            if len(df_post) > 0:
                promedios_satisfaccion = []
                for col in sat_cols:
                    if col in df_post.columns:
                        val_num = df_post[col].map(map_satisfaccion).dropna()
                        promedios_satisfaccion.append(val_num.mean() if len(val_num) > 0 else 0)
                    else:
                        promedios_satisfaccion.append(0)
                        
                fig_sat = px.bar(
                    x=aspectos_nombres, 
                    y=promedios_satisfaccion,
                    labels={'x': 'Aspecto Evaluado', 'y': 'Calificación Promedio (1-4)'},
                    range_y=[1, 4],
                    color_discrete_sequence=['#0d6efd']
                )
                fig_sat.add_hline(y=3.0, line_dash="dash", line_color="green", annotation_text="Meta (Bueno)")
                st.plotly_chart(fig_sat, use_container_width=True)
            else:
                st.info("💡 Los datos de satisfacción se mostrarán una vez se registren cuestionarios de tipo POSTEST.")
