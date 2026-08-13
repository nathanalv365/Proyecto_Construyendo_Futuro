import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import json
import os
from datetime import datetime

# Intentar importar librerías opcionales para Google Sheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# Intentar importar la librería de Gemini
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# Configuración de Página de Streamlit (Modo Móvil Adaptivo por defecto)
st.set_page_config(
    page_title="SRPA - Construyendo Futuro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOCAL_EXCEL_PATH = "plantilla_encuestas_srpa.xlsx"

# -----------------------------------------------------------------------------
# 1. CAPA DE DATOS (CONECTOR DE BASE DE DATOS LOCAL EXCEL O GOOGLE SHEETS)
# -----------------------------------------------------------------------------

class DatabaseConnector:
    """Clase encargada de abstraer si guardamos en Google Sheets o en Excel local."""
    def __init__(self):
        self.use_sheets = False
        self.sheet_client = None
        self.spreadsheet = None
        self.init_connection()

    def init_connection(self):
        # Intentar conectar con Google Sheets si las credenciales están en los secretos
        if GSPREAD_AVAILABLE and "gcp_service_account" in st.secrets:
            try:
                scope = [
                    'https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive'
                ]
                creds_dict = dict(st.secrets["gcp_service_account"])
                # Corregir saltos de línea en la llave privada
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
                
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                self.sheet_client = gspread.authorize(creds)
                
                # Intentar abrir el libro de Google Sheets
                # Debe existir un Google Sheets compartido con este correo electrónico de servicio
                self.spreadsheet = self.sheet_client.open("Base_Encuestas_SRPA")
                self.use_sheets = True
            except Exception as e:
                st.sidebar.warning(f"No se pudo conectar a Google Sheets: {e}. Usando Excel local.")
                self.use_sheets = False
        else:
            self.use_sheets = False
            
        # Si no se usa Google Sheets, asegurar existencia de Excel local estructurado
        if not self.use_sheets:
            self.ensure_local_excel_structure()

    def ensure_local_excel_structure(self):
        """Crea el archivo Excel con las pestañas necesarias si no existe."""
        if not os.path.exists(LOCAL_EXCEL_PATH):
            cola_cols = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"]
            resp_cols = [
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
                "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                "Verificado_Por", "Fecha_Aprobacion"
            ]
            
            # Crear datos demo
            demo_respuestas = [
                ["CF-001", "PRETEST", "2026-07-30", "Cartagena", "Promesa de Dios", "Estudiante", "a", "b", "b", "a", "d", "a", "a", "a", "", "", "", "", "", "", "", "", "", "Coordinador Campo", "2026-08-11"],
                ["CF-002", "POSTEST", "2026-07-30", "Cartagena", "Promesa de Dios", "Estudiante", "", "b", "b", "b", "d", "a", "", "", "Excelente", "Excelente", "Bueno", "Excelente", "Excelente", "Bueno", "Excelente", "Excelente", "Excelente", "Coordinador Campo", "2026-08-11"],
                ["CF-003", "PRETEST", "2026-07-30", "Cartagena", "Promesa de Dios", "Docente", "b", "a", "b", "c", "a", "b", "c", "b", "", "", "", "", "", "", "", "", "", "Coordinador Campo", "2026-08-11"],
                ["CF-004", "POSTEST", "2026-07-30", "Cartagena", "Promesa de Dios", "Docente", "", "b", "b", "b", "d", "a", "", "", "Excelente", "Excelente", "Excelente", "Excelente", "Excelente", "Excelente", "Excelente", "Excelente", "Excelente", "Coordinador Campo", "2026-08-11"],
                ["CF-005", "PRETEST", "2026-08-01", "Turbaco", "I.E. Turbaco", "Padre de Familia", "c", "d", "b", "a", "d", "d", "a", "c", "", "", "", "", "", "", "", "", "", "Coordinador Campo", "2026-08-12"],
                ["CF-006", "POSTEST", "2026-08-01", "Turbaco", "I.E. Turbaco", "Padre de Familia", "", "b", "b", "b", "d", "a", "", "", "Bueno", "Excelente", "Bueno", "Bueno", "Excelente", "Excelente", "Bueno", "Excelente", "Excelente", "Coordinador Campo", "2026-08-12"],
            ]
            
            demo_cola = [
                ["CF-PEND-01", "2026-08-13", "POSTEST", "Cartagena", "Pr0mesa de Di0s", "Estudiante", '{"p1":"", "p2":"b", "p3":"b", "p4":"b", "p5":"d", "s1":"Excelente", "s2":"Bueno", "s3":"Excelente", "s4":"Excelente", "s5":"Bueno", "s6":"Excelente", "s7":"Bueno", "s8":"Excelente", "s9":"Excelente"}', "Pendiente"],
                ["CF-PEND-02", "2026-08-13", "PRETEST", "Turbaco", "I.E. Turbac0", "Líder comunitario", '{"p1":"a", "p2":"b", "p3":"b", "p4":"a", "p5":"d", "p6":"a", "p7":"a", "p8":"ICBF"}', "Pendiente"]
            ]
            
            with pd.ExcelWriter(LOCAL_EXCEL_PATH, engine="openpyxl") as writer:
                pd.DataFrame(demo_cola, columns=cola_cols).to_excel(writer, sheet_name="Cola_Revision", index=False)
                pd.DataFrame(demo_respuestas, columns=resp_cols).to_excel(writer, sheet_name="Respuestas_SRPA", index=False)
            st.sidebar.info("Plantilla de Excel local creada de forma automática.")

    def read_sheet(self, sheet_name) -> pd.DataFrame:
        """Lee datos desde Google Sheets o Excel local según configuración."""
        if self.use_sheets:
            try:
                ws = self.spreadsheet.worksheet(sheet_name)
                records = ws.get_all_records()
                return pd.DataFrame(records)
            except Exception as e:
                st.error(f"Error leyendo de Google Sheets ({sheet_name}): {e}. Intentando Excel local.")
                return pd.read_excel(LOCAL_EXCEL_PATH, sheet_name=sheet_name)
        else:
            try:
                return pd.read_excel(LOCAL_EXCEL_PATH, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error leyendo de Excel local ({sheet_name}): {e}")
                return pd.DataFrame()

    def append_row(self, sheet_name, row_data: list):
        """Inserta una fila al final de la pestaña correspondiente."""
        if self.use_sheets:
            try:
                ws = self.spreadsheet.worksheet(sheet_name)
                ws.append_row(row_data)
            except Exception as e:
                st.error(f"Error guardando en Google Sheets ({sheet_name}): {e}. Guardando localmente.")
                self.append_row_local(sheet_name, row_data)
        else:
            self.append_row_local(sheet_name, row_data)

    def append_row_local(self, sheet_name, row_data: list):
        """Inserta una fila en el Excel local."""
        try:
            wb = pd.ExcelFile(LOCAL_EXCEL_PATH)
            sheets = {s: wb.parse(s) for s in wb.sheet_names}
            
            # Agregar la fila
            new_row_df = pd.DataFrame([row_data], columns=sheets[sheet_name].columns)
            sheets[sheet_name] = pd.concat([sheets[sheet_name], new_row_df], ignore_index=True)
            
            with pd.ExcelWriter(LOCAL_EXCEL_PATH, engine="openpyxl") as writer:
                for s, df in sheets.items():
                    df.to_excel(writer, sheet_name=s, index=False)
        except Exception as e:
            st.error(f"Error escribiendo en Excel local ({sheet_name}): {e}")

    def update_cola_status_to_approved(self, id_encuesta):
        """Marca una encuesta en la cola de revisión como Aprobada."""
        if self.use_sheets:
            try:
                ws = self.spreadsheet.worksheet("Cola_Revision")
                cell = ws.find(id_encuesta)
                if cell:
                    # La columna Estado es la 8
                    ws.update_cell(cell.row, 8, "Aprobado")
            except Exception as e:
                st.error(f"Error actualizando estado en Google Sheets: {e}. Aplicando cambio local.")
                self.update_cola_status_to_approved_local(id_encuesta)
        else:
            self.update_cola_status_to_approved_local(id_encuesta)

    def update_cola_status_to_approved_local(self, id_encuesta):
        """Marca una encuesta en la cola de revisión de Excel como Aprobada."""
        try:
            wb = pd.ExcelFile(LOCAL_EXCEL_PATH)
            sheets = {s: wb.parse(s) for s in wb.sheet_names}
            
            df_cola = sheets["Cola_Revision"]
            df_cola.loc[df_cola["ID_Encuesta"] == id_encuesta, "Estado"] = "Aprobado"
            sheets["Cola_Revision"] = df_cola
            
            with pd.ExcelWriter(LOCAL_EXCEL_PATH, engine="openpyxl") as writer:
                for s, df in sheets.items():
                    df.to_excel(writer, sheet_name=s, index=False)
        except Exception as e:
            st.error(f"Error actualizando cola local: {e}")

# -----------------------------------------------------------------------------
# 2. CONFIGURACIÓN DEL ENGINE DE OCR (GEMINI API)
# -----------------------------------------------------------------------------

def inicializar_gemini_api(api_key_custom=None):
    """Inicializa el SDK de Gemini utilizando una clave de la UI o de los Secretos."""
    api_key = api_key_custom or st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        return False
    
    try:
        genai.configure(api_key=api_key)
        return True
    except Exception:
        return False

# Prompt estructurado de análisis para encuestas físicas del Proyecto Construyendo Futuro
PROMPT_ANALISIS_OCR = """
Eres el sistema OCR inteligente oficial del "Proyecto Construyendo Futuro". Tu objetivo es leer y digitalizar de manera precisa el contenido manuscrito y las marcas de verificación de las dos fotos correspondientes a una encuesta física (Página 1 y Página 2).

INSTRUCCIONES CLAVE:
1. Analiza de manera conjunta las dos imágenes provistas. Ambas pertenecen al mismo participante.
2. Identifica el tipo de formulario: "PRETEST" o "POSTEST" según se titule en la cabecera.
3. Extrae la información manuscrita e identificativa de la Página 1:
   - Fecha (escrita en formato manuscrito en la parte superior).
   - Municipio (generalmente Cartagena, Turbaco, etc. escrito a puño y letra).
   - Institución Educativa (Nombre de la escuela escrito por el facilitador o el participante). ¡Presta extrema atención a la ortografía y legibilidad!
   - Rol del participante (Estudiante, Docente, Padre de Familia, Líder comunitario). Está marcado con una "X" en una tabla.
4. Extrae las respuestas a las preguntas de conocimiento (marcadas con "X"):
   - Si es PRETEST: Hay 8 preguntas de conocimientos (P1 a P8).
   - Si es POSTEST: Hay 5 preguntas de conocimiento en la SECCIÓN A (P1 a P5) y 9 preguntas de satisfacción en la SECCIÓN B (S1 a S9, cuyos valores pueden ser: Excelente, Bueno, Regular, Deficiente).
5. Retorna la información EXCLUSIVAMENTE en un formato JSON plano, estricto, sin decoraciones de código como ```json, de acuerdo a la siguiente estructura:

{
  "tipo_formulario": "PRETEST",
  "fecha": "2026-07-30",
  "municipio": "Turbaco",
  "institucion_educativa": "I.E. Turbaco",
  "rol": "Estudiante",
  "respuestas_conocimiento": {
    "p1": "a",
    "p2": "b",
    "p3": "b",
    "p4": "a",
    "p5": "d",
    "p6": "a",
    "p7": "a",
    "p8": "ICBF"
  },
  "satisfaccion": {
    "s1": "", "s2": "", "s3": "", "s4": "", "s5": "", "s6": "", "s7": "", "s8": "", "s9": ""
  }
}

Si es POSTEST, las respuestas de conocimientos irán en "p1" a "p5" y el objeto "satisfaccion" tendrá las marcas correspondientes ("Excelente", "Bueno", "Regular", "Deficiente"). Si una casilla está en blanco o no es legible, devuélvela vacía.
"""

def analizar_encuestas_con_ia(img_pag1, img_pag2, api_key_custom, simulated_mode):
    """Envía la página 1 y la página 2 conjuntamente a Gemini o simula un resultado realista."""
    if simulated_mode:
        import time
        time.sleep(2)  # Simular latencia de red
        
        # Simulación de extracción realista
        es_postest = np.random.choice([True, False])
        roles_demo = ["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"]
        muni_demo = ["Cartagena", "Turbaco", "Arjona", "El Carmen de Bolívar"]
        ie_demo = ["Promesa de Dios", "I.E. San Juan", "I.E. San José de Turbaco", "I.E. Madre Bernarda"]
        
        tipo = "POSTEST" if es_postest else "PRETEST"
        resp_c = {}
        for i in range(1, 9 if not es_postest else 6):
            resp_c[f"p{i}"] = np.random.choice(["a", "b", "c", "d", ""]) if i not in [1, 7] else np.random.choice(["a", "b", "c"])
            
        sat_c = {}
        if es_postest:
            for s in range(1, 10):
                sat_c[f"s{s}"] = np.random.choice(["Excelente", "Bueno", "Regular", "Deficiente"], p=[0.6, 0.3, 0.08, 0.02])
        else:
            sat_c = {f"s{s}": "" for s in range(1, 10)}
            
        return {
            "tipo_formulario": tipo,
            "fecha": "2026-08-13",
            "municipio": np.random.choice(muni_demo),
            "institucion_educativa": np.random.choice(ie_demo) + " (Sugerido por IA)",
            "rol": np.random.choice(roles_demo),
            "respuestas_conocimiento": resp_c,
            "satisfaccion": sat_c
        }
    
    # Modo API Real
    if not GEMINI_AVAILABLE:
        st.error("La librería `google-generativeai` no está instalada. Ejecuta `pip install google-generativeai`.")
        return None
        
    inicializado = inicializar_gemini_api(api_key_custom)
    if not inicializado:
        st.error("Por favor, configura tu API Key de Gemini en la barra lateral para utilizar el modo real.")
        return None
        
    try:
        # Configurar y mandar las imágenes
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # Convertir imágenes de streamlit (UploadFile) a imágenes PIL legibles por la API de Google
        pil_p1 = Image.open(img_pag1).convert("RGB")
        pil_p2 = Image.open(img_pag2).convert("RGB")
        
        # Opcional: Redimensionar imágenes muy grandes para optimizar costos de ancho de banda y velocidad de tokenización
        pil_p1.thumbnail((1200, 1600))
        pil_p2.thumbnail((1200, 1600))
        
        # Realizar la consulta a Gemini
        response = model.generate_content([
            pil_p1,
            pil_p2,
            PROMPT_ANALISIS_OCR
        ])
        
        # Limpiar la respuesta para asegurar JSON válido
        text_response = response.text.strip()
        if "```json" in text_response:
            text_response = text_response.split("```json")[1].split("```")[0].strip()
        elif "```" in text_response:
            text_response = text_response.split("```")[1].strip()
            
        data = json.loads(text_response)
        return data
    except Exception as e:
        st.error(f"Error procesando con Gemini API: {e}")
        return None

# -----------------------------------------------------------------------------
# 3. INTERFAZ DE USUARIO (STREAMLIT APP)
# -----------------------------------------------------------------------------

# Inicializar conector de base de datos
if 'db' not in st.session_state:
    st.session_state.db = DatabaseConnector()

db = st.session_state.db

# Barra Lateral: Configuración de Entorno e Indicadores
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/e/ea/Gobernaci%C3%B3n_de_Bol%C3%ADvar.jpg", width=120)
st.sidebar.title("Proyecto Construyendo Futuro")
st.sidebar.markdown("**Evaluación de Conocimientos SRPA**  \n*Bolívar Nos Une*")
st.sidebar.divider()

# Indicador de base de datos activa
if db.use_sheets:
    st.sidebar.success("🟢 Conectado a Google Sheets")
else:
    st.sidebar.info("📂 Almacenamiento local: Excel activo")

st.sidebar.divider()

# Configuración de Clave API de Gemini
st.sidebar.subheader("🔑 Inteligencia Artificial")
api_key_custom = st.sidebar.text_input("API Key de Gemini", type="password", help="Genera una clave gratuita en Google AI Studio")
st.sidebar.markdown("[Obtener API Key gratis ↗](https://aistudio.google.com/)")

simulated_mode = True
if api_key_custom or "GEMINI_API_KEY" in st.secrets:
    simulated_mode = st.sidebar.checkbox("Usar modo simulado (demo)", value=False)
else:
    st.sidebar.warning("Usando Modo Simulado debido a falta de API Key.")

# Botones de utilidad
st.sidebar.divider()
if st.sidebar.button("🔄 Recargar datos"):
    st.session_state.db.init_connection()
    st.toast("Datos recargados desde la base de datos central.")

# -----------------------------------------------------------------------------
# TABS PRINCIPALES DE LA APLICACIÓN
# -----------------------------------------------------------------------------
tab_carga, tab_cola, tab_dashboard = st.tabs([
    "📥 Cargar Encuesta (Doble Página)", 
    "✏️ Cola de Revisión (Validar I.E.)", 
    "📊 Dashboard Estadístico"
])

# -----------------------------------------------------------------------------
# TAB 1: CARGAR ENCUESTA (MÓDULO DE FOTOGRAFÍA)
# -----------------------------------------------------------------------------
with tab_carga:
    st.header("📥 Captura y Carga de Encuestas Físicas")
    st.write("Para consolidar una encuesta, debes subir la **Página 1** (que contiene la cabecera e institución) y la **Página 2** (que contiene las preguntas de conocimientos y satisfacción).")
    
    col_input1, col_input2 = st.columns(2)
    
    with col_input1:
        st.markdown("### 📸 Página 1")
        img_p1 = st.file_uploader("Tomar foto o subir Página 1", type=["jpg", "jpeg", "png"], key="p1_upload", help="Asegúrate de que la cabecera y el nombre manuscrito de la institución sean nítidos.")
        if img_p1:
            st.image(img_p1, caption="Página 1 cargada", use_container_width=True)
            
    with col_input2:
        st.markdown("### 📸 Página 2")
        img_p2 = st.file_uploader("Tomar foto o subir Página 2", type=["jpg", "jpeg", "png"], key="p2_upload", help="Asegúrate de capturar bien las marcas de respuestas o de satisfacción.")
        if img_p2:
            st.image(img_p2, caption="Página 2 cargada", use_container_width=True)

    st.divider()
    
    # Acción de procesamiento
    if st.button("🔍 Iniciar Lectura con Inteligencia Artificial", use_container_width=True, type="primary"):
        if not img_p1 or not img_p2:
            st.warning("⚠️ Debes subir obligatoriamente ambas páginas (Página 1 y Página 2) para procesar la encuesta.")
        else:
            with st.spinner("Procesando imágenes con Gemini 1.5 Flash... Extrayendo marcas de casilla y caligrafía manuscrita."):
                datos_extraidos = analizar_encuestas_con_ia(img_p1, img_p2, api_key_custom, simulated_mode)
                
                if datos_extraidos:
                    # Generar ID único temporal
                    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    id_temp = f"CF-PEND-{timestamp}"
                    
                    # Convertir datos
                    tipo_form = datos_extraidos.get("tipo_formulario", "PRETEST")
                    muni = datos_extraidos.get("municipio", "Sin clasificar")
                    ie_ia = datos_extraidos.get("institucion_educativa", "Pendiente lectura")
                    rol = datos_extraidos.get("rol", "Estudiante")
                    
                    # Guardar respuestas estructuradas juntas
                    respuestas_juntas = {
                        "respuestas_conocimiento": datos_extraidos.get("respuestas_conocimiento", {}),
                        "satisfaccion": datos_extraidos.get("satisfaccion", {})
                    }
                    json_resp = json.dumps(respuestas_juntas)
                    
                    # Fila para la cola de revisión
                    nueva_fila_cola = [
                        id_temp,
                        datetime.now().strftime("%Y-%m-%d"),
                        tipo_form,
                        muni,
                        ie_ia,
                        rol,
                        json_resp,
                        "Pendiente"
                    ]
                    
                    db.append_row("Cola_Revision", nueva_fila_cola)
                    
                    st.success(f"🎉 ¡Encuesta leída con éxito! Se ha creado el registro temporal **{id_temp}**.")
                    st.info(f"📍 **I.E. Detectada por IA:** {ie_ia}. Ve al menú superior **'Cola de Revisión'** para verificar la caligrafía manuscrita e ingresarla definitivamente en la base de datos.")

# -----------------------------------------------------------------------------
# TAB 2: COLA DE REVISIÓN (HUMAN-IN-THE-LOOP)
# -----------------------------------------------------------------------------
with tab_cola:
    st.header("✏️ Banco de Revisión Humana: Validar Caligrafía de I.E.")
    st.write("Dado que los nombres de los colegios e instituciones se escriben a puño y letra, revisa aquí las lecturas de la IA antes de insertarlas de manera definitiva en el Excel/Google Sheets consolidado.")
    
    # Leer datos frescos de la cola de revisión
    df_cola_full = db.read_sheet("Cola_Revision")
    
    if df_cola_full.empty:
        st.info("No hay encuestas pendientes de validación en la cola.")
    else:
        # Filtrar solo registros pendientes
        df_cola_pendientes = df_cola_full[df_cola_full["Estado"] == "Pendiente"]
        
        if df_cola_pendientes.empty:
            st.success("✨ ¡Felicidades! Todas las encuestas cargadas han sido verificadas y aprobadas.")
        else:
            st.warning(f"Tienes **{len(df_cola_pendientes)}** encuestas listas para revisar en la cola.")
            
            # Selector de encuesta para revisar
            lista_pendientes = df_cola_pendientes["ID_Encuesta"].tolist()
            encuesta_seleccionada = st.selectbox("Selecciona la encuesta a revisar:", lista_pendientes)
            
            # Obtener datos de la encuesta elegida
            datos_encuesta = df_cola_pendientes[df_cola_pendientes["ID_Encuesta"] == encuesta_seleccionada].iloc[0]
            
            # Parsear respuestas JSON
            try:
                respuestas_dict = json.loads(datos_encuesta["JSON_Respuestas"])
            except Exception:
                respuestas_dict = {"respuestas_conocimiento": {}, "satisfaccion": {}}
                
            st.divider()
            
            # Interfaz de edición lado a lado
            col_edicion_izq, col_edicion_der = st.columns([1, 1])
            
            with col_edicion_izq:
                st.subheader("💡 Información Detectada por la IA")
                st.info(f"**ID Temporal:** {datos_encuesta['ID_Encuesta']}  \n**Fecha de Carga:** {datos_encuesta['Fecha_Carga']}")
                
                # Campos de edición para control humano
                ie_editada = st.text_input("🏫 Institución Educativa (Corregir caligrafía):", value=datos_encuesta["Institucion_Educativa_IA"])
                muni_editado = st.text_input("📍 Municipio:", value=datos_encuesta["Municipio"])
                rol_editado = st.selectbox("👤 Rol de Participante:", ["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"].index(datos_encuesta["Rol"]) if datos_encuesta["Rol"] in ["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"] else 0)
                tipo_form_editado = st.selectbox("📝 Tipo de Formulario:", ["PRETEST", "POSTEST"], index=0 if datos_encuesta["Tipo_Formulario"] == "PRETEST" else 1)
                
            with col_edicion_der:
                st.subheader("🔎 Detalle de Respuestas Extraídas")
                st.write("**Sección Conocimientos (Marcadas):**")
                st.write(respuestas_dict.get("respuestas_conocimiento", {}))
                
                if datos_encuesta["Tipo_Formulario"] == "POSTEST":
                    st.write("**Sección Satisfacción:**")
                    st.write(respuestas_dict.get("satisfaccion", {}))
                
                st.divider()
                
                revisor_nombre = st.text_input("💻 Nombre del Revisor (Responsable):", value="Coordinador de Campo")
                
                # BOTÓN CRÍTICO DE APROBACIÓN
                if st.button("✅ Aprobar e Insertar en Base de Datos Final", use_container_width=True, type="primary"):
                    # Extraer respuestas de conocimiento ordenadas (p1 a p8)
                    rc = respuestas_dict.get("respuestas_conocimiento", {})
                    con_p1 = rc.get("p1", "")
                    con_p2 = rc.get("p2", "")
                    con_p3 = rc.get("p3", "")
                    con_p4 = rc.get("p4", "")
                    con_p5 = rc.get("p5", "")
                    con_p6 = rc.get("p6", "")
                    con_p7 = rc.get("p7", "")
                    con_p8 = rc.get("p8", "")
                    
                    # Extraer respuestas de satisfacción ordenadas (s1 a s9)
                    rs = respuestas_dict.get("satisfaccion", {})
                    sat_p1 = rs.get("s1", "")
                    sat_p2 = rs.get("s2", "")
                    sat_p3 = rs.get("s3", "")
                    sat_p4 = rs.get("s4", "")
                    sat_p5 = rs.get("s5", "")
                    sat_p6 = rs.get("s6", "")
                    sat_p7 = rs.get("s7", "")
                    sat_p8 = rs.get("s8", "")
                    sat_p9 = rs.get("s9", "")
                    
                    # Construir la fila final consolidada
                    fila_consolidada = [
                        datos_encuesta["ID_Encuesta"],
                        tipo_form_editado,
                        datetime.now().strftime("%Y-%m-%d"),
                        muni_editado,
                        ie_editada,
                        rol_editado,
                        con_p1, con_p2, con_p3, con_p4, con_p5, con_p6, con_p7, con_p8,
                        sat_p1, sat_p2, sat_p3, sat_p4, sat_p5, sat_p6, sat_p7, sat_p8, sat_p9,
                        revisor_nombre,
                        datetime.now().strftime("%Y-%m-%d")
                    ]
                    
                    # 1. Guardar en Respuestas_SRPA
                    db.append_row("Respuestas_SRPA", fila_consolidada)
                    
                    # 2. Actualizar estado en la Cola
                    db.update_cola_status_to_approved(datos_encuesta["ID_Encuesta"])
                    
                    st.success(f"🚀 ¡Excelente! El registro `{datos_encuesta['ID_Encuesta']}` de `{ie_editada}` ha sido verificado y aprobado con éxito en la base de datos.")
                    
                    # Esperar y recargar
                    st.rerun()

# -----------------------------------------------------------------------------
# TAB 3: DASHBOARD ESTADÍSTICO INTERACTIVO (TIEMPO REAL)
# -----------------------------------------------------------------------------
with tab_dashboard:
    st.header("📊 Panel Estadístico del Impacto Educativo (SRPA)")
    st.write("Visualiza en tiempo real los indicadores de avance, efectividad en el aprendizaje y satisfacción de las jornadas de prevención del ingreso al Sistema de Responsabilidad Penal.")
    
    # Cargar datos aprobados
    df_resp = db.read_sheet("Respuestas_SRPA")
    
    if df_resp.empty:
        st.warning("No hay registros aprobados en la base de datos para generar estadísticas.")
    else:
        # Reemplazar valores nulos para evitar fallas en gráficos
        df_resp = df_resp.fillna("")
        
        # Filtros dinámicos en la parte superior del Dashboard
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1:
            municipios_disponibles = ["Todos"] + sorted(df_resp["Municipio"].unique().tolist())
            muni_filtro = st.selectbox("Filtrar por Municipio:", municipios_disponibles)
        with col_filtro2:
            ie_disponibles = ["Todas"] + sorted(df_resp["Institucion_Educativa_Verificada"].unique().tolist())
            ie_filtro = st.selectbox("Filtrar por Institución Educativa:", ie_disponibles)
            
        # Aplicar filtros
        df_filtrado = df_resp.copy()
        if muni_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Municipio"] == muni_filtro]
        if ie_filtro != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Institucion_Educativa_Verificada"] == ie_filtro]
            
        st.divider()
        
        # 1. TARJETAS DE MÉTRICAS CLAVE (KPIs)
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        with kpi1:
            st.metric(label="✅ Encuestas Consolidadas", value=len(df_filtrado))
        with kpi2:
            pretests_count = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"])
            st.metric(label="📋 Total Pretests", value=pretests_count)
        with kpi3:
            postests_count = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"])
            st.metric(label="📝 Total Postests", value=postests_count)
        with kpi4:
            # Conteo de pendientes desde la base de datos de cola
            df_cola_fresh = db.read_sheet("Cola_Revision")
            pendientes_reales = len(df_cola_fresh[df_cola_fresh["Estado"] == "Pendiente"]) if not df_cola_fresh.empty else 0
            st.metric(label="⏳ Pendientes en Cola", value=pendientes_reales)
            
        st.divider()
        
        # 2. COMPARATIVA DE APRENDIZAJE: PRETEST VS POSTEST (MEDICIÓN DE IMPACTO)
        st.subheader("💡 Efectividad del Taller: % Respuestas Correctas (Antes vs Después)")
        st.write("Compara el nivel de asimilación de conceptos clave midiendo el acierto en las preguntas idénticas compartidas entre el Pretest y el Postest.")
        
        # Filtrar Pretests y Postests para el cálculo
        df_pre = df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"]
        df_post = df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"]
        
        if df_pre.empty or df_post.empty:
            st.info("Para visualizar la comparación de impacto educativo (Antes vs Después) se requiere tener registrado al menos un PRETEST y un POSTEST con el filtro actual.")
        else:
            # Calcular porcentajes de acierto por pregunta clave
            # Concepto 1: Finalidad del SRPA (Pretest Q2 == 'b', Postest Q1 == 'b')
            pre_c1 = (df_pre["Conocimientos_P2"] == "b").mean() * 100
            post_c1 = (df_post["Conocimientos_P1"] == "b").mean() * 100
            
            # Concepto 2: Factor de Riesgo (Pretest Q3 == 'b', Postest Q2 == 'b')
            pre_c2 = (df_pre["Conocimientos_P3"] == "b").mean() * 100
            post_c2 = (df_post["Conocimientos_P2"] == "b").mean() * 100
            
            # Concepto 3: Factor Protector (Pretest Q4 == 'a', Postest Q3 == 'b')
            pre_c3 = (df_pre["Conocimientos_P4"] == "a").mean() * 100
            post_c3 = (df_post["Conocimientos_P3"] == "b").mean() * 100
            
            # Concepto 4: Responsable de Prevención (Pretest Q5 == 'd', Postest Q4 == 'd')
            pre_c4 = (df_pre["Conocimientos_P5"] == "d").mean() * 100
            post_c4 = (df_post["Conocimientos_P4"] == "d").mean() * 100
            
            # Estructurar datos para gráfico comparativo de barras agrupadas
            conceptos = [
                "Finalidad SRPA",
                "Factor Riesgo (SPA)",
                "Factor Protector",
                "Responsabilidad Prevención"
            ]
            
            fig_impacto = go.Figure(data=[
                go.Bar(name='Antes del Taller (Pretest)', x=conceptos, y=[pre_c1, pre_c2, pre_c3, pre_c4], marker_color='#9E9E9E'),
                go.Bar(name='Después del Taller (Postest)', x=conceptos, y=[post_c1, post_c2, post_c3, post_c4], marker_color='#1E3D59')
            ])
            
            fig_impacto.update_layout(
                barmode='group',
                yaxis_title='% de Respuestas Correctas',
                yaxis=dict(ticksuffix="%", range=[0, 105]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(l=40, r=40, t=50, b=40)
            )
            
            st.plotly_chart(fig_impacto, use_container_width=True)
            
        st.divider()
        
        # 3. ROLES Y EVALUACIÓN DE SATISFACCIÓN LADO A LADO
        col_graf1, col_graf2 = st.columns(2)
        
        with col_graf1:
            st.subheader("👥 Perfil de los Participantes")
            df_roles = df_filtrado["Rol"].value_counts().reset_index()
            df_roles.columns = ["Rol", "Cantidad"]
            
            fig_roles = px.pie(
                df_roles, 
                values="Cantidad", 
                names="Rol", 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            fig_roles.update_layout(
                margin=dict(l=20, r=20, t=20, b=20),
                legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_roles, use_container_width=True)
            
        with col_graf2:
            st.subheader("⭐ Satisfacción de la Jornada (Escala 1 a 4)")
            if df_post.empty:
                st.info("No hay datos de Postests para mostrar evaluación de satisfacción.")
            else:
                # Mapear respuestas cualitativas a escala numérica para calcular promedio
                satisfaccion_mapping = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": np.nan}
                
                sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
                df_sat_numeric = df_post[sat_cols].copy()
                
                for c in sat_cols:
                    df_sat_numeric[c] = df_sat_numeric[c].map(satisfaccion_mapping)
                    
                promedios_sat = df_sat_numeric.mean().values
                
                aspectos = [
                    "1. Claridad Información",
                    "2. Dominio Facilitadores",
                    "3. Metodología",
                    "4. Participación",
                    "5. Utilidad Temas",
                    "6. Organización",
                    "7. Materiales",
                    "8. Fortaleció Conocimientos",
                    "9. Recomendaría Jornada"
                ]
                
                # Crear gráfico de barras horizontal de promedios
                fig_sat = go.Figure()
                fig_sat.add_trace(go.Bar(
                    y=aspectos,
                    x=promedios_sat,
                    orientation='h',
                    marker_color='#17B890',
                    name="Calificación Promedio"
                ))
                
                # Añadir línea de meta mínima aceptable (e.g. Bueno = 3.0)
                fig_sat.add_vline(x=3.0, line_dash="dash", line_color="red", annotation_text="Meta Mínima (Bueno)")
                
                fig_sat.update_layout(
                    xaxis_title="Puntaje Promedio (Excelente=4, Bueno=3)",
                    xaxis=dict(range=[1, 4.1]),
                    yaxis=dict(autorange="reversed"),
                    margin=dict(l=20, r=20, t=20, b=20)
                )
                st.plotly_chart(fig_sat, use_container_width=True)

        st.divider()
        
        # 4. TABLA DE REGISTROS CONSOLIDADOS Y EXPORTACIÓN
        st.subheader("📥 Exportación de Registros Consolidados")
        st.write("Visualiza la base de datos de encuestas procesadas y aprobadas. Puedes descargar este archivo corregido en cualquier momento en formato CSV.")
        
        st.dataframe(df_filtrado, use_container_width=True)
        
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="💾 Descargar Base de Datos Corregida (CSV)",
            data=csv_data,
            file_name="base_datos_srpa_consolidada.csv",
            mime="text/csv",
            use_container_width=True
        )
