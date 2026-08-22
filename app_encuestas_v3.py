import streamlit as st
import pandas as pd
import json
import base64
import time
from datetime import datetime
from PIL import Image
import io

# Setup page config
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        color: #1e3a8a;
        font-family: 'Helvetica Neue', Arial, sans-serif;
        font-weight: 700;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-header {
        color: #4b5563;
        font-family: 'Helvetica Neue', Arial, sans-serif;
        text-align: center;
        margin-bottom: 25px;
        font-size: 1.1rem;
    }
    .card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #e5e7eb;
        margin-bottom: 20px;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 600;
    }
    .status-ok {
        background-color: #ecfdf5;
        border: 1px solid #10b981;
        padding: 15px;
        border-radius: 8px;
        color: #065f46;
        font-weight: 500;
        margin-bottom: 15px;
    }
    .status-error {
        background-color: #fef2f2;
        border: 1px solid #ef4444;
        padding: 15px;
        border-radius: 8px;
        color: #991b1b;
        font-weight: 500;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Try importing Google API client and GSheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    # Google GenAI support
    try:
        from google import genai
        from google.genai import types
        USE_NEW_SDK = True
    except ImportError:
        import google.generativeai as genai_legacy
        USE_NEW_SDK = False
    GSHEETS_AVAILABLE = True
except ImportError as e:
    GSHEETS_AVAILABLE = False
    st.error(f"Faltan dependencias críticas de Google: {e}. Asegúrese de tener 'gspread', 'oauth2client' y 'google-genai' en requirements.txt")

# Helper to load and initialize Google Sheets connection
def get_gspread_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.sidebar.error(f"Error de autenticación con Google: {e}")
        return None

def init_google_sheets():
    client = get_gspread_client()
    if not client:
        return None, None
    try:
        try:
            spreadsheet = client.open("Base_Encuestas_SRPA")
        except gspread.exceptions.SpreadsheetNotFound:
            # Create sheet if not exists
            spreadsheet = client.create("Base_Encuestas_SRPA")
            # Share with client email if needed or just let it exist
            st.sidebar.info("Base de datos creada en Google Sheets.")

        # Ensure Cola_Revision worksheet exists
        try:
            ws_cola = spreadsheet.worksheet("Cola_Revision")
        except gspread.exceptions.WorksheetNotFound:
            ws_cola = spreadsheet.add_worksheet(title="Cola_Revision", rows="1000", cols="10")
            headers = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Img_Pag1_Base64", "Img_Pag2_Base64"]
            ws_cola.update('A1:J1', [headers])

        # Ensure Respuestas_SRPA worksheet exists
        try:
            ws_resp = spreadsheet.worksheet("Respuestas_SRPA")
        except gspread.exceptions.WorksheetNotFound:
            ws_resp = spreadsheet.add_worksheet(title="Respuestas_SRPA", rows="2000", cols="25")
            headers = [
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
                "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                "Verificado_Por", "Fecha_Aprobacion"
            ]
            ws_resp.append_row(headers)

        return ws_cola, ws_resp
    except Exception as e:
        st.sidebar.error(f"Error conectando a las pestañas de Google Sheets: {e}")
        return None, None

# Connect sheets
ws_cola, ws_resp = init_google_sheets()

# Helper to parse sheet values to DataFrame safely
def load_worksheet_to_df(ws, expected_cols):
    if not ws:
        return pd.DataFrame(columns=expected_cols)
    try:
        raw_rows = ws.get_all_values()
        if not raw_rows or len(raw_rows) <= 1:
            return pd.DataFrame(columns=expected_cols)
        
        headers = [h.strip() for h in raw_rows[0]]
        data_rows = raw_rows[1:]
        
        # Create dictionary representation matching headers
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Ensure all expected columns are present
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""
                
        return df[expected_cols]
    except Exception as e:
        st.error(f"Error cargando los datos desde Google Sheets: {e}")
        return pd.DataFrame(columns=expected_cols)

# Define column specifications
COLS_COLA = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Img_Pag1_Base64", "Img_Pag2_Base64"]
COLS_RESP = [
    "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
    "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
    "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
    "Verificado_Por", "Fecha_Aprobacion"
]

# Google Sheets write verification function
def append_row_and_verify(sheet, row_data, id_col_index, id_value):
    """
    Inserta un registro en Google Sheets y verifica inmediatamente su existencia y consistencia.
    """
    if not sheet:
        return False, "La conexión a Google Sheets no está disponible."
    try:
        # Append row
        sheet.append_row(row_data)
        
        # Wait briefly for propagation
        time.sleep(1.5)
        
        # Verification: search the sheet for the unique ID
        cells = sheet.findall(id_value)
        if not cells:
            return False, f"Verificación fallida: El registro con ID {id_value} no se encontró en la base de datos tras la inserción."
            
        # Check that the row was written in the correct column
        matching_row_index = None
        for cell in cells:
            if cell.col == id_col_index:
                matching_row_index = cell.row
                break
                
        if matching_row_index is None:
            return False, f"Verificación fallida: El ID {id_value} se encontró en una columna incorrecta (columna {cell.col})."
            
        # Load row back and compare critical metadata
        retrieved_row = sheet.row_values(matching_row_index)
        
        # Verificar la coincidencia del ID para garantizar la integridad de la fila
        if not retrieved_row:
            return False, "La fila recuperada de Google Sheets está vacía o no se pudo leer."
        ret_id = str(retrieved_row[0]).strip()
        if ret_id != str(id_value).strip():
            return False, f"Fallo de integridad: El ID guardado '{ret_id}' no coincide con el ID enviado '{id_value}'."
                
        return True, f"Registro {id_value} guardado y verificado correctamente en la fila {matching_row_index}."
    except Exception as e:
        return False, f"Error durante la inserción o verificación: {e}"

# Update row status or values in sheet with verification
def update_row_and_verify(sheet, id_value, update_col_name, new_value, expected_cols_list):
    """
    Actualiza una celda específica buscando por ID de encuesta y verifica que el cambio sea persistente.
    """
    if not sheet:
        return False, "La conexión a Google Sheets no está disponible."
    try:
        raw_rows = sheet.get_all_values()
        if not raw_rows:
            return False, "La hoja de cálculo está vacía."
            
        headers = [h.strip() for h in raw_rows[0]]
        if update_col_name not in headers:
            return False, f"Columna '{update_col_name}' no encontrada en la hoja."
            
        col_idx = headers.index(update_col_name) + 1
        
        # Find row index by ID
        matching_row_index = None
        for r_idx, row in enumerate(raw_rows[1:], start=2):
            if row[0].strip() == id_value.strip():
                matching_row_index = r_idx
                break
                
        if matching_row_index is None:
            return False, f"No se encontró el registro con ID {id_value} para actualizar."
            
        # Perform update
        sheet.update_cell(matching_row_index, col_idx, new_value)
        
        # Verify
        time.sleep(1.0)
        updated_value = sheet.cell(matching_row_index, col_idx).value
        if str(updated_value).strip() != str(new_value).strip():
            return False, f"Error de persistencia: Se intentó actualizar a '{new_value}' pero la celda devolvió '{updated_value}'."
            
        return True, "Actualización completada y verificada."
    except Exception as e:
        return False, f"Error actualizando o verificando celda: {e}"

# Compression helper for Base64 storage
def compress_and_base64(uploaded_file, max_size=(350, 350)):
    if uploaded_file is None:
        return ""
    try:
        img = Image.open(uploaded_file)
        img.thumbnail(max_size)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=75)
        img_bytes = buffer.getvalue()
        return base64.b64encode(img_bytes).decode("utf-8")
    except Exception as e:
        st.warning(f"No se pudo comprimir la imagen: {e}")
        return ""

# Sidebar instructions and credentials panel
with st.sidebar:
    st.image("https://www.bolivar.gov.co/images/Gobernacion/Logo-Gobernacion-Color.png", use_container_width=True)
    st.markdown("### Proyecto Construyendo Futuro")
    st.markdown("Sistematización de Encuestas SRPA con Verificación de Escritura Humana.")
    
    # Check GSheets credentials status
    if "gcp_service_account" in st.secrets:
        st.success("🔌 Google Sheets: Conectado")
    else:
        st.warning("⚠️ Google Sheets: No Conectado")
        with st.expander("🔑 Guía de Configuración"):
            st.info("Pega tus credenciales JSON de la cuenta de servicio de Google Cloud en los Secretos de Streamlit Cloud con el nombre '[gcp_service_account]'.")

    # Check Gemini API Key
    if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
        st.success("🤖 Google Gemini API: Activo")
    else:
        st.error("❌ Google Gemini API: Faltante")
        st.info("Agregue la variable 'GEMINI_API_KEY' en los Secretos de Streamlit Cloud para habilitar el OCR inteligente.")

# Main title
st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Evaluación de Conocimientos y Satisfacción SRPA - Gobernación de Bolívar</p>", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3 = st.tabs(["📤 Carga de Encuestas", "🔍 Banco de Verificación", "📊 Dashboard Estadístico"])

# --- TAB 1: CARGA DE ENCUESTAS ---
with tab1:
    st.markdown("### Cargar Encuesta de Doble Página")
    st.markdown("Suba las fotografías de la Página 1 (Cabecera y preguntas 1-3) y la Página 2 (Preguntas 4-8 y Matriz de Satisfacción) del mismo participante.")

    col1, col2 = st.columns(2)
    with col1:
        file_pag1 = st.file_uploader("Subir Imagen Página 1 (Obligatorio)", type=["jpg", "png", "jpeg"], key="uploader_p1")
    with col2:
        file_pag2 = st.file_uploader("Subir Imagen Página 2 (Obligatorio)", type=["jpg", "png", "jpeg"], key="uploader_p2")

    municipio_input = st.selectbox("Seleccione Municipio del Taller", ["Cartagena", "Turbaco", "Arjona", "Carmen de Bolívar", "Magangué", "María La Baja"])

    if st.button("🚀 Procesar e Inserción con IA", key="btn_procesar"):
        if not file_pag1 or not file_pag2:
            st.error("Debe subir obligatoriamente ambas páginas (Página 1 y Página 2) para procesar el cuestionario de manera integrada.")
        elif "GEMINI_API_KEY" not in st.secrets:
            st.error("Por favor, ingrese su GEMINI_API_KEY en los secretos de la aplicación para habilitar el motor OCR.")
        elif not ws_cola:
            st.error("No se puede guardar el registro debido a que Google Sheets no está correctamente conectado. Verifique sus secretos 'gcp_service_account'.")
        else:
            with st.spinner("Analizando imágenes y extrayendo caligrafía manuscrita con Gemini..."):
                try:
                    # Convert images to PIL
                    img1 = Image.open(file_pag1)
                    img2 = Image.open(file_pag2)
                    
                    # Call Gemini API with Fallback Models
                    api_key = st.secrets["GEMINI_API_KEY"]
                    models_to_try = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.5-pro"]
                    
                    response_text = ""
                    success_model = None
                    
                    prompt = """
                    Eres un asistente OCR especializado en procesamiento de encuestas del "Proyecto Construyendo Futuro" (SRPA).
                    Analiza las dos páginas adjuntas que pertenecen al mismo participante y extrae la información en formato JSON estricto sin incluir texto introductorio ni explicaciones de ningún tipo. No rodees el JSON con bloques de código markdown, entrega solo las llaves crudas.

                    Reglas de extracción:
                    1. Identifica si el formulario es "PRETEST" o "POSTEST".
                    2. Extrae los datos de cabecera de la Página 1:
                       - Fecha (conviértelo a formato YYYY-MM-DD si es legible, de lo contrario deja vacío)
                       - Municipio (el usuario indicó que es de este municipio, pero reporta lo que esté escrito si lo hay)
                       - Institución Educativa: lee cuidadosamente la caligrafía manuscrita a puño y letra en la cabecera. Reporta exactamente lo que diga (ej: "Promesa de Dios").
                       - Rol del participante (Estudiante, Docente, Padre de Familia, Lider comunitario). Debe marcarse de acuerdo al cuadro con la 'X'.
                    3. Lee las respuestas a las preguntas de conocimientos (marcadas con 'X' u otra marca):
                       - En la Página 1 están las preguntas de conocimientos 1 a 3.
                       - En la Página 2 están las preguntas de conocimientos restantes (preguntas 4 a 8 si es Pretest o conocimientos de postest si es Postest).
                       - Para cada pregunta reporta la opción marcada (ej: "a", "b", "c", "d"). SI NO HAY MARCA o el participante la dejó en blanco, reporta exactamente un campo de texto vacío "". NUNCA infieras ni rellenes respuestas falsas.
                    4. En la Página 2 del POSTEST, lee la tabla de satisfacción (preguntas 1 a 9 con valores: Excelente, Bueno, Regular, Deficiente). Si no hay marcas, reporta "".

                    Formato JSON estricto esperado:
                    {
                      "tipo_formulario": "PRETEST" o "POSTEST",
                      "fecha": "YYYY-MM-DD o vacío",
                      "municipio": "Municipio",
                      "institucion_educativa": "Nombre de la I.E. escrito a mano",
                      "rol": "Estudiante/Docente/Padre de Familia/Lider comunitario",
                      "conocimientos": {
                        "p1": "opción o vacío",
                        "p2": "opción o vacío",
                        "p3": "opción o vacío",
                        "p4": "opción o vacío",
                        "p5": "opción o vacío",
                        "p6": "opción o vacío",
                        "p7": "opción o vacío",
                        "p8": "opción o vacío"
                      },
                      "satisfaccion": {
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
                    """
                    
                    if USE_NEW_SDK:
                        client = genai.Client(api_key=api_key)
                        for model_name in models_to_try:
                            try:
                                # Convert images to bytes for the new SDK
                                img1_byte_arr = io.BytesIO()
                                img1.save(img1_byte_arr, format='JPEG')
                                img1_bytes = img1_byte_arr.getvalue()
                                
                                img2_byte_arr = io.BytesIO()
                                img2.save(img2_byte_arr, format='JPEG')
                                img2_bytes = img2_byte_arr.getvalue()
                                
                                response = client.models.generate_content(
                                    model=model_name,
                                    contents=[
                                        types.Part.from_bytes(data=img1_bytes, mime_type="image/jpeg"),
                                        types.Part.from_bytes(data=img2_bytes, mime_type="image/jpeg"),
                                        prompt
                                    ]
                                )
                                if response.text:
                                    response_text = response.text
                                    success_model = model_name
                                    break
                            except Exception as model_err:
                                continue
                    else:
                        genai_legacy.configure(api_key=api_key)
                        for model_name in models_to_try:
                            try:
                                model = genai_legacy.GenerativeModel(model_name)
                                response = model.generate_content([img1, img2, prompt])
                                if response.text:
                                    response_text = response.text
                                    success_model = model_name
                                    break
                            except Exception as model_err:
                                continue
                                
                    if not response_text:
                        raise ValueError("Ninguno de los modelos de Gemini pudo procesar la solicitud.")
                        
                    # Clean response text from potential markdown block wrappers
                    clean_json = response_text.strip()
                    if clean_json.startswith("```json"):
                        clean_json = clean_json[7:]
                    if clean_json.endswith("```"):
                        clean_json = clean_json[:-3]
                    clean_json = clean_json.strip()
                    
                    parsed_data = json.loads(clean_json)
                    
                    # Create Unique ID
                    timestamp_id = datetime.now().strftime("%Y%m%d_%H%M%S")
                    id_encuesta = f"ENC_{timestamp_id}"
                    
                    # Compress and encode images for layout verification
                    img1_base64 = compress_and_base64(file_pag1)
                    img2_base64 = compress_and_base64(file_pag2)
                    
                    # Prepare GSheets row for Cola_Revision
                    # ID_Encuesta, Fecha_Carga, Tipo_Formulario, Municipio, Institucion_Educativa_IA, Rol, JSON_Respuestas, Estado, Img_Pag1_Base64, Img_Pag2_Base64
                    row_data = [
                        id_encuesta,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        parsed_data.get("tipo_formulario", "PRETEST"),
                        municipio_input,
                        parsed_data.get("institucion_educativa", ""),
                        parsed_data.get("rol", "Estudiante"),
                        json.dumps(parsed_data),
                        "Pendiente",
                        img1_base64,
                        img2_base64
                    ]
                    
                    # Write and Verify Save in Cola_Revision (Col A = col index 1)
                    is_ok, msg_verif = append_row_and_verify(ws_cola, row_data, 1, id_encuesta)
                    
                    if is_ok:
                        st.markdown(f"<div class='status-ok'>🎉 {msg_verif}<br>Vaya a la pestaña 'Banco de Verificación' para validar la escritura a mano e ingresar los datos a la base consolidada.</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div class='status-error'>❌ Fallo de Verificación: {msg_verif}</div>", unsafe_allow_html=True)
                        
                except Exception as ex:
                    st.error(f"Error procesando la encuesta con Gemini ({success_model or 'Error en API'}): {ex}")

# --- TAB 2: BANCO DE VERIFICACIÓN ---
with tab2:
    st.markdown("### Banco de Verificación de Escritura a Mano")
    st.markdown("Verifique y corrija el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")

    # Reload data live
    df_cola = load_worksheet_to_df(ws_cola, COLS_COLA)
    
    # Filter for Pendiente
    df_pendientes = df_cola[df_cola["Estado"] == "Pendiente"] if not df_cola.empty else pd.DataFrame()
    
    if df_pendientes.empty:
        st.info("No hay encuestas pendientes de verificación en este momento. Cargue encuestas en la pestaña de carga.")
    else:
        st.markdown(f"Hay **{len(df_pendientes)}** encuestas esperando revisión humana de caligrafía.")
        
        # Select survey to review
        survey_ids = df_pendientes["ID_Encuesta"].tolist()
        selected_id = st.selectbox("Seleccione Encuesta a Verificar:", survey_ids)
        
        row_review = df_pendientes[df_pendientes["ID_Encuesta"] == selected_id].iloc[0]
        
        # Display side-by-side comparison images
        col1_img, col2_img = st.columns(2)
        
        with col1_img:
            st.markdown("**Página 1 (Verificación de Cabecera):**")
            if row_review["Img_Pag1_Base64"]:
                try:
                    img_data1 = base64.b64decode(row_review["Img_Pag1_Base64"])
                    st.image(Image.open(io.BytesIO(img_data1)), use_container_width=True)
                except Exception as img_err:
                    st.warning("No se pudo desplegar la imagen de la Página 1.")
            else:
                st.info("Sin previsualización de Página 1.")
                
        with col2_img:
            st.markdown("**Página 2:**")
            if row_review["Img_Pag2_Base64"]:
                try:
                    img_data2 = base64.b64decode(row_review["Img_Pag2_Base64"])
                    st.image(Image.open(io.BytesIO(img_data2)), use_container_width=True)
                except Exception as img_err:
                    st.warning("No se pudo desplegar la imagen de la Página 2.")
            else:
                st.info("Sin previsualización de Página 2.")

        # Display metadata & verification form
        st.markdown("---")
        st.markdown("#### Datos Sugeridos por la IA")
        
        try:
            full_json_data = json.loads(row_review["JSON_Respuestas"])
        except Exception:
            full_json_data = {}

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            tipo_form = st.text_input("Tipo de Formulario", row_review["Tipo_Formulario"], disabled=True)
        with col_f2:
            municipio_val = st.text_input("Municipio", row_review["Municipio"], disabled=True)
        with col_f3:
            rol_val = st.text_input("Rol", row_review["Rol"], disabled=True)

        # Field for human-corrected school name
        ie_sugerida = row_review["Institucion_Educativa_IA"]
        ie_corregida = st.text_input("✏️ Nombre de la Institución Educativa Manuscrita (Verifique con la imagen superior):", ie_sugerida)
        
        # Option buttons
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True):
                if not ie_corregida.strip():
                    st.error("Por favor, ingrese un nombre válido para la Institución Educativa antes de aprobar.")
                else:
                    with st.spinner("Guardando en la base de datos de producción consolidada..."):
                        # Structure final responses row for Respuestas_SRPA
                        conocimientos = full_json_data.get("conocimientos", {})
                        satisfaccion = full_json_data.get("satisfaccion", {})
                        
                        # Prepare finalized row
                        final_row = [
                            row_review["ID_Encuesta"],
                            row_review["Tipo_Formulario"],
                            datetime.now().strftime("%Y-%m-%d"),
                            row_review["Municipio"],
                            ie_corregida.strip(),
                            row_review["Rol"],
                            conocimientos.get("p1", ""),
                            conocimientos.get("p2", ""),
                            conocimientos.get("p3", ""),
                            conocimientos.get("p4", ""),
                            conocimientos.get("p5", ""),
                            conocimientos.get("p6", ""),
                            conocimientos.get("p7", ""),
                            conocimientos.get("p8", ""),
                            satisfaccion.get("sat_p1", ""),
                            satisfaccion.get("sat_p2", ""),
                            satisfaccion.get("sat_p3", ""),
                            satisfaccion.get("sat_p4", ""),
                            satisfaccion.get("sat_p5", ""),
                            satisfaccion.get("sat_p6", ""),
                            satisfaccion.get("sat_p7", ""),
                            satisfaccion.get("sat_p8", ""),
                            satisfaccion.get("sat_p9", ""),
                            "Verificador_Bolivar", # Verificador
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ]
                        
                        # 1. Append and verify in Respuestas_SRPA (Col A index = 1)
                        is_written, msg_write = append_row_and_verify(ws_resp, final_row, 1, row_review["ID_Encuesta"])
                        
                        if is_written:
                            # 2. Update status in Cola_Revision to Aprobado and verify
                            is_updated, msg_up = update_row_and_verify(ws_cola, row_review["ID_Encuesta"], "Estado", "Aprobado", COLS_COLA)
                            if is_updated:
                                st.success(f"¡Encuesta aprobada con éxito! {msg_write}")
                                time.sleep(1.0)
                                st.rerun()
                            else:
                                st.error(f"Fallo de actualización de estado: {msg_up}")
                        else:
                            st.error(f"⚠️ Error Crítico de Integridad: No se pudo verificar la inserción en la base de datos de producción. {msg_write}")
                            
        with col_btn2:
            if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                with st.spinner("Descartando registro..."):
                    is_updated, msg_up = update_row_and_verify(ws_cola, row_review["ID_Encuesta"], "Estado", "Rechazado", COLS_COLA)
                    if is_updated:
                        st.warning("Encuesta rechazada y eliminada de la cola activa.")
                        time.sleep(1.0)
                        st.rerun()
                    else:
                        st.error(f"Error rechazando entrada: {msg_up}")

# --- TAB 3: DASHBOARD ESTADÍSTICO ---
with tab3:
    st.markdown("### Dashboard Estadístico de Conocimientos e Impacto")
    
    # Load responses
    df_resp = load_worksheet_to_df(ws_resp, COLS_RESP)
    
    if df_resp.empty:
        st.info("La base de datos consolidada se encuentra limpia y vacía. Aprobando encuestas en el 'Banco de Verificación' se poblarán las estadísticas automáticamente en tiempo real.")
    else:
        st.markdown(f"Total de Registros Consolidados Reales: **{len(df_resp)}**")
        
        # Filter widgets
        filter_mun = st.multiselect("Filtrar por Municipio", df_resp["Municipio"].unique().tolist(), default=df_resp["Municipio"].unique().tolist())
        filter_ie = st.multiselect("Filtrar por Institución Educativa", df_resp["Institucion_Educativa_Verificada"].unique().tolist(), default=df_resp["Institucion_Educativa_Verificada"].unique().tolist())
        
        df_filtered = df_resp[(df_resp["Municipio"].isin(filter_mun)) & (df_resp["Institucion_Educativa_Verificada"].isin(filter_ie))]
        
        if df_filtered.empty:
            st.warning("No hay datos para los filtros seleccionados.")
        else:
            # Layout metrics
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric("Total Encuestas Filtradas", len(df_filtered))
            with col_m2:
                pre_count = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
                st.metric("Cuestionarios Pretest", pre_count)
            with col_m3:
                post_count = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
                st.metric("Cuestionarios Postest", post_count)

            # Display Data Frame for confirmation
            st.markdown("#### Tabla de Datos Registrados Reales")
            st.dataframe(df_filtered[["ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol"]])
            
            # Simple Chart analysis comparing Pretest and Postest responses
            st.markdown("#### Análisis de Conocimientos (Comparativa Pretest vs Postest)")
            
            # Count answers to Question 2 (finalidad del SRPA)
            # Correct answer is option 'b' (Promover la responsabilidad, la protección...)
            p2_pre = df_filtered[(df_filtered["Tipo_Formulario"] == "PRETEST") & (df_filtered["Conocimientos_P2"].str.lower().str.startswith("b"))]
            p2_post = df_filtered[(df_filtered["Tipo_Formulario"] == "POSTEST") & (df_filtered["Conocimientos_P1"].str.lower().str.startswith("b"))] # Note: In Postest question 1 is about SRPA purpose
            
            p2_pre_pct = (len(p2_pre) / max(1, pre_count)) * 100
            p2_post_pct = (len(p2_post) / max(1, post_count)) * 100
            
            st.info(f"💡 **Porcentaje de respuestas correctas sobre la finalidad del SRPA:** Pretest: **{p2_pre_pct:.1f}%** | Postest: **{p2_post_pct:.1f}%**")
            
            # Satisfacción Donuts Chart
            st.markdown("#### Matriz de Satisfacción del Taller (Resultados de Postest)")
            sat_cols = ["Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9"]
            
            df_sat = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]
            if not df_sat.empty:
                sat_summary = {}
                for col in sat_cols:
                    val_counts = df_sat[col].value_counts().to_dict()
                    sat_summary[col] = val_counts
                
                st.write(pd.DataFrame(sat_summary).fillna(0))
            else:
                st.info("Cargue cuestionarios de tipo POSTEST para visualizar la matriz de satisfacción.")
