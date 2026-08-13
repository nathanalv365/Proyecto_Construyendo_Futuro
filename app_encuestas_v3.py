import streamlit as st
import pandas as pd
import json
from PIL import Image
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Intentar importar librerías de Google Sheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    G_SHEETS_AVAILABLE = True
except ImportError:
    G_SHEETS_AVAILABLE = False

# Intentar importar librerías de Gemini
GEMINI_LIB = None
try:
    from google import genai
    from google.genai import types
    GEMINI_LIB = "genai"
except ImportError:
    try:
        import google.generativeai as google_genai
        GEMINI_LIB = "generativeai"
    except ImportError:
        GEMINI_LIB = None

# ==========================================================
# CONFIGURACIÓN DE PÁGINA Y DISEÑO CSS (CORREGIDO)
# ==========================================================
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado para celulares e interfaz moderna
st.markdown("""
<style>
    .main {
        background-color: #f8fafc;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
    }
    .success-box {
        padding: 15px;
        background-color: #f0fdf4;
        border-left: 5px solid #16a34a;
        border-radius: 4px;
        color: #14532d;
        margin-bottom: 15px;
    }
    .warning-box {
        padding: 15px;
        background-color: #fffbeb;
        border-left: 5px solid #d97706;
        border-radius: 4px;
        color: #78350f;
        margin-bottom: 15px;
    }
    .error-box {
        padding: 15px;
        background-color: #fef2f2;
        border-left: 5px solid #dc2626;
        border-radius: 4px;
        color: #7f1d1d;
        margin-bottom: 15px;
    }
    .card {
        padding: 20px;
        background-color: #ffffff;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 20px;
        border: 1px solid #e2e8f0;
    }
    h1, h2, h3 {
        color: #1e293b;
    }
</style>
""", unsafe_allow_html=True) # <-- AQUÍ ESTÁ LA CORRECCIÓN CRÍTICA

# ==========================================================
# CONEXIÓN CON GOOGLE SHEETS
# ==========================================================
def obtener_cliente_sheets():
    if not G_SHEETS_AVAILABLE:
        return None, "Librerías de Google Sheets no instaladas en el servidor."
    
    if "gcp_service_account" not in st.secrets:
        return None, "No se encontraron las credenciales 'gcp_service_account' en los Secretos de Streamlit."
    
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Limpiar saltos de línea en la clave privada
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, f"Error de autenticación con Google Cloud: {str(e)}"

def inicializar_hoja_sheets():
    client, error = obtener_cliente_sheets()
    if error:
        return None, error
    
    try:
        # Intentar abrir la hoja
        spreadsheet = client.open("Base_Encuestas_SRPA")
    except gspread.SpreadsheetNotFound:
        try:
            # Crear si no existe
            spreadsheet = client.create("Base_Encuestas_SRPA")
            # Compartir con el usuario que creó la cuenta de servicio si es necesario
            # spreadsheet.share('tu_correo@gmail.com', perm_type='user', role='writer')
        except Exception as e:
            return None, f"No se encontró la hoja y no se pudo crear: {str(e)}"
    
    # Asegurar existencia de pestañas
    try:
        cola_sheet = spreadsheet.worksheet("Cola_Revision")
    except gspread.WorksheetNotFound:
        cola_sheet = spreadsheet.add_worksheet(title="Cola_Revision", rows="1000", cols="8")
        cola_sheet.append_row([
            "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
            "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
        ])
        
    try:
        respuestas_sheet = spreadsheet.worksheet("Respuestas_SRPA")
    except gspread.WorksheetNotFound:
        respuestas_sheet = spreadsheet.add_worksheet(title="Respuestas_SRPA", rows="5000", cols="24")
        headers = [
            "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
            "Institucion_Educativa_Verificada", "Rol",
            "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
            "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
            "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
            "Verificado_Por", "Fecha_Aprobacion"
        ]
        respuestas_sheet.append_row(headers)
        
    return spreadsheet, None

# ==========================================================
# GESTIÓN DE DATOS EN LA NUBE
# ==========================================================
def cargar_cola_revision(spreadsheet):
    try:
        sheet = spreadsheet.worksheet("Cola_Revision")
        records = sheet.get_all_records()
        if not records:
            return pd.DataFrame(columns=[
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
            ])
        df = pd.DataFrame(records)
        # Asegurar columnas críticas por seguridad
        for col in ["ID_Encuesta", "Estado", "Institucion_Educativa_IA", "JSON_Respuestas"]:
            if col not in df.columns:
                df[col] = ""
        return df[df["Estado"] == "Pendiente"]
    except Exception as e:
        st.error(f"Error al cargar cola de revisión: {str(e)}")
        return pd.DataFrame()

def cargar_respuestas_validadas(spreadsheet):
    try:
        sheet = spreadsheet.worksheet("Respuestas_SRPA")
        records = sheet.get_all_records()
        if not records:
            return pd.DataFrame()
        return pd.DataFrame(records)
    except Exception as e:
        st.error(f"Error al cargar respuestas validadas: {str(e)}")
        return pd.DataFrame()

# ==========================================================
# INTEGRACIÓN CON GEMINI VISION API (OCR)
# ==========================================================
def procesar_encuesta_con_ia(pag1_bytes, pag2_bytes):
    # Obtener clave de API
    api_key = st.secrets.get("GEMINI_API_KEY", None)
    if not api_key:
        return None, "No se configuró la variable GEMINI_API_KEY en los secretos."
    
    if not GEMINI_LIB:
        return None, "No se encuentran instaladas las librerías de Google Gemini en el servidor."

    # Prompt estructurado de análisis
    prompt = """
    Eres un sistema de Inteligencia Artificial especializado en procesamiento de encuestas del "Proyecto Construyendo Futuro" (SRPA).
    Analiza las dos imágenes proporcionadas (Página 1 y Página 2 de un mismo cuestionario físico) y extrae la información con alta precisión.

    Instrucciones críticas:
    1. Identifica el tipo de formulario: "PRETEST" o "POSTEST" buscando el título en la cabecera.
    2. En la cabecera, extrae:
       - Fecha (en formato YYYY-MM-DD o lo más cercano posible)
       - Municipio
       - Institución Educativa (presta atención a la escritura a mano alzada)
       - Rol del participante: marca Estudiante, Docente, Padre de Familia o Líder comunitario.
    3. Para las respuestas de conocimiento (Sección A):
       - Identifica qué opción (a, b, c, d) tiene una "X" marcada. 
       - Si una pregunta está vacía, no tiene ninguna X o está en blanco, devuélvela estrictamente como una cadena vacía "". No asumas ni inventes respuestas.
    4. Para la Sección B de Satisfacción (solo si es POSTEST):
       - Extrae la opción marcada (Excelente, Bueno, Regular, Deficiente) para cada una de las 9 filas.
       - Si una fila está vacía, devuélvela como una cadena vacía "". No asumas respuestas.

    Formato de salida requerido: JSON puro, sin markdown, que coincida con el siguiente esquema:
    {
      "tipo_formulario": "PRETEST" o "POSTEST",
      "fecha": "2026-08-11",
      "municipio": "Cartagena",
      "institucion_educativa": "Nombre de la escuela escrito a mano",
      "rol": "Estudiante",
      "respuestas_conocimiento": {
        "p1": "a", "p2": "b", "p3": "b", "p4": "a", 
        "p5": "d", "p6": "a", "p7": "a", "p8": "ICBF"
      },
      "evaluacion_satisfaccion": {
        "s1": "Excelente", "s2": "Bueno", "s3": "", "s4": "Excelente",
        "s5": "Excelente", "s6": "Bueno", "s7": "Excelente", "s8": "Excelente", "s9": "Excelente"
      }
    }
    """

    try:
        if GEMINI_LIB == "genai":
            client = genai.Client(api_key=api_key)
            # Preparar partes de bytes
            part1 = types.Part.from_bytes(data=pag1_bytes, mime_type="image/jpeg")
            part2 = types.Part.from_bytes(data=pag2_bytes, mime_type="image/jpeg")
            
            response = client.models.generate_content(
                model='gemini-2.0-flash',
                contents=[part1, part2, prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            text_result = response.text
        else:
            # google-generativeai clásico
            google_genai.configure(api_key=api_key)
            model = google_genai.GenerativeModel('gemini-1.5-flash')
            
            img1 = Image.open(io.BytesIO(pag1_bytes))
            img2 = Image.open(io.BytesIO(pag2_bytes))
            
            response = model.generate_content([img1, img2, prompt])
            text_result = response.text
        
        # Limpieza por seguridad si la IA devuelve bloques markdown de código ```json
        if "```json" in text_result:
            text_result = text_result.split("```json")[1].split("```")[0].strip()
        elif "```" in text_result:
            text_result = text_result.split("```")[1].split("```")[0].strip()
            
        return json.loads(text_result), None
    except Exception as e:
        return None, f"Error de comunicación con la API de Gemini: {str(e)}"

# ==========================================================
# APLICACIÓN PRINCIPAL - CONTROLADOR
# ==========================================================
def main():
    st.title("Proyecto Construyendo Futuro 📱")
    st.subheader("Evaluación de Conocimientos y Satisfacción - SRPA")
    
    # Intentar inicializar Google Sheets
    spreadsheet = None
    sheet_error = None
    
    if G_SHEETS_AVAILABLE:
        spreadsheet, sheet_error = inicializar_hoja_sheets()
    else:
        sheet_error = "Las librerías de Google Sheets (gspread, oauth2client) no están disponibles."

    # Mostrar error si no se pudo conectar con Sheets de manera clara
    if sheet_error:
        st.markdown(f"""
        <div class="error-box">
            <h4>⚠️ Error de Conexión a Base de Datos en la Nube</h4>
            <p>{sheet_error}</p>
            <p><strong>¿Cómo solucionarlo?</strong> Asegúrate de que las credenciales de tu cuenta de servicio de Google Cloud estén correctamente configuradas en el apartado <strong>Secrets</strong> de Streamlit Cloud como <code>[gcp_service_account]</code>.</p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # Creación del menú de pestañas interactivo
    tab_upload, tab_review, tab_dashboard = st.tabs([
        "📸 Capturar / Cargar Encuesta", 
        "✏️ Banco de Verificación de Caligrafía", 
        "📊 Dashboard Estadístico"
    ])

    # ------------------------------------------------------
    # PESTAÑA 1: CARGAR ENCUESTA (MÓVIL)
    # ------------------------------------------------------
    with tab_upload:
        st.markdown("### Digitalización de Encuesta por Foto")
        st.write("Sube la **Página 1** y la **Página 2** del mismo cuestionario para que el sistema las analice de forma conjunta.")
        
        col1, col2 = st.columns(2)
        with col1:
            file_pag1 = st.file_uploader("Subir foto de la Página 1 (Cabecera y Preguntas 1-3)", type=["jpg", "png", "jpeg"])
        with col2:
            file_pag2 = st.file_uploader("Subir foto de la Página 2 (Preguntas de Conocimiento y Satisfacción)", type=["jpg", "png", "jpeg"])
            
        if file_pag1 and file_pag2:
            st.markdown("""
            <div class="success-box">
                ✔️ Las imágenes se cargaron correctamente. ¡Estás listo para enviarlas a procesar con Inteligencia Artificial!
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("🔍 Iniciar Análisis con IA"):
                with st.spinner("La Inteligencia Artificial está procesando ambas páginas simultáneamente..."):
                    # Leer bytes de imágenes
                    pag1_bytes = file_pag1.read()
                    pag2_bytes = file_pag2.read()
                    
                    resultado, error = procesar_encuesta_con_ia(pag1_bytes, pag2_bytes)
                    
                    if error:
                        st.error(error)
                    else:
                        # Guardar temporalmente en la cola de revisión
                        try:
                            cola_sheet = spreadsheet.worksheet("Cola_Revision")
                            id_encuesta = f"EN_SRPA_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                            
                            # Preparar datos temporales
                            tipo_f = resultado.get("tipo_formulario", "PRETEST")
                            municipio = resultado.get("municipio", "")
                            ie_ia = resultado.get("institucion_educativa", "")
                            rol = resultado.get("rol", "Estudiante")
                            
                            # Respuestas empaquetadas en JSON
                            json_respuestas = json.dumps({
                                "respuestas_conocimiento": resultado.get("respuestas_conocimiento", {}),
                                "evaluacion_satisfaccion": resultado.get("evaluacion_satisfaccion", {})
                            })
                            
                            cola_sheet.append_row([
                                id_encuesta,
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                tipo_f,
                                municipio,
                                ie_ia,
                                rol,
                                json_respuestas,
                                "Pendiente"
                            ])
                            
                            st.balloons()
                            st.success(f"¡Análisis exitoso! Encuesta '{id_encuesta}' enviada a revisión de caligrafía.")
                            # Forzar recarga rápida
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al escribir en la cola de revisión en Google Sheets: {str(e)}")

    # ------------------------------------------------------
    # PESTAÑA 2: BANCO DE VERIFICACIÓN (HUMAN-IN-THE-LOOP)
    # ------------------------------------------------------
    with tab_review:
        st.markdown("### Banco de Verificación de Escritura a Mano")
        st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
        
        df_cola = cargar_cola_revision(spreadsheet)
        
        if df_cola.empty:
            st.info("🎉 ¡Excelente trabajo! No hay encuestas pendientes de verificación de caligrafía en este momento.")
        else:
            st.warning(f"Hay {len(df_cola)} encuestas en cola pendientes de revisión.")
            
            # Tomar la primera encuesta en cola
            fila_actual = df_cola.iloc[0]
            id_encuesta = fila_actual["ID_Encuesta"]
            ie_ia = fila_actual["Institucion_Educativa_IA"]
            tipo_f = fila_actual["Tipo_Formulario"]
            municipio = fila_actual["Municipio"]
            rol = fila_actual["Rol"]
            
            # Decodificar respuestas
            try:
                resp_data = json.loads(fila_actual["JSON_Respuestas"])
                conocimientos = resp_data.get("respuestas_conocimiento", {})
                satisfaccion = resp_data.get("evaluacion_satisfaccion", {})
            except Exception:
                conocimientos = {}
                satisfaccion = {}
                
            st.markdown(f"**Revisando Registro:** `{id_encuesta}` | **Tipo:** `{tipo_f}`")
            
            # Formulario interactivo de corrección
            col_rev1, col_rev2 = st.columns(2)
            with col_rev1:
                st.markdown("<div class='card'>", unsafe_allow_html=True)
                st.markdown("#### ✏️ Validación de Cabecera")
                ie_verificada = st.text_input("Institución Educativa (Corrige la caligrafía si es necesario):", value=ie_ia)
                municipio_verificado = st.text_input("Municipio:", value=municipio)
                rol_verificado = st.selectbox("Rol del Participante:", ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(rol) if rol in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0)
                st.markdown("</div>", unsafe_allow_html=True)
                
            with col_rev2:
                st.markdown("<div class='card'>", unsafe_allow_html=True)
                st.markdown("#### 📋 Respuestas Extraídas")
                st.write("**Respuestas de Conocimiento:**")
                st.write(conocimientos)
                if tipo_f == "POSTEST":
                    st.write("**Evaluación de Satisfacción:**")
                    st.write(satisfaccion)
                st.markdown("</div>", unsafe_allow_html=True)
                
            # Botones de Acción
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("✅ Aprobar e Ingresar a Base de Datos", type="primary"):
                    with st.spinner("Guardando en la base de datos oficial..."):
                        try:
                            respuestas_sheet = spreadsheet.worksheet("Respuestas_SRPA")
                            
                            # Formatear fila para Base de Datos
                            fila_final = [
                                id_encuesta,
                                tipo_f,
                                datetime.now().strftime("%Y-%m-%d"),
                                municipio_verificado,
                                ie_verificada,
                                rol_verificado,
                                conocimientos.get("p1", ""), conocimientos.get("p2", ""),
                                conocimientos.get("p3", ""), conocimientos.get("p4", ""),
                                conocimientos.get("p5", ""), conocimientos.get("p6", ""),
                                conocimientos.get("p7", ""), conocimientos.get("p8", ""),
                                satisfaccion.get("s1", ""), satisfaccion.get("s2", ""),
                                satisfaccion.get("s3", ""), satisfaccion.get("s4", ""),
                                satisfaccion.get("s5", ""), satisfaccion.get("s6", ""),
                                satisfaccion.get("s7", ""), satisfaccion.get("s8", ""),
                                satisfaccion.get("s9", ""),
                                "Revisor_Movil",
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            ]
                            
                            # Escribir en base final
                            respuestas_sheet.append_row(fila_final)
                            
                            # Actualizar estado en la cola de revisión de gsheets
                            cola_sheet = spreadsheet.worksheet("Cola_Revision")
                            celda = cola_sheet.find(id_encuesta)
                            # Actualizar columna "Estado" (columna 8)
                            cola_sheet.update_cell(celda.row, 8, "Aprobado")
                            
                            st.success("¡Registro consolidado y aprobado con éxito!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al consolidar la información: {str(e)}")
                            
            with col_btn2:
                if st.button("❌ Rechazar Entrada (Eliminar de la Cola)"):
                    with st.spinner("Eliminando entrada defectuosa..."):
                        try:
                            cola_sheet = spreadsheet.worksheet("Cola_Revision")
                            celda = cola_sheet.find(id_encuesta)
                            cola_sheet.update_cell(celda.row, 8, "Rechazado")
                            st.success("Registro rechazado y eliminado de la cola.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al rechazar el registro: {str(e)}")

    # ------------------------------------------------------
    # PESTAÑA 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
    # ------------------------------------------------------
    with tab_dashboard:
        st.markdown("### Dashboard Estadístico del Proyecto")
        
        df_final = cargar_respuestas_validadas(spreadsheet)
        
        if df_final.empty:
            st.markdown("""
            <div class="warning-box">
                <h4>📊 No hay información registrada</h4>
                <p>La base de datos se encuentra vacía actualmente. Comienza a subir fotos en el apartado <strong>Capturar / Cargar Encuesta</strong> para ver gráficos automatizados en tiempo real aquí.</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            # Mostrar métricas clave de Cobertura
            total_registros = len(df_final)
            pretests = len(df_final[df_final["Tipo_Formulario"] == "PRETEST"])
            postests = len(df_final[df_final["Tipo_Formulario"] == "POSTEST"])
            
            st.markdown("#### 📈 Resumen General de Cobertura")
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric("Total de Participantes", f"{total_registros}")
            with col_m2:
                st.metric("Pretests Realizados", f"{pretests}")
            with col_m3:
                st.metric("Postests Realizados", f"{postests}")
                
            # Filtro dinámico por Municipio
            municipios_disponibles = list(df_final["Municipio"].unique())
            municipio_sel = st.selectbox("Filtrar por Municipio:", ["Todos"] + municipios_disponibles)
            
            df_filtrado = df_final if municipio_sel == "Todos" else df_final[df_final["Municipio"] == municipio_sel]
            
            # Gráficos Dinámicos
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.markdown("<div class='card'>", unsafe_allow_html=True)
                st.write("**Participantes por Rol**")
                rol_counts = df_filtrado["Rol"].value_counts().reset_index()
                rol_counts.columns = ["Rol", "Cantidad"]
                fig_rol = px.pie(rol_counts, names="Rol", values="Cantidad", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_rol, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)
                
            with col_g2:
                st.markdown("<div class='card'>", unsafe_allow_html=True)
                st.write("**Participación por Institución Educativa**")
                ie_counts = df_filtrado["Institucion_Educativa_Verificada"].value_counts().reset_index()
                ie_counts.columns = ["IE", "Cantidad"]
                fig_ie = px.bar(ie_counts, x="Cantidad", y="IE", orientation="h", color_discrete_sequence=["#2563eb"])
                st.plotly_chart(fig_ie, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)

            # Cálculo de progreso e Impacto Educativo (Preguntas compartidas)
            # P1 en Pretest / Postest (en el pretest es p2, en el postest es p1 - Finalidad del SRPA)
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.write("#### 🎯 Comparativa de Respuestas Correctas (Pretest vs. Postest)")
            st.write("Análisis de asimilación del conocimiento en las jornadas del Proyecto.")
            
            # Filtrar pretest y postest del dataset filtrado
            df_pre = df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"]
            df_pos = df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"]
            
            if not df_pre.empty and not df_pos.empty:
                # Calcular aciertos en Finalidad del SRPA (opción 'b' es correcta)
                aciertos_p1_pre = (df_pre["Conocimientos_P2"] == "b").mean() * 100
                aciertos_p1_pos = (df_pos["Conocimientos_P1"] == "b").mean() * 100
                
                # Calcular aciertos en Factores de Riesgo (opción 'b' es correcta)
                aciertos_p2_pre = (df_pre["Conocimientos_P3"] == "b").mean() * 100
                aciertos_p2_pos = (df_pos["Conocimientos_P2"] == "b").mean() * 100
                
                # Gráfico comparativo
                fig_comp = go.Figure(data=[
                    go.Bar(name='PRETEST (Antes)', x=['Finalidad del SRPA', 'Factores de Riesgo'], y=[aciertos_p1_pre, aciertos_p2_pre], marker_color='#93c5fd'),
                    go.Bar(name='POSTEST (Después)', x=['Finalidad del SRPA', 'Factores de Riesgo'], y=[aciertos_p1_pos, aciertos_p2_pos], marker_color='#2563eb')
                ])
                fig_comp.update_layout(barmode='group', yaxis_title='Porcentaje de Aciertos (%)', yaxis_range=[0, 100])
                st.plotly_chart(fig_comp, use_container_width=True)
            else:
                st.info("Se requiere contar con al menos un registro de Pretest y un registro de Postest aprobados para visualizar la gráfica comparativa de impacto.")
            st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()

                st.plotly_chart(fig_sat, use_container_width=True)
            else:
                st.info("ℹ️ Para ver las métricas de satisfacción y el desglose de calidad de las jornadas, aprueba al menos una encuesta de tipo **POSTEST**.")

