import streamlit as st
import pandas as pd
import json
import os
from PIL import Image
import io
import datetime
import plotly.express as px
import plotly.graph_objects as go
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuración de página
st.set_page_config(
    page_title="Sistematización SRPA - Gobernación de Bolívar",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header {
        color: #1E3A8A;
        font-weight: bold;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-header {
        color: #4B5563;
        text-align: center;
        margin-bottom: 25px;
    }
    .card {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #E5E7EB;
    }
    .stButton>button {
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Conectar con Google Sheets
def conectar_google_sheets():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Limpiar saltos de línea en la clave privada
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n").replace("\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        return None

# Obtener o crear pestaña
def obtener_worksheet(client, sheet_name, ws_name, expected_headers):
    try:
        try:
            sh = client.open(sheet_name)
        except gspread.exceptions.SpreadsheetNotFound:
            sh = client.create(sheet_name)
            
        try:
            ws = sh.worksheet(ws_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=ws_name, rows="1000", cols=str(len(expected_headers) + 5))
            ws.append_row(expected_headers)
            
        return ws
    except Exception as e:
        st.error(f"Error accediendo a la pestaña {ws_name}: {e}")
        return None

# Cargar Cola de Revisión de forma segura sin usar get_all_records()
def cargar_cola_revision():
    headers = [
        'ID_Encuesta', 'Fecha_Carga', 'Tipo_Formulario', 'Municipio', 
        'Institucion_Educativa_IA', 'Rol', 'JSON_Respuestas', 'Estado', 
        'URL_Foto_P1', 'URL_Foto_P2'
    ]
    client = conectar_google_sheets()
    if not client:
        return pd.DataFrame(columns=headers)
        
    ws = obtener_worksheet(client, "Base_Encuestas_SRPA", "Cola_Revision", headers)
    if not ws:
        return pd.DataFrame(columns=headers)
        
    try:
        all_values = ws.get_all_values()
        if not all_values:
            ws.append_row(headers)
            return pd.DataFrame(columns=headers)
            
        current_headers = [h.strip() for h in all_values[0]]
        # Filtrar columnas vacías para evitar errores de pandas
        valid_indices = [i for i, h in enumerate(current_headers) if h]
        clean_headers = [current_headers[i] for i in valid_indices]
        
        rows = []
        for r in all_values[1:]:
            padded = r + [''] * (len(current_headers) - len(r))
            clean_row = [padded[i] for i in valid_indices]
            rows.append(clean_row)
            
        df = pd.DataFrame(rows, columns=clean_headers)
        
        # Asegurar que todas las columnas requeridas existan
        for col in headers:
            if col not in df.columns:
                df[col] = ""
                
        # Filtrar solo pendientes
        df = df[df['Estado'] == 'Pendiente']
        return df
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame(columns=headers)

# Cargar Respuestas de forma segura sin usar get_all_records()
def cargar_respuestas_validadas():
    headers = [
        'ID_Encuesta', 'Tipo_Formulario', 'Fecha', 'Municipio', 'Institucion_Educativa_Verificada', 'Rol',
        'Conocimientos_P1', 'Conocimientos_P2', 'Conocimientos_P3', 'Conocimientos_P4', 'Conocimientos_P5',
        'Conocimientos_P6', 'Conocimientos_P7', 'Conocimientos_P8',
        'Sat_P1', 'Sat_P2', 'Sat_P3', 'Sat_P4', 'Sat_P5', 'Sat_P6', 'Sat_P7', 'Sat_P8', 'Sat_P9',
        'Verificado_Por', 'Fecha_Aprobacion'
    ]
    client = conectar_google_sheets()
    if not client:
        return pd.DataFrame(columns=headers)
        
    ws = obtener_worksheet(client, "Base_Encuestas_SRPA", "Respuestas_SRPA", headers)
    if not ws:
        return pd.DataFrame(columns=headers)
        
    try:
        all_values = ws.get_all_values()
        if not all_values:
            ws.append_row(headers)
            return pd.DataFrame(columns=headers)
            
        current_headers = [h.strip() for h in all_values[0]]
        valid_indices = [i for i, h in enumerate(current_headers) if h]
        clean_headers = [current_headers[i] for i in valid_indices]
        
        rows = []
        for r in all_values[1:]:
            padded = r + [''] * (len(current_headers) - len(r))
            clean_row = [padded[i] for i in valid_indices]
            rows.append(clean_row)
            
        df = pd.DataFrame(rows, columns=clean_headers)
        
        # Asegurar que todas las columnas requeridas existan
        for col in headers:
            if col not in df.columns:
                df[col] = ""
                
        return df
    except Exception as e:
        st.error(f"Error cargando Respuestas SRPA: {e}")
        return pd.DataFrame(columns=headers)

# Llamar a la API de Gemini (Soporta SDK nuevo y antiguo de forma transparente)
def llamar_gemini(api_key, img_p1, img_p2, prompt):
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)
        
        contents = []
        if img_p1:
            img_byte_arr = io.BytesIO()
            img_p1.save(img_byte_arr, format='JPEG')
            contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type='image/jpeg'))
        if img_p2:
            img_byte_arr = io.BytesIO()
            img_p2.save(img_byte_arr, format='JPEG')
            contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type='image/jpeg'))
            
        contents.append(prompt)
        
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.1
            )
        )
        return response.text
    except Exception as e:
        # Fallback a google-generativeai clásico
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-3.6-flash')
            
            contents = []
            if img_p1:
                contents.append(img_p1)
            if img_p2:
                contents.append(img_p2)
            contents.append(prompt)
            
            response = model.generate_content(
                contents,
                generation_config={"response_mime_type": "application/json", "temperature": 0.1}
            )
            return response.text
        except Exception as ex:
            raise Exception(f"Error procesando con Gemini: {ex}")

# Menú lateral
st.sidebar.image("https://www.bolivar.gov.co/images/Gobernacion/Logo-Gobernacion-Color.png", use_container_width=True, error_bad_lines=False)
st.sidebar.title("Navegación")
app_mode = st.sidebar.radio("Ir a:", ["Cargar Encuestas", "Cola de Revisión", "Dashboard Estadístico"])

# Verificar configuración inicial
client = conectar_google_sheets()
api_key_configured = "GEMINI_API_KEY" in st.secrets

if not client:
    st.warning("⚠️ **Conexión con Google Sheets no configurada en Secrets.**")
    st.markdown("""
    ### ¿Cómo conectar tu aplicación a Google Sheets?
    Para que las encuestas se guarden directamente en la nube de forma colaborativa, debes pegar tus credenciales en el cuadro **Secrets** de Streamlit Cloud.
    
    Sigue estos pasos rápidos:
    1. Genera un archivo de credenciales JSON de cuenta de servicio en la consola de Google Cloud.
    2. Comparte tu archivo de Google Sheets llamado `Base_Encuestas_SRPA` con el correo de tu cuenta de servicio en modo **Editor**.
    3. Pega tus secretos en Streamlit Cloud usando este formato:
    ```toml
    GEMINI_API_KEY = "tu_clave_api_aquí"
    
    [gcp_service_account]
    type = "service_account"
    project_id = "tu-proyecto-id"
    private_key_id = "tu-private-key-id"
    private_key = "-----BEGIN PRIVATE KEY-----\nTU_LLAVE_PRIVADA_AQUÍ\n-----END PRIVATE KEY-----\n"
    client_email = "tu-cuenta-de-servicio@proyecto.iam.gserviceaccount.com"
    client_id = "tu-client-id"
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.google.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/tu-cuenta-de-servicio"
    ```
    """, unsafe_allow_html=True)
    st.stop()

if not api_key_configured:
    st.warning("🔑 **Por favor, ingresa tu GEMINI_API_KEY en la configuración de Streamlit Cloud Secrets.**")
    st.stop()

# ----------------- PANTALLA 1: CARGAR ENCUESTAS -----------------
if app_mode == "Cargar Encuestas":
    st.markdown("<h1 class='main-header'>Cargar Encuestas Físicas</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Toma o selecciona fotos de las dos páginas correspondientes a una sola encuesta.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Página 1: Encabezado y Datos Generales")
        foto_p1 = st.file_uploader("Subir Página 1", type=["jpg", "jpeg", "png"], key="p1")
        if foto_p1:
            st.image(foto_p1, caption="Vista previa Página 1", use_container_width=True)
            
    with col2:
        st.subheader("Página 2: Cuestionario y Satisfacción")
        foto_p2 = st.file_uploader("Subir Página 2", type=["jpg", "jpeg", "png"], key="p2")
        if foto_p2:
            st.image(foto_p2, caption="Vista previa Página 2", use_container_width=True)
            
    if foto_p1 and foto_p2:
        if st.button("🚀 Procesar Encuesta con IA", use_container_width=True, type="primary"):
            with st.spinner("Procesando imágenes con Gemini 3.6 Flash..."):
                try:
                    img1 = Image.open(foto_p1)
                    img2 = Image.open(foto_p2)
                    
                    prompt = """
                    Analiza las dos imágenes correspondientes a la Página 1 y Página 2 de una encuesta del "Proyecto Construyendo Futuro".
                    Extrae la información con precisión matemática de OCR.
                    
                    Reglas estrictas de extracción:
                    1. Identifica el tipo de formulario: "PRETEST" o "POSTEST" de acuerdo con el encabezado.
                    2. Lee el encabezado: fecha, municipio, institución educativa (manuscrita) y el rol del participante marcado con una equis (X).
                    3. Lee las respuestas a las preguntas de conocimientos y satisfacción (marcas con X).
                    4. IMPORTANTE: Si una pregunta está en blanco, no tiene marcas o es ilegible, devuelve un string vacío "" para esa pregunta. No infieras datos falsos.
                    
                    Formato de retorno estrictamente JSON:
                    {
                      "tipo_formulario": "PRETEST",
                      "fecha": "2026-08-18",
                      "municipio": "NombreMunicipio",
                      "institucion_educativa": "NombreIE",
                      "rol": "Estudiante",
                      "conocimientos": {
                        "p1": "a", "p2": "", "p3": "b"
                      },
                      "satisfaccion": {
                        "s1": "Excelente", "s2": ""
                      }
                    }
                    """
                    
                    gemini_key = st.secrets["GEMINI_API_KEY"]
                    respuesta_json = llamar_gemini(gemini_key, img1, img2, prompt)
                    
                    # Limpiar markdown de respuesta si existiera
                    if respuesta_json.startswith("```json"):
                        respuesta_json = respuesta_json[7:-3].strip()
                    elif respuesta_json.startswith("```"):
                        respuesta_json = respuesta_json[3:-3].strip()
                        
                    data_ext = json.loads(respuesta_json)
                    
                    # Guardar temporalmente en la cola de revisión
                    ws_cola = obtener_worksheet(client, "Base_Encuestas_SRPA", "Cola_Revision", [])
                    id_encuesta = f"ENC_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    
                    row = [
                        id_encuesta,
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        data_ext.get("tipo_formulario", "PRETEST"),
                        data_ext.get("municipio", ""),
                        data_ext.get("institucion_educativa", ""),
                        data_ext.get("rol", "Estudiante"),
                        json.dumps(data_ext),
                        "Pendiente",
                        "", # URL foto P1 (temporal/local)
                        ""  # URL foto P2 (temporal/local)
                    ]
                    ws_cola.append_row(row)
                    st.success("✅ ¡Encuesta procesada con éxito! Se ha añadido a la **Cola de Revisión** para que valides el nombre de la Institución Educativa.")
                    st.balloons()
                    
                    # Forzar recarga de estado
                    st.rerun()
                except Exception as e:
                    st.error(f"Error procesando la encuesta: {e}")

# ----------------- PANTALLA 2: COLA DE REVISIÓN -----------------
elif app_mode == "Cola de Revisión":
    st.markdown("<h1 class='main-header'>Banco de Verificación de Escritura</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Valida o corrige el nombre de la Institución Educativa escrito a puño y letra antes de guardarlo en la base definitiva.</p>", unsafe_allow_html=True)
    
    df_cola = cargar_cola_revision()
    
    if df_cola.empty:
        st.info("🎉 ¡Excelente! No hay encuestas pendientes en la cola de revisión.")
    else:
        st.write(f"Tienes **{len(df_cola)}** encuestas listas para revisar.")
        
        # Seleccionar registro para revisar
        selected_id = st.selectbox("Seleccione ID de Encuesta para verificar:", df_cola['ID_Encuesta'].tolist())
        row_sel = df_cola[df_cola['ID_Encuesta'] == selected_id].iloc[0]
        
        try:
            datos_ia = json.loads(row_sel['JSON_Respuestas'])
        except:
            datos_ia = {}
            
        st.subheader("Datos Extraídos por la IA")
        col_form1, col_form2 = st.columns(2)
        with col_form1:
            municipio_sel = st.text_input("Municipio:", value=row_sel['Municipio'])
            rol_sel = st.text_input("Rol de Participante:", value=row_sel['Rol'])
            tipo_f = row_sel['Tipo_Formulario']
            st.info(f"Tipo de Formulario: **{tipo_f}**")
            
        with col_form2:
            ie_manuscrita = st.text_input("🏫 Institución Educativa (Corrige de ser necesario):", value=row_sel['Institucion_Educativa_IA'])
            fecha_sel = st.text_input("Fecha de Captura:", value=datos_ia.get("fecha", datetime.datetime.now().strftime('%Y-%m-%d')))

        st.markdown("---")
        
        # Botones de acción
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True, type="primary"):
                with st.spinner("Consolidando registro en Google Sheets..."):
                    try:
                        ws_resp = obtener_worksheet(client, "Base_Encuestas_SRPA", "Respuestas_SRPA", [])
                        
                        conocimientos = datos_ia.get("conocimientos", {})
                        satisfaccion = datos_ia.get("satisfaccion", {})
                        
                        # Armar la fila con todas las columnas correspondientes
                        new_row = [
                            row_sel['ID_Encuesta'],
                            tipo_f,
                            fecha_sel,
                            municipio_sel,
                            ie_manuscrita, # Institución validada
                            rol_sel,
                            conocimientos.get("p1", ""),
                            conocimientos.get("p2", ""),
                            conocimientos.get("p3", ""),
                            conocimientos.get("p4", ""),
                            conocimientos.get("p5", ""),
                            conocimientos.get("p6", ""),
                            conocimientos.get("p7", ""),
                            conocimientos.get("p8", ""),
                            satisfaccion.get("s1", ""),
                            satisfaccion.get("s2", ""),
                            satisfaccion.get("s3", ""),
                            satisfaccion.get("s4", ""),
                            satisfaccion.get("s5", ""),
                            satisfaccion.get("s6", ""),
                            satisfaccion.get("s7", ""),
                            satisfaccion.get("s8", ""),
                            satisfaccion.get("s9", ""),
                            "Verificador Campo",
                            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        ]
                        
                        ws_resp.append_row(new_row)
                        
                        # Actualizar estado en la cola de revisión
                        ws_cola = obtener_worksheet(client, "Base_Encuestas_SRPA", "Cola_Revision", [])
                        # Buscar fila en gspread por ID_Encuesta
                        cell = ws_cola.find(selected_id)
                        if cell:
                            ws_cola.update_cell(cell.row, 8, "Aprobado") # Columna 8: Estado
                            
                        st.success(f"¡Registro {selected_id} aprobado y consolidado con éxito!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error consolidando registro: {e}")
                        
        with btn_col2:
            if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                with st.spinner("Descartando registro..."):
                    try:
                        ws_cola = obtener_worksheet(client, "Base_Encuestas_SRPA", "Cola_Revision", [])
                        cell = ws_cola.find(selected_id)
                        if cell:
                            ws_cola.update_cell(cell.row, 8, "Rechazado")
                        st.warning("Se ha rechazado la encuesta y se eliminó de la lista de pendientes.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error rechazando registro: {e}")

# ----------------- PANTALLA 3: DASHBOARD ESTADÍSTICO -----------------
elif app_mode == "Dashboard Estadístico":
    st.markdown("<h1 class='main-header'>Dashboard Estadístico Real</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Estadísticas, cobertura e impacto del Proyecto 'Construyendo Futuro' consolidadas directamente desde Google Sheets.</p>", unsafe_allow_html=True)
    
    df_resp = cargar_respuestas_validadas()
    
    if df_resp.empty:
        st.info("📈 **Aún no hay encuestas validadas para mostrar en el panel.** Una vez que proceses y apruebes registros en la pestaña 'Cola de Revisión', este tablero se actualizará dinámicamente con los indicadores en tiempo real.")
    else:
        st.write(f"Visualizando análisis para **{len(df_resp)}** registros validados.")
        
        # Filtros dinámicos
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros del Dashboard")
        
        municipios_list = ["Todos"] + df_resp['Municipio'].unique().tolist()
        municipio_filtro = st.sidebar.selectbox("Filtrar por Municipio:", municipios_list)
        
        roles_list = ["Todos"] + df_resp['Rol'].unique().tolist()
        rol_filtro = st.sidebar.selectbox("Filtrar por Rol:", roles_list)
        
        # Aplicar filtros
        df_plot = df_resp.copy()
        if municipio_filtro != "Todos":
            df_plot = df_plot[df_plot['Municipio'] == municipio_filtro]
        if rol_filtro != "Todos":
            df_plot = df_plot[df_plot['Rol'] == rol_filtro]
            
        if df_plot.empty:
            st.warning("No hay registros que coincidan con los filtros seleccionados.")
        else:
            # Indicadores Clave KPI
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            with kpi_col1:
                st.metric("Total Registros Consolidados", len(df_plot))
            with kpi_col2:
                pre_cnt = len(df_plot[df_plot['Tipo_Formulario'] == 'PRETEST'])
                st.metric("Encuestas Pretest", pre_cnt)
            with kpi_col3:
                post_cnt = len(df_plot[df_plot['Tipo_Formulario'] == 'POSTEST'])
                st.metric("Encuestas Postest", post_cnt)
                
            st.markdown("---")
            
            # Gráficos
            g1, g2 = st.columns(2)
            with g1:
                st.subheader("Distribución de Participantes por Rol")
                rol_counts = df_plot['Rol'].value_counts().reset_index()
                rol_counts.columns = ['Rol', 'Cantidad']
                fig_rol = px.pie(rol_counts, values='Cantidad', names='Rol', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_rol, use_container_width=True)
                
            with g2:
                st.subheader("Respuestas Correctas: Pretest vs. Postest")
                
                # Definir respuestas correctas de conocimientos basadas en PRETEST - POSTEST.docx
                # Pregunta 2 Pretest / Pregunta 1 Postest: Finalidad del SRPA -> 'b' (Promover la responsabilidad...)
                # Pregunta 3 Pretest / Pregunta 2 Postest: Factor de riesgo -> 'b' (Consumir sustancias psicoactivas)
                # Pregunta 4 Pretest / Pregunta 3 Postest: Factor protector -> 'a'/'b' (En postest es 'b': Participar en actividades...)
                # Pregunta 5 Pretest / Pregunta 4 Postest: Responsabilidad -> 'd' (La familia, escuela, comunidad...)
                
                # Calcular porcentajes
                df_pre = df_plot[df_plot['Tipo_Formulario'] == 'PRETEST']
                df_post = df_plot[df_plot['Tipo_Formulario'] == 'POSTEST']
                
                pre_pct = 0.0
                post_pct = 0.0
                
                if not df_pre.empty:
                    # Contar respuestas correctas en columna p2 (finalidad del SRPA)
                    pre_correct = df_pre['Conocimientos_P2'].str.lower().str.strip() == 'b'
                    pre_pct = pre_correct.mean() * 100
                    
                if not df_post.empty:
                    post_correct = df_post['Conocimientos_P1'].str.lower().str.strip() == 'b'
                    post_pct = post_correct.mean() * 100
                    
                fig_comp = go.Figure(data=[
                    go.Bar(name='Pretest', x=['Finalidad del SRPA'], y=[pre_pct], marker_color='#3B82F6'),
                    go.Bar(name='Postest', x=['Finalidad del SRPA'], y=[post_pct], marker_color='#10B981')
                ])
                fig_comp.update_layout(yaxis_title='Porcentaje de Aciertos (%)', barmode='group', yaxis=dict(range=[0, 100]))
                st.plotly_chart(fig_comp, use_container_width=True)

            # Sección de Satisfacción (solo para Postests validados)
            df_sat = df_plot[(df_plot['Tipo_Formulario'] == 'POSTEST')]
            if not df_sat.empty:
                st.markdown("---")
                st.subheader("Evaluación de Satisfacción de la Jornada (Postest)")
                
                sat_cols = [f'Sat_P{i}' for i in range(1, 10)]
                aspectos = [
                    "Claridad de Información",
                    "Dominio de Facilitadores",
                    "Metodología",
                    "Participación",
                    "Utilidad de Temas",
                    "Organización",
                    "Materiales",
                    "Fortaleció Conocimientos",
                    "Recomendaría"
                ]
                
                # Calcular promedio ponderado Excelente=4, Bueno=3, Regular=2, Deficiente=1
                promedios = []
                for col in sat_cols:
                    if col in df_sat.columns:
                        vals = df_sat[col].str.lower().str.strip()
                        puntos = vals.map({'excelente': 4, 'bueno': 3, 'regular': 2, 'deficiente': 1}).fillna(0)
                        # Filtrar ceros para el promedio
                        valid_puntos = puntos[puntos > 0]
                        if not valid_puntos.empty:
                            promedios.append(valid_puntos.mean())
                        else:
                            promedios.append(0.0)
                    else:
                        promedios.append(0.0)
                        
                fig_sat = go.Figure(go.Bar(
                    x=promedios,
                    y=aspectos,
                    orientation='h',
                    marker_color='#8B5CF6'
                ))
                fig_sat.update_layout(xaxis_title='Puntuación Promedio (Escala 1 a 4)', xaxis=dict(range=[1, 4]))
                st.plotly_chart(fig_sat, use_container_width=True)
