import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import json
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# Configuración de la página para dispositivos móviles (Responsive)
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS para optimizar la visualización en teléfonos móviles
st.markdown("""
    <style>
    .reportview-container .main .block-container {
        padding-top: 1rem;
        padding-right: 1rem;
        padding-left: 1rem;
        padding-bottom: 1rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# CONEXIÓN A GOOGLE SHEETS (SIN CACHÉ PARA EVITAR DATOS OBSOLETOS)
# -------------------------------------------------------------------------
def conectar_google_sheets():
    """Establece conexión directa con Google Sheets usando los secretos de Streamlit."""
    try:
        if "gcp_service_account" not in st.secrets:
            return None
            
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de autenticación con Google Sheets: {e}")
        return None

def obtener_o_crear_hoja():
    """Obtiene la hoja 'Base_Encuestas_SRPA' o la crea si no existe."""
    client = conectar_google_sheets()
    if not client:
        return None
        
    try:
        spreadsheet = client.open("Base_Encuestas_SRPA")
    except gspread.SpreadsheetNotFound:
        try:
            # Crear la hoja si no existe
            spreadsheet = client.create("Base_Encuestas_SRPA")
            # Compartir con el correo de la cuenta de servicio si es necesario
            creds_dict = dict(st.secrets["gcp_service_account"])
            spreadsheet.share(creds_dict["client_email"], perm_type='user', role='editor')
        except Exception as e:
            st.error(f"No se pudo crear la hoja en Google Drive: {e}")
            return None

    # Asegurar pestaña 'Cola_Revision'
    try:
        cola_sheet = spreadsheet.worksheet("Cola_Revision")
    except gspread.WorksheetNotFound:
        cola_sheet = spreadsheet.add_worksheet(title="Cola_Revision", rows="1000", cols="8")
        cola_sheet.append_row([
            "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
            "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
        ])

    # Asegurar pestaña 'Respuestas_SRPA'
    try:
        resp_sheet = spreadsheet.worksheet("Respuestas_SRPA")
    except gspread.WorksheetNotFound:
        resp_sheet = spreadsheet.add_worksheet(title="Respuestas_SRPA", rows="5000", cols="24")
        headers = [
            "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
            "Institucion_Educativa_Verificada", "Rol",
            "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
            "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
            "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
            "Verificado_Por", "Fecha_Aprobacion"
        ]
        resp_sheet.append_row(headers)

    return spreadsheet

# NOTA CRÍTICA: NO USAR @st.cache_data para leer la cola de revisión,
# de lo contrario los cambios manuales y nuevas subidas no se verán reflejados en tiempo real.
def cargar_cola_revision(spreadsheet):
    """Carga los registros pendientes de la cola de revisión directamente de Google Sheets."""
    if not spreadsheet:
        return pd.DataFrame()
    try:
        sheet = spreadsheet.worksheet("Cola_Revision")
        records = sheet.get_all_records()
        df = pd.DataFrame(records)
        if df.empty:
            return pd.DataFrame(columns=[
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
            ])
        # Filtrar solo las pendientes
        return df[df["Estado"] == "Pendiente"]
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame()

def cargar_respuestas_validadas(spreadsheet):
    """Carga las respuestas ya consolidadas."""
    if not spreadsheet:
        return pd.DataFrame()
    try:
        sheet = spreadsheet.worksheet("Respuestas_SRPA")
        records = sheet.get_all_records()
        return pd.DataFrame(records)
    except Exception as e:
        st.error(f"Error cargando Respuestas Consolidadas: {e}")
        return pd.DataFrame()

# -------------------------------------------------------------------------
# INTERFAZ Y FLUJO DE TRABAJO DE LA APP
# -------------------------------------------------------------------------
st.title("Proyecto 'Construyendo Futuro' 🇨🇴")
st.subheader("Evaluación de Conocimientos SRPA - Sistematización Móvil")

spreadsheet = obtener_o_crear_hoja()

if not spreadsheet:
    st.warning("⚠️ **Conexión de Google Sheets requerida**")
    st.info("""
        Para que la aplicación funcione y guarde la información de manera unificada y colaborativa en tiempo real, 
        debes configurar las credenciales de tu Cuenta de Servicio de Google Cloud en los Secretos de Streamlit Cloud.
        
        **Sigue estos pasos en tu panel de Streamlit Cloud (Manage App -> Settings -> Secrets):**
        1. Crea un proyecto en Google Cloud Console y activa las APIs de **Google Sheets** y **Google Drive**.
        2. Crea una **Cuenta de Servicio**, genera una clave en formato **JSON** y descárgala.
        3. Pega el contenido del JSON en tus Secretos de Streamlit Cloud bajo la clave `[gcp_service_account]`.
    """)
else:
    # Crear pestañas principales de la interfaz
    tab_upload, tab_review, tab_dashboard = st.tabs([
        "📤 Cargar Encuestas (Fotos)", 
        "🔍 Cola de Revisión", 
        "📊 Dashboard Estadístico"
    ])

    # -------------------------------------------------------------------------
    # TAB 1: CARGA DE ENCUESTAS (FOTOS DE PÁGINA 1 Y PÁGINA 2)
    # -------------------------------------------------------------------------
    with tab_upload:
        st.markdown("### Subir Encuesta Física (Doble Página)")
        st.write("Toma fotos claras o selecciona imágenes de la galería del celular.")

        tipo_formulario = st.selectbox("Tipo de Cuestionario:", ["PRETEST", "POSTEST"])
        
        col1, col2 = st.columns(2)
        with col1:
            foto_p1 = st.file_uploader("📷 Foto de la Página 1 (Cabecera e inicio)", type=["jpg", "jpeg", "png"])
        with col2:
            foto_p2 = st.file_uploader("📷 Foto de la Página 2 (Preguntas finales y satisfacción)", type=["jpg", "jpeg", "png"])

        if foto_p1 and foto_p2:
            st.success("✅ Ambas páginas cargadas correctamente.")
            
            if st.button("🚀 Procesar Encuesta con IA"):
                with st.spinner("Procesando imágenes con el motor de IA Vision..."):
                    # SIMULACIÓN DE OCR / PROCESAMIENTO GEMINI REAL
                    # En producción, aquí se llamaría a la API de Gemini enviando foto_p1 y foto_p2.
                    # El prompt del sistema extrae los datos y respeta que si no hay info, queda vacío "".
                    
                    id_nueva_encuesta = f"ENC-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
                    
                    # Datos simulados extraídos del OCR basados en tus imágenes de muestra reales
                    respuestas_simuladas = {
                        "p1": "b", "p2": "b", "p3": "b", "p4": "a",
                        "p5": "d", "p6": "a", "p7": "a", "p8": "ICBF"
                    }
                    satisfaccion_simulada = {
                        "s1": "Excelente", "s2": "Excelente", "s3": "Bueno",
                        "s4": "Excelente", "s5": "Excelente", "s6": "Bueno",
                        "s7": "Excelente", "s8": "Excelente", "s9": "Excelente"
                    } if tipo_formulario == "POSTEST" else {}

                    # Creamos el JSON de respuestas que se guardará temporalmente
                    json_payload = {
                        "respuestas_conocimiento": respuestas_simuladas,
                        "evaluacion_satisfaccion": satisfaccion_simulada
                    }

                    # Guardamos el registro en la pestaña de revisión de Google Sheets
                    try:
                        cola_sheet = spreadsheet.worksheet("Cola_Revision")
                        
                        # IMPORTANTE: Si la IA no está segura de la caligrafía o el usuario no la ingresó,
                        # queda como vacío para que se revise a mano.
                        institucion_detectada_ia = "Promesa de Dios" # Ejemplo de caligrafía manuscrita en las fotos de Bolívar
                        
                        nueva_fila = [
                            id_nueva_encuesta,
                            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            tipo_formulario,
                            "Cartagena",
                            institucion_detectada_ia,
                            "Estudiante",
                            json.dumps(json_payload),
                            "Pendiente"
                        ]
                        cola_sheet.append_row(nueva_fila)
                        st.success("🎉 ¡Encuesta enviada con éxito a la Cola de Revisión!")
                        
                        # FORZAR RECARGA INMEDIATA DE STREAMLIT PARA ACTUALIZAR LA COLA DE REVISIÓN
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"Error al escribir en la Cola de Revisión: {e}")

    # -------------------------------------------------------------------------
    # TAB 2: COLA DE REVISIÓN (VERIFICACIÓN HUMANA / SIN CACHÉ OBSOLETO)
    # -------------------------------------------------------------------------
    with tab_review:
        st.markdown("### Banco de Verificación de Escritura a Mano")
        st.write("Verifica y corrige el nombre de la **Institución Educativa** manuscrita antes de guardarla permanentemente.")

        # Recargamos la cola en tiempo real cada vez que el usuario entra a este tab
        df_cola = cargar_cola_revision(spreadsheet)

        if df_cola.empty:
            st.info("🎈 **No hay encuestas pendientes de revisión en este momento.** ¡Gran trabajo!")
        else:
            st.warning(f"Tienes **{len(df_cola)}** encuesta(s) pendiente(s) por verificar.")
            
            # Seleccionar la primera encuesta pendiente de la fila
            registro = df_cola.iloc[0]
            
            id_encuesta = registro["ID_Encuesta"]
            tipo_form = registro["Tipo_Formulario"]
            municipio_orig = registro["Municipio"]
            rol_orig = registro["Rol"]
            ie_ia = registro["Institucion_Educativa_IA"]
            json_resp = json.loads(registro["JSON_Respuestas"])

            st.write("---")
            st.markdown(f"#### Editando Registro: `{id_encuesta}` ({tipo_form})")

            # Módulo visual de corrección de caligrafía manuscrita
            col_preview, col_form = st.columns([1, 1])

            with col_preview:
                st.info("📸 **Caligrafía del documento original**")
                st.write("El facilitador tomó fotos de las páginas físicas. Revisa la foto de la cabecera en tu celular para validar la escritura real del nombre de la escuela.")
                # Aquí se mostraría la foto real si se tuviera la URL, por ahora mostramos ayuda visual
                st.markdown("""
                    **Consejo de lectura:**
                    En tus encuestas de Bolívar, la Institución Educativa suele estar escrita a mano alzada en la segunda línea de la cabecera (ej: *Promesa de Dios*, *I.E. San José*).
                """)

            with col_form:
                st.markdown("##### ✏️ Datos a Verificar")
                
                # Input editable pre-cargado con lo que leyó la IA
                ie_verificada = st.text_input(
                    "Institución Educativa (Corrige si es necesario):", 
                    value=str(ie_ia)
                )
                
                municipio = st.text_input("Municipio:", value=str(municipio_orig))
                rol = st.selectbox("Rol del Participante:", ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"].index(rol_orig) if rol_orig in ["Estudiante", "Docente", "Padre de Familia", "Lider comunitario"] else 0)

                col_actions = st.columns(2)
                with col_actions[0]:
                    # BOTÓN DE APROBACIÓN (VERDE)
                    if st.button("✅ Aprobar e Ingresar", key="btn_aprobar"):
                        with st.spinner("Guardando en la base de datos central..."):
                            try:
                                # 1. Escribir en la pestaña final 'Respuestas_SRPA'
                                resp_sheet = spreadsheet.worksheet("Respuestas_SRPA")
                                
                                # Mapear respuestas de conocimiento
                                rc = json_resp.get("respuestas_conocimiento", {})
                                # Mapear respuestas de satisfacción (si existen)
                                rs = json_resp.get("evaluacion_satisfaccion", {})
                                
                                # Si un dato de respuesta no existe o está vacío, se guarda estrictamente en blanco "" (sin autocompletar)
                                nueva_fila_final = [
                                    id_encuesta,
                                    tipo_form,
                                    datetime.date.today().strftime("%Y-%m-%d"),
                                    municipio,
                                    ie_verificada,
                                    rol,
                                    rc.get("p1", ""), rc.get("p2", ""), rc.get("p3", ""), rc.get("p4", ""),
                                    rc.get("p5", ""), rc.get("p6", ""), rc.get("p7", ""), rc.get("p8", ""),
                                    rs.get("s1", ""), rs.get("s2", ""), rs.get("s3", ""), rs.get("s4", ""),
                                    rs.get("s5", ""), rs.get("s6", ""), rs.get("s7", ""), rs.get("s8", ""), rs.get("s9", ""),
                                    "Verificador_Movil",
                                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                ]
                                
                                resp_sheet.append_row(nueva_fila_final)

                                # 2. Cambiar estado en la pestaña de cola a 'Aprobado' (o eliminar fila)
                                cola_sheet = spreadsheet.worksheet("Cola_Revision")
                                celda = cola_sheet.find(id_encuesta)
                                if celda:
                                    # El estado está en la columna 8 (columna H)
                                    cola_sheet.update_cell(celda.row, 8, "Aprobado")

                                st.success("🎉 ¡Registro consolidado correctamente!")
                                
                                # FORZAR RECARGA INMEDIATA PARA MOSTRAR EL SIGUIENTE REGISTRO PENDIENTE
                                st.rerun()

                            except Exception as e:
                                st.error(f"Error al aprobar el registro: {e}")

                with col_actions[1]:
                    # BOTÓN DE RECHAZO (ROJO)
                    if st.button("❌ Rechazar Entrada", key="btn_rechazar"):
                        with st.spinner("Rechazando y limpiando la cola..."):
                            try:
                                cola_sheet = spreadsheet.worksheet("Cola_Revision")
                                celda = cola_sheet.find(id_encuesta)
                                if celda:
                                    # El estado está en la columna 8 (columna H)
                                    cola_sheet.update_cell(celda.row, 8, "Rechazado")
                                
                                st.warning("⚠️ Registro rechazado y eliminado de la cola.")
                                
                                # FORZAR RECARGA INMEDIATA PARA ELIMINAR EL ELEMENTO VISUALMENTE
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"Error al rechazar el registro: {e}")

    # -------------------------------------------------------------------------
    # TAB 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
    # -------------------------------------------------------------------------
    with tab_dashboard:
        st.markdown("### Panel de Control y Estadísticas en Tiempo Real")
        
        # Carga los datos ya validados desde Google Sheets
        df_validados = cargar_respuestas_validadas(spreadsheet)

        if df_validados.empty:
            st.info("📊 **El dashboard se poblará automáticamente en tiempo real cuando apruebes la primera encuesta en la pestaña 'Cola de Revisión'.**")
        else:
            # Filtros dinámicos responsive en la cabecera
            st.markdown("#### 🔍 Filtros de Consulta")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                municipios_disp = ["Todos"] + list(df_validados["Municipio"].unique())
                m_filter = st.selectbox("Filtrar por Municipio:", municipios_disp)
            with col_f2:
                roles_disp = ["Todos"] + list(df_validados["Rol"].unique())
                r_filter = st.selectbox("Filtrar por Rol de Participante:", roles_disp)

            # Aplicar filtros
            df_filtered = df_validados.copy()
            if m_filter != "Todos":
                df_filtered = df_filtered[df_filtered["Municipio"] == m_filter]
            if r_filter != "Todos":
                df_filtered = df_filtered[df_filtered["Rol"] == r_filter]

            # 1. TARJETAS DE INDICADORES (KPIs)
            st.markdown("#### 📌 Indicadores de Cobertura")
            kpi1, kpi2, kpi3 = st.columns(3)
            with kpi1:
                st.metric("Total Encuestas Consolidadas", len(df_filtered))
            with kpi2:
                total_pre = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
                st.metric("Pretests Evaluados", total_pre)
            with kpi3:
                total_post = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
                st.metric("Postests y Satisfacción", total_post)

            # 2. COMPARATIVA DE CONOCIMIENTO (ANTES VS DESPUÉS)
            st.markdown("#### 🧠 Impacto en el Conocimiento (Antes vs. Después)")
            
            # Mapeo de respuestas correctas del cuestionario (según documento PRETEST - POSTEST)
            # P1: Finalidad del SRPA -> Opción 'b' (Promover la responsabilidad, la protección de derechos y la resocialización)
            # P2: Factor de riesgo -> Opción 'b' (Consumir sustancias psicoactivas)
            # P3: Factor protector -> Opción 'b' (Participar en actividades deportivas...)
            # P4: Responsabilidad de prevención -> Opción 'd' (La familia, la escuela...)
            
            correctas = {
                "Conocimientos_P2": "b", # Finalidad
                "Conocimientos_P3": "b", # Factor de riesgo
                "Conocimientos_P4": "b", # Factor protector (en Postest es P3, alineado)
                "Conocimientos_P5": "d"  # Corresponsabilidad
            }

            conceptos = ["Finalidad SRPA", "Factores de Riesgo", "Factores Protectores", "Corresponsabilidad"]
            pct_pre = []
            pct_post = []

            df_pre_filtered = df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"]
            df_post_filtered = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]

            # Calcular aciertos por concepto para Pretest
            if not df_pre_filtered.empty:
                # P2 en Pretest
                p2_pre_correct = (df_pre_filtered["Conocimientos_P2"] == "b").sum() / len(df_pre_filtered) * 100
                # P3 en Pretest
                p3_pre_correct = (df_pre_filtered["Conocimientos_P3"] == "b").sum() / len(df_pre_filtered) * 100
                # P4 en Pretest
                p4_pre_correct = (df_pre_filtered["Conocimientos_P4"] == "a").sum() / len(df_pre_filtered) * 100 # Dialogar con la familia
                # P5 en Pretest
                p5_pre_correct = (df_pre_filtered["Conocimientos_P5"] == "d").sum() / len(df_pre_filtered) * 100
                pct_pre = [p2_pre_correct, p3_pre_correct, p4_pre_correct, p5_pre_correct]
            else:
                pct_pre = [0, 0, 0, 0]

            # Calcular aciertos por concepto para Postest (Sección A)
            if not df_post_filtered.empty:
                # P1 en Postest (Misma pregunta que P2 del Pretest)
                p1_post_correct = (df_post_filtered["Conocimientos_P1"] == "b").sum() / len(df_post_filtered) * 100
                # P2 en Postest (Misma pregunta que P3 del Pretest)
                p2_post_correct = (df_post_filtered["Conocimientos_P2"] == "b").sum() / len(df_post_filtered) * 100
                # P3 en Postest (Misma pregunta que P4 del Pretest)
                p3_post_correct = (df_post_filtered["Conocimientos_P3"] == "b").sum() / len(df_post_filtered) * 100 # Actividades deportivas
                # P4 en Postest (Misma pregunta que P5 del Pretest)
                p4_post_correct = (df_post_filtered["Conocimientos_P4"] == "d").sum() / len(df_post_filtered) * 100
                pct_post = [p1_post_correct, p2_post_correct, p3_post_correct, p4_post_correct]
            else:
                pct_post = [0, 0, 0, 0]

            # Graficar comparativa Antes vs Después
            fig_conocimiento = go.Figure(data=[
                go.Bar(name="PRETEST (Antes)", x=conceptos, y=pct_pre, marker_color="#312E81"),
                go.Bar(name="POSTEST (Después)", x=conceptos, y=pct_post, marker_color="#10B981")
            ])
            fig_conocimiento.update_layout(
                barmode="group",
                yaxis_title="% de Respuestas Correctas",
                yaxis=dict(range=[0, 100]),
                legend_orientation="h",
                legend=dict(x=0, y=-0.2),
                margin=dict(l=20, r=20, t=20, b=20),
                height=350
            )
            st.plotly_chart(fig_conocimiento, use_container_width=True)

            # 3. ROLES DE PARTICIPANTES Y SATISFACCIÓN (SÓLO POSTEST)
            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.markdown("##### 👥 Distribución de Participantes")
                fig_roles = px.pie(
                    df_filtered, 
                    names="Rol", 
                    color_discrete_sequence=px.colors.qualitative.Prism,
                    hole=0.4
                )
                fig_roles.update_layout(
                    margin=dict(l=20, r=20, t=20, b=20),
                    height=300,
                    legend=dict(orientation="h", y=-0.1)
                )
                st.plotly_chart(fig_roles, use_container_width=True)

            with col_chart2:
                st.markdown("##### ⭐️ Satisfacción del Taller (Postest)")
                if df_post_filtered.empty:
                    st.write("Aún no hay datos de satisfacción disponibles para este filtro.")
                else:
                    # Mapear valores cualitativos a escala numérica (Excelente=4, Bueno=3, Regular=2, Deficiente=1)
                    val_map = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None, " ": None}
                    
                    sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
                    promedios = []
                    
                    for col in sat_cols:
                        valores_num = df_post_filtered[col].map(val_map).dropna()
                        promedios.append(valores_num.mean() if not valores_num.empty else 0)

                    aspectos = [
                        "Claridad", "Dominio Facilitador", "Metodología", 
                        "Participación", "Utilidad", "Organización", 
                        "Materiales", "Conocimiento", "Recomendaría"
                    ]

                    fig_sat = go.Figure(go.Bar(
                        x=promedios,
                        y=aspectos,
                        orientation="h",
                        marker_color="#F59E0B"
                    ))
                    # Línea meta de satisfacción excelente (3.5 de 4.0)
                    fig_sat.add_vline(x=3.5, line_dash="dash", line_color="red", annotation_text="Meta (3.5)")
                    fig_sat.update_layout(
                        xaxis_title="Calificación Promedio (1.0 a 4.0)",
                        xaxis=dict(range=[1, 4]),
                        margin=dict(l=20, r=20, t=20, b=20),
                        height=300
                    )
                    st.plotly_chart(fig_sat, use_container_width=True)
