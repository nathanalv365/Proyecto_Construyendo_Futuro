# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import json
import os
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import io
from datetime import datetime

# Importación segura de Gemini API
gemini_api_working = False
try:
    import google.generativeai as genai_classic
    gemini_api_working = True
    USE_CLASSIC = True
except ImportError:
    try:
        from google import genai as genai_new
        from google.genai import types as genai_types
        gemini_api_working = True
        USE_CLASSIC = False
    except ImportError:
        pass

# Configuración de la página para móvil
st.set_page_config(
    page_title="Proyecto Construyendo Futuro - SRPA",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS responsive personalizados para celulares
st.markdown("""
<style>
    .reportview-container .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 1.5rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
    }
    .card {
        padding: 1.5rem;
        border-radius: 10px;
        background-color: #ffffff;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        margin-bottom: 1rem;
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

COLUMNS_COLA = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"]
COLUMNS_RESPUESTAS = [
    "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
    "Institucion_Educativa_Verificada", "Rol",
    "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
    "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
    "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
    "Verificado_Por", "Fecha_Aprobacion"
]

# --- CONEXIÓN DE GOOGLE SHEETS ---
def get_gspread_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_info = dict(st.secrets["gcp_service_account"])
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
        from oauth2client.service_account import ServiceAccountCredentials
        import gspread
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de autenticación con Google: {e}")
        return None

def init_sheets():
    client = get_gspread_client()
    if not client:
        return None, None
    try:
        import gspread
        try:
            spreadsheet = client.open("Base_Encuestas_SRPA")
        except gspread.exceptions.SpreadsheetNotFound:
            spreadsheet = client.create("Base_Encuestas_SRPA")
            
        try:
            sheet_cola = spreadsheet.worksheet("Cola_Revision")
        except gspread.exceptions.WorksheetNotFound:
            sheet_cola = spreadsheet.add_worksheet("Cola_Revision", rows=100, cols=8)
            sheet_cola.append_row(COLUMNS_COLA)
            
        try:
            sheet_respuestas = spreadsheet.worksheet("Respuestas_SRPA")
        except gspread.exceptions.WorksheetNotFound:
            sheet_respuestas = spreadsheet.add_worksheet("Respuestas_SRPA", rows=1000, cols=25)
            sheet_respuestas.append_row(COLUMNS_RESPUESTAS)
            
        return sheet_cola, sheet_respuestas
    except Exception as e:
        st.error(f"Error inicializando hojas de cálculo de Google: {e}")
        return None, None

def load_cola_data(sheet_cola):
    try:
        records = sheet_cola.get_all_records()
        if not records:
            return pd.DataFrame(columns=COLUMNS_COLA)
        df = pd.DataFrame(records)
        for col in COLUMNS_COLA:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e:
        st.warning(f"La Cola de revisión está vacía o cargando estructura por defecto: {e}")
        return pd.DataFrame(columns=COLUMNS_COLA)

def load_respuestas_data(sheet_resp):
    try:
        records = sheet_resp.get_all_records()
        if not records:
            return pd.DataFrame(columns=COLUMNS_RESPUESTAS)
        df = pd.DataFrame(records)
        for col in COLUMNS_RESPUESTAS:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e:
        st.warning(f"La base de respuestas está vacía o cargando estructura por defecto: {e}")
        return pd.DataFrame(columns=COLUMNS_RESPUESTAS)

# --- LLAMADO A GEMINI VISION OCR ---
def run_gemini_ocr(img1_bytes, img2_bytes, mime_type="image/jpeg"):
    api_key = st.secrets.get("GEMINI_API_KEY", None)
    if not api_key:
        st.error("Por favor, ingresa tu GEMINI_API_KEY en la configuración de Streamlit Cloud Secrets.")
        return None
        
    prompt = """
    Eres un sistema OCR de alta precisión para el "Proyecto Construyendo Futuro" (SRPA) de la Gobernación de Bolívar.
    Se te proporcionan dos imágenes que corresponden a la Página 1 y la Página 2 del mismo formulario físico de encuesta (Pretest o Postest).
    
    Analiza ambas imágenes juntas para consolidar un único registro en formato JSON estricto.
    
    Instrucciones de extracción:
    1. Identifica el "tipo_formulario": Debe ser "PRETEST" o "POSTEST".
    2. Lee los datos manuscritos de la cabecera de la Página 1:
       - "fecha": Extrae la fecha (ej. "30/07/2026"). Si está vacía o no se entiende, pon "".
       - "municipio": Extrae el municipio. Si está vacío, pon "".
       - "institucion_educativa": Lee la escuela escrita a mano (ej. "Promesa de Dios"). Si está vacía, pon "".
       - "rol": Determina el rol del participante (Estudiante, Docente, Padre de Familia, Lider comunitario) según la casilla marcada.
    3. Lee las respuestas a las preguntas de conocimientos de opción múltiple marcadas con una (X).
       - En PRETEST hay 8 preguntas (Sección de Conocimientos 1 a 8).
       - En POSTEST hay 5 preguntas (Sección A. Postest de Conocimientos 1 a 5).
       - Extrae la letra seleccionada ("a", "b", "c", "d") o el texto si aplica (ej. pregunta 8 de Pretest).
       - SI UNA PREGUNTA ESTÁ EN BLANCO, DEBES DEVOLVER "". NO INVENTES RESPUESTAS.
    4. Lee las respuestas de satisfacción de la matriz (Página 2 - Solo si es POSTEST):
       - Hay 9 aspectos evaluados (1 a 9).
       - Las opciones son: Excelente, Bueno, Regular, Deficiente.
       - SI UN ASPECTO ESTÁ EN BLANCO, DEBES DEVOLVER "". NO INVENTES RESPUESTAS.
    
    Devuelve estrictamente un JSON sin bloques de código markdown, sin explicaciones ni rodeos:
    {
      "tipo_formulario": "PRETEST",
      "encabezado": {
        "fecha": "30/07/2026",
        "municipio": "Cartagena",
        "institucion_educativa": "Promesa de Dios",
        "rol": "Estudiante"
      },
      "respuestas_conocimiento": {
        "p1": "a", "p2": "b", "p3": "b", "p4": "a",
        "p5": "d", "p6": "a", "p7": "a", "p8": "ICBF"
      },
      "evaluacion_satisfaccion": {
        "s1": "", "s2": "Excelente", "s3": "Bueno", "s4": "Excelente",
        "s5": "Excelente", "s6": "", "s7": "Excelente", "s8": "Excelente", "s9": "Excelente"
      }
    }
    """
    
    try:
        if USE_CLASSIC:
            genai_classic.configure(api_key=api_key)
            model = genai_classic.GenerativeModel('gemini-1.5-flash')
            img1 = {"mime_type": mime_type, "data": img1_bytes}
            img2 = {"mime_type": mime_type, "data": img2_bytes}
            response = model.generate_content([prompt, img1, img2])
            text_resp = response.text.strip()
        else:
            client = genai_new.Client(api_key=api_key)
            part1 = genai_types.Part.from_bytes(data=img1_bytes, mime_type=mime_type)
            part2 = genai_types.Part.from_bytes(data=img2_bytes, mime_type=mime_type)
            response = client.models.generate_content(
                model='gemini-1.5-flash',
                contents=[prompt, part1, part2]
            )
            text_resp = response.text.strip()
            
        if text_resp.startswith("```json"):
            text_resp = text_resp.split("```json")[1].split("```")[0].strip()
        elif text_resp.startswith("```"):
            text_resp = text_resp.split("```")[1].split("```")[0].strip()
            
        return json.loads(text_resp)
    except Exception as e:
        st.error(f"Error procesando la imagen con Gemini: {e}")
        return None

# --- INICIO DE INTERFAZ DE USUARIO ---
st.title('Proyecto "Construyendo Futuro"')
st.subheader('Evaluación de Conocimientos y Satisfacción SRPA - Gobernación de Bolívar')

sheet_cola, sheet_respuestas = init_sheets()

if not sheet_cola or not sheet_respuestas:
    st.error("⚠️ Error de Conexión a Base de Datos")
    st.markdown("""
    ### ¿Cómo solucionar este error?
    Para que la aplicación funcione y se conecte a Google Sheets, necesitas proveer tus credenciales de Google Cloud en Streamlit Secrets.
    
    Sigue estos pasos rápidos:
    1. Crea un proyecto en **Google Cloud Console**.
    2. Activa **Google Sheets API** y **Google Drive API**.
    3. Crea una **Cuenta de Servicio**, genera una clave **JSON** y descárgala.
    4. Abre tu Google Sheet, nómbralo exactamente como **`Base_Encuestas_SRPA`** y comparte permisos de **Editor** con el correo de la cuenta de servicio.
    5. Ve al panel de control de tu aplicación en **Streamlit Cloud** -> **Settings** -> **Secrets** e introduce las credenciales en formato TOML:
    
    ```toml
    GEMINI_API_KEY = "tu_api_key_de_gemini"
    
    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "..."
    ...
    ```
    """)
else:
    # Definición de Pestañas
    tab1, tab2, tab3 = st.tabs(["📥 Cargar Nueva Encuesta", "✍️ Banco de Verificación", "📊 Dashboard Estadístico"])
    
    # --- PESTAÑA 1: CARGA DE ENCUESTAS ---
    with tab1:
        st.markdown("### 📥 Procesamiento de Encuestas de Doble Página")
        st.write("Sube la Página 1 y la Página 2 del cuestionario de forma obligatoria para realizar la integración con IA.")
        
        c_tipo = st.selectbox("Tipo de Encuesta a Subir", ["PRETEST", "POSTEST"])
        
        col_img1, col_img2 = st.columns(2)
        with col_img1:
            st.markdown("**Página 1 (Cabecera y Preguntas Iniciales)**")
            p1_file = st.file_uploader("Subir foto Página 1", type=["png", "jpg", "jpeg", "webp"], key="p1")
            if p1_file:
                st.image(p1_file, use_container_width=True)
                
        with col_img2:
            st.markdown("**Página 2 (Preguntas Finales o Satisfacción)**")
            p2_file = st.file_uploader("Subir foto Página 2", type=["png", "jpg", "jpeg", "webp"], key="p2")
            if p2_file:
                st.image(p2_file, use_container_width=True)
                
        if p1_file and p2_file:
            st.success("✅ ¡Páginas 1 y 2 cargadas correctamente! Listas para procesar.")
            btn_ocr = st.button("🔍 Procesar Encuesta con IA")
            
            if btn_ocr:
                with st.spinner("Procesando imagen con Gemini AI Vision..."):
                    img1_bytes = p1_file.read()
                    img2_bytes = p2_file.read()
                    
                    data_ia = run_gemini_ocr(img1_bytes, img2_bytes)
                    if data_ia:
                        # Generar ID único para la encuesta
                        enc_id = f"SRPA-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                        fecha_carga = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        enc_dict = data_ia.get("encabezado", {})
                        municipio = enc_dict.get("municipio", "")
                        ie_ia = enc_dict.get("institucion_educativa", "")
                        rol = enc_dict.get("rol", "")
                        
                        # Guardar en Cola de Revisión
                        try:
                            sheet_cola.append_row([
                                enc_id,
                                fecha_carga,
                                c_tipo,
                                municipio,
                                ie_ia,
                                rol,
                                json.dumps(data_ia),
                                "Pendiente"
                            ])
                            st.success("🎉 ¡Encuesta procesada exitosamente! Se ha guardado en la 'Cola de Revisión' para la verificación humana.")
                            st.balloons()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error guardando en la cola de revisión: {e}")
                            
    # --- PESTAÑA 2: BANCO DE VERIFICACIÓN ---
    with tab2:
        st.markdown("### ✍️ Banco de Verificación de Escritura a Mano")
        st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de consolidarla permanentemente.")
        
        df_cola = load_cola_data(sheet_cola)
        
        if df_cola.empty or len(df_cola[df_cola["Estado"] == "Pendiente"]) == 0:
            st.info("No hay encuestas pendientes de verificación. ¡Buen trabajo!")
        else:
            df_pending = df_cola[df_cola["Estado"] == "Pendiente"]
            st.success(f"Hay {len(df_pending)} encuesta(s) pendiente(s) de revisión.")
            
            survey_options = df_pending["ID_Encuesta"].tolist()
            sel_survey = st.selectbox("Selecciona la Encuesta a Revisar", survey_options)
            
            row_data = df_pending[df_pending["ID_Encuesta"] == sel_survey].iloc[0]
            
            try:
                raw_json = json.loads(row_data["JSON_Respuestas"])
                enc_ia = raw_json.get("encabezado", {})
                conocimiento = raw_json.get("respuestas_conocimiento", {})
                satisfaccion = raw_json.get("evaluacion_satisfaccion", {})
            except Exception as e:
                st.error(f"Error procesando JSON de respuestas: {e}")
                raw_json = {}
                enc_ia = {}
                conocimiento = {}
                satisfaccion = {}
                
            st.markdown("---")
            col_rev1, col_rev2 = st.columns(2)
            
            with col_rev1:
                st.markdown("#### Datos de Cabecera Leídos por la IA")
                st.write(f"**ID Encuesta:** {sel_survey}")
                st.write(f"**Tipo Formulario:** {row_data['Tipo_Formulario']}")
                st.write(f"**Fecha Captura:** {enc_ia.get('fecha', '')}")
                st.write(f"**Municipio:** {row_data['Municipio']}")
                st.write(f"**Rol:** {row_data['Rol']}")
                
                st.markdown("##### ✏️ Corrección de Caligrafía Manuscríba")
                ie_verificada = st.text_input(
                    "Nombre de la Institución Educativa (Revisar y Corregir):", 
                    value=row_data["Institucion_Educativa_IA"]
                )
                
            with col_rev2:
                st.markdown("#### Respuestas Extraídas de las Preguntas")
                # Mostrar respuestas de conocimiento
                st.write("**Conocimientos:**")
                st.json(conocimiento)
                
                # Mostrar respuestas de satisfacción si aplica
                if row_data["Tipo_Formulario"] == "POSTEST":
                    st.write("**Evaluación de Satisfacción:**")
                    st.json(satisfaccion)
                    
            st.markdown("---")
            btn_col1, btn_col2 = st.columns(2)
            
            with btn_col1:
                btn_approve = st.button("✅ Aprobar e Ingresar a Base de Datos", type="primary")
                if btn_approve:
                    with st.spinner("Guardando en la base de datos de producción..."):
                        # Construir la fila completa con celdas limpias
                        new_row = [
                            sel_survey,
                            row_data["Tipo_Formulario"],
                            enc_ia.get("fecha", ""),
                            row_data["Municipio"],
                            ie_verificada,
                            row_data["Rol"],
                            conocimiento.get("p1", ""),
                            conocimiento.get("p2", ""),
                            conocimiento.get("p3", ""),
                            conocimiento.get("p4", ""),
                            conocimiento.get("p5", ""),
                            conocimiento.get("p6", ""),
                            conocimiento.get("p7", ""),
                            conocimiento.get("p8", ""),
                            satisfaccion.get("s1", ""),
                            satisfaccion.get("s2", ""),
                            satisfaccion.get("s3", ""),
                            satisfaccion.get("s4", ""),
                            satisfaccion.get("s5", ""),
                            satisfaccion.get("s6", ""),
                            satisfaccion.get("s7", ""),
                            satisfaccion.get("s8", ""),
                            satisfaccion.get("s9", ""),
                            "Facilitador Campo",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ]
                        
                        try:
                            # 1. Insertar en Respuestas_SRPA
                            sheet_respuestas.append_row(new_row)
                            
                            # 2. Eliminar o actualizar en la Cola_Revision
                            records = sheet_cola.get_all_records()
                            row_idx = 2
                            for rec in records:
                                if rec["ID_Encuesta"] == sel_survey:
                                    sheet_cola.delete_rows(row_idx)
                                    break
                                row_idx += 1
                                
                            st.success(f"🎉 ¡Encuesta '{sel_survey}' aprobada con éxito y añadida a la base de datos consolidadas!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error guardando registro consolidado: {e}")
                            
            with btn_col2:
                btn_reject = st.button("❌ Rechazar Entrada (Eliminar de la Cola)")
                if btn_reject:
                    with st.spinner("Eliminando entrada de la cola..."):
                        try:
                            records = sheet_cola.get_all_records()
                            row_idx = 2
                            for rec in records:
                                if rec["ID_Encuesta"] == sel_survey:
                                    sheet_cola.delete_rows(row_idx)
                                    break
                                row_idx += 1
                            st.warning(f"Entrada '{sel_survey}' rechazada y eliminada de la cola.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error rechazando entrada: {e}")
                            
    # --- PESTAÑA 3: DASHBOARD ESTADÍSTICO ---
    with tab3:
        st.markdown("### 📊 Dashboard Estadístico en Tiempo Real")
        
        df_resp = load_respuestas_data(sheet_respuestas)
        
        if df_resp.empty:
            st.info("La base de datos de respuestas validadas está vacía. ¡Empieza a verificar encuestas en la pestaña 'Banco de Verificación' para ver estadísticas aquí!")
        else:
            st.subheader("Métricas de Cobertura e Impacto")
            
            # Filtros dinámicos en el dashboard
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                municipios = ["Todos"] + sorted(list(df_resp["Municipio"].dropna().unique()))
                muni_sel = st.selectbox("Filtrar por Municipio", municipios)
            with col_f2:
                roles = ["Todos"] + sorted(list(df_resp["Rol"].dropna().unique()))
                rol_sel = st.selectbox("Filtrar por Rol", roles)
                
            df_filtered = df_resp.copy()
            if muni_sel != "Todos":
                df_filtered = df_filtered[df_filtered["Municipio"] == muni_sel]
            if rol_sel != "Todos":
                df_filtered = df_filtered[df_filtered["Rol"] == rol_sel]
                
            if df_filtered.empty:
                st.warning("No hay datos para la combinación de filtros seleccionada.")
            else:
                # KPIs en la parte superior
                tot_enc = len(df_filtered)
                tot_pre = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
                tot_post = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
                
                kpi1, kpi2, kpi3 = st.columns(3)
                kpi1.metric("Total Encuestas Validadas", tot_enc)
                kpi2.metric("Pretests (Línea Base)", tot_pre)
                kpi3.metric("Postests (Evaluaciones)", tot_post)
                
                st.markdown("---")
                
                col_g1, col_g2 = st.columns(2)
                
                with col_g1:
                    st.markdown("#### Participación por Rol")
                    rol_counts = df_filtered["Rol"].value_counts().reset_index()
                    rol_counts.columns = ["Rol", "Cantidad"]
                    fig_rol = px.pie(rol_counts, names="Rol", values="Cantidad", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_rol.update_layout(margin=dict(l=20, r=20, t=30, b=20))
                    st.plotly_chart(fig_rol, use_container_width=True)
                    
                # Calcular aciertos para Pretest y Postest de conceptos comunes
                pre_df = df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"]
                post_df = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]
                
                topics = ["Finalidad SRPA", "Factor Riesgo", "Factor Protector", "Responsabilidad"]
                pre_scores = []
                post_scores = []
                
                # 1. Finalidad SRPA (Pretest P2 == "b", Postest P1 == "b")
                if len(pre_df) > 0:
                    pre_scores.append( (pre_df["Conocimientos_P2"].astype(str).str.lower().str.strip() == "b").mean() * 100 )
                else:
                    pre_scores.append(0)
                    
                if len(post_df) > 0:
                    post_scores.append( (post_df["Conocimientos_P1"].astype(str).str.lower().str.strip() == "b").mean() * 100 )
                else:
                    post_scores.append(0)
                    
                # 2. Factor Riesgo (Pretest P3 == "b", Postest P2 == "b")
                if len(pre_df) > 0:
                    pre_scores.append( (pre_df["Conocimientos_P3"].astype(str).str.lower().str.strip() == "b").mean() * 100 )
                else:
                    pre_scores.append(0)
                    
                if len(post_df) > 0:
                    post_scores.append( (post_df["Conocimientos_P2"].astype(str).str.lower().str.strip() == "b").mean() * 100 )
                else:
                    post_scores.append(0)
                    
                # 3. Factor Protector (Pretest P4 == "a", Postest P3 == "b")
                if len(pre_df) > 0:
                    pre_scores.append( (pre_df["Conocimientos_P4"].astype(str).str.lower().str.strip() == "a").mean() * 100 )
                else:
                    pre_scores.append(0)
                    
                if len(post_df) > 0:
                    post_scores.append( (post_df["Conocimientos_P3"].astype(str).str.lower().str.strip() == "b").mean() * 100 )
                else:
                    post_scores.append(0)
                    
                # 4. Responsabilidad (Pretest P5 == "d", Postest P4 == "d")
                if len(pre_df) > 0:
                    pre_scores.append( (pre_df["Conocimientos_P5"].astype(str).str.lower().str.strip() == "d").mean() * 100 )
                else:
                    pre_scores.append(0)
                    
                if len(post_df) > 0:
                    post_scores.append( (post_df["Conocimientos_P4"].astype(str).str.lower().str.strip() == "d").mean() * 100 )
                else:
                    post_scores.append(0)
                    
                with col_g2:
                    st.markdown("#### % Respuestas Correctas (Antes vs Después)")
                    fig_comp = go.Figure()
                    fig_comp.add_trace(go.Bar(
                        x=topics,
                        y=pre_scores,
                        name="PRETEST (Antes)",
                        marker_color="#1f77b4"
                    ))
                    fig_comp.add_trace(go.Bar(
                        x=topics,
                        y=post_scores,
                        name="POSTEST (Después)",
                        marker_color="#2ca02c"
                    ))
                    fig_comp.update_layout(
                        barmode="group",
                        yaxis_title="Porcentaje (%) de Aciertos",
                        yaxis_range=[0, 100],
                        margin=dict(l=20, r=20, t=30, b=20)
                    )
                    st.plotly_chart(fig_comp, use_container_width=True)
                    
                # Gráficos de Satisfacción (solo si hay Postests)
                if len(post_df) > 0:
                    st.markdown("---")
                    st.markdown("#### Evaluación de Satisfacción del Taller (Postest)")
                    
                    val_map = {"excelente": 4, "bueno": 3, "regular": 2, "deficiente": 1}
                    
                    sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
                    sat_names = [
                        "1. Claridad de Información",
                        "2. Dominio del Tema",
                        "3. Metodología",
                        "4. Participación",
                        "5. Utilidad de Temas",
                        "6. Organización",
                        "7. Materiales/Recursos",
                        "8. Fortaleció Conocimiento",
                        "9. Recomendaría Jornada"
                    ]
                    
                    sat_means = []
                    for col in sat_cols:
                        vals = post_df[col].astype(str).str.lower().str.strip().map(val_map)
                        mean_val = vals.mean()
                        if pd.isna(mean_val):
                            mean_val = 0
                        sat_means.append(mean_val)
                        
                    fig_sat = go.Figure()
                    fig_sat.add_trace(go.Bar(
                        x=sat_names,
                        y=sat_means,
                        marker_color="#ff7f0e",
                        name="Calificación Promedio"
                    ))
                    fig_sat.add_trace(go.Scatter(
                        x=sat_names,
                        y=[3.0]*9,
                        mode="lines",
                        name="Meta de Calidad (Bueno = 3.0)",
                        line=dict(color="red", dash="dash")
                    ))
                    fig_sat.update_layout(
                        yaxis_title="Calificación Promedio (1 a 4)",
                        yaxis_range=[1, 4],
                        xaxis_tickangle=-45,
                        margin=dict(l=20, r=20, t=30, b=100)
                    )
                    st.plotly_chart(fig_sat, use_container_width=True)
                else:
                    st.info("La evaluación de satisfacción solo se genera con registros de tipo POSTEST.")
