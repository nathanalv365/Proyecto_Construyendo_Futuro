import streamlit as st
import pandas as pd
import json
import os
import io
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Evaluación de Conocimientos SRPA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para celulares y diseño moderno
st.markdown("""
    <style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        border-left: 5px solid #007bff;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    .kpi-title {
        font-size: 0.9rem;
        color: #6c757d;
        font-weight: bold;
    }
    .kpi-value {
        font-size: 1.8rem;
        color: #212529;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- CONFIGURACIÓN E INICIALIZACIÓN DE GOOGLE SHEETS ---
def conectar_google_sheets():
    """Establece conexión con Google Sheets usando Secrets de Streamlit."""
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Limpiar saltos de línea en la llave privada
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        # Intentar abrir la hoja
        try:
            doc = client.open("Base_Encuestas_SRPA")
        except gspread.SpreadsheetNotFound:
            # Crear la hoja si no existe
            doc = client.create("Base_Encuestas_SRPA")
            # Compartir con el correo de la cuenta de servicio y el creador si estuviera especificado
            st.info("Creada nueva hoja de cálculo 'Base_Encuestas_SRPA' en Google Drive.")
        
        return doc
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
        return None

def inicializar_hojas(doc):
    """Asegura que existan las pestañas necesarias con sus columnas oficiales."""
    if doc is None:
        return
    
    # 1. Inicializar Cola_Revision
    try:
        sh_cola = doc.worksheet("Cola_Revision")
    except gspread.WorksheetNotFound:
        sh_cola = doc.add_worksheet("Cola_Revision", rows=1000, cols=8)
        sh_cola.append_row([
            "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
            "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
        ])
        
    # 2. Inicializar Respuestas_SRPA
    try:
        sh_resp = doc.worksheet("Respuestas_SRPA")
    except gspread.WorksheetNotFound:
        sh_resp = doc.add_worksheet("Respuestas_SRPA", rows=5000, cols=25)
        # Encabezado estructurado para almacenar las respuestas limpias
        headers = [
            "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
            "Institucion_Educativa_Verificada", "Rol",
            "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
            "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
            "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
            "Verificado_Por", "Fecha_Aprobacion"
        ]
        sh_resp.append_row(headers)

# Intentar conectar
google_doc = conectar_google_sheets()
if google_doc:
    inicializar_hojas(google_doc)

# --- CARGA DE DATOS DESDE GOOGLE SHEETS ---
def cargar_cola_revision():
    """Lee registros pendientes de validación."""
    if google_doc is None:
        return pd.DataFrame()
    try:
        sheet = google_doc.worksheet("Cola_Revision")
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            # Crear DataFrame con estructura esperada si la hoja está vacía
            df = pd.DataFrame(columns=[
                "ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", 
                "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado"
            ])
        return df[df["Estado"] == "Pendiente"]
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame()

def cargar_respuestas_validadas():
    """Lee registros validados finales."""
    if google_doc is None:
        return pd.DataFrame()
    try:
        sheet = google_doc.worksheet("Respuestas_SRPA")
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            df = pd.DataFrame(columns=[
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", 
                "Institucion_Educativa_Verificada", "Rol",
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
                "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                "Verificado_Por", "Fecha_Aprobacion"
            ])
        return df
    except Exception as e:
        st.error(f"Error cargando Base de Datos: {e}")
        return pd.DataFrame()

# --- LLAMADA A LA API DE GEMINI (SDK ACTUALIZADO AL MODELO GEMINI-2.0-FLASH) ---
def procesar_imagenes_con_gemini(img1_bytes, img2_bytes, api_key):
    """Envia ambas páginas de la encuesta a Gemini 2.0 Flash para consolidar la extracción."""
    # Intentar importar el cliente de Google GenAI
    # Google ha deprecado la disponibilidad de gemini-1.5-flash en API v1beta.
    # Usamos estrictamente el modelo moderno 'gemini-2.0-flash'.
    model_name = 'gemini-2.0-flash'
    
    prompt_instrucciones = """
    Eres un sistema de OCR e IA de alta precisión para el Proyecto "Construyendo Futuro" (SRPA) en Bolívar, Colombia.
    Analiza las dos imágenes adjuntas (Página 1 y Página 2 de la misma encuesta) y extrae la información.
    
    Reglas críticas de extracción:
    1. Determina el tipo de encuesta: "PRETEST" o "POSTEST" (Se lee en el encabezado de la Página 1).
    2. Lee la Fecha y el Municipio en el encabezado.
    3. Lee la "Institución Educativa" escrita a mano en la parte superior. Si la caligrafía es ilegible, escribe lo más aproximado posible.
    4. Identifica el Rol del Participante marcando con una X (Estudiante, Docente, Padre de Familia, Lider comunitario).
    5. Extrae las respuestas marcadas con una equis (X):
       - En PRETEST: Preguntas de conocimiento de 1 a 8.
       - En POSTEST SECCIÓN A: Preguntas de conocimiento de 1 a 5.
       - En POSTEST SECCIÓN B (Satisfacción): Preguntas de satisfacción de 1 a 9. Los valores posibles son: Excelente, Bueno, Regular, Deficiente.
    6. REGLA DE INTEGRIDAD: Si una pregunta o casilla está en blanco (sin marcar), escribe un string vacío "". NO asumas ni inventes respuestas que el participante no marcó.

    Responde ESTRICTAMENTE con un objeto JSON con este formato estructurado:
    {
      "tipo_formulario": "PRETEST",
      "encabezado": {
        "fecha": "30/07/2026",
        "municipio": "Cartagena",
        "institucion_educativa": "Nombre Detectado",
        "rol": "Estudiante"
      },
      "respuestas_conocimiento": {
        "p1": "a", "p2": "b", "p3": "b", "p4": "a",
        "p5": "d", "p6": "a", "p7": "a", "p8": "ICBF"
      },
      "evaluacion_satisfaccion": {
        "s1": "Excelente", "s2": "Excelente", "s3": "Bueno",
        "s4": "", "s5": "Excelente", "s6": "Bueno",
        "s7": "", "s8": "Excelente", "s9": "Excelente"
      }
    }
    """
    
    try:
        # Intentar conectar con el nuevo SDK de Google GenAI
        try:
            from google import genai
            from google.genai import types
            
            client = genai.Client(api_key=api_key)
            
            # Convertir imágenes de bytes a partes multimedia
            part1 = types.Part.from_bytes(data=img1_bytes, mime_type="image/jpeg")
            part2 = types.Part.from_bytes(data=img2_bytes, mime_type="image/jpeg")
            
            response = client.models.generate_content(
                model=model_name,
                contents=[part1, part2, prompt_instrucciones],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            return json.loads(response.text)
            
        except ImportError:
            # Fallback al SDK clásico de google-generativeai
            import google.generativeai as genai_legacy
            genai_legacy.configure(api_key=api_key)
            
            img1 = Image.open(io.BytesIO(img1_bytes))
            img2 = Image.open(io.BytesIO(img2_bytes))
            
            model = genai_legacy.GenerativeModel(model_name)
            response = model.generate_content([img1, img2, prompt_instrucciones])
            
            # Limpiar posibles bloques de código markdown ```json ... ```
            text_cleaned = response.text.strip()
            if text_cleaned.startswith("```json"):
                text_cleaned = text_cleaned[7:]
            if text_cleaned.endswith("```"):
                text_cleaned = text_cleaned[:-3]
            
            return json.loads(text_cleaned.strip())
            
    except Exception as e:
        raise ValueError(f"Error procesando con Gemini ({model_name}): {e}")

# --- COMPRESIÓN DE IMÁGENES ANTES DE LA API (Para ahorrar ancho de banda móvil) ---
def optimizar_imagen(uploaded_file):
    """Comprime y redimensiona la imagen para optimizar el envío de datos móviles."""
    if uploaded_file is None:
        return None
    image = Image.open(uploaded_file)
    # Convertir a RGB si tiene canal Alfa
    if image.mode in ("RGBA", "P"):
        image = image.convert("RGB")
    
    # Redimensionar conservando el aspecto si excede un límite de seguridad
    max_size = 1600
    if max(image.size) > max_size:
        image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    
    # Guardar en memoria comprimida
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=80)
    return buffer.getvalue()

# --- INTERFAZ DE USUARIO (STREAMLIT) ---

if google_doc is None:
    st.title("⚙️ Configuración Requerida")
    st.warning("La aplicación móvil está lista, pero requiere conexión a Google Sheets para operar de forma colaborativa.")
    st.markdown("""
    ### Pasos para conectar Google Sheets en Streamlit Cloud:
    1. Genera un archivo de credenciales de Cuenta de Servicio en tu consola de Google Cloud.
    2. Comparte tu archivo de Google Sheets llamado exactamente **Base_Encuestas_SRPA** con el correo electrónico de la cuenta de servicio como **Editor**.
    3. Pega la estructura de secretos de tu cuenta de servicio en el panel de **Secrets** de Streamlit Cloud.
    """)
    st.stop()

# Menú principal superior estilo móvil
tabs = st.tabs(["📸 Cargar Encuestas", "✏️ Banco de Verificación", "📈 Dashboard Estadístico"])

# TAB 1: CARGA DE ENCUESTAS
with tabs[0]:
    st.title("📸 Captura de Encuestas")
    st.write("Sube las dos hojas pertenecientes a la misma encuesta para consolidarlas.")
    
    col1, col2 = st.columns(2)
    with col1:
        img_pag1 = st.file_uploader("Subir foto de la Página 1 (Cabecera y Preguntas 1-3)", type=["jpg", "jpeg", "png"], key="upload_p1")
    with col2:
        img_pag2 = st.file_uploader("Subir foto de la Página 2 (Preguntas de salida y Satisfacción)", type=["jpg", "jpeg", "png"], key="upload_p2")
    
    api_key_disponible = "GEMINI_API_KEY" in st.secrets
    api_key_manual = ""
    if not api_key_disponible:
        api_key_manual = st.text_input("Ingresa tu clave de Gemini API (Google AI Studio)", type="password")
    
    if st.button("🚀 Enviar a Procesamiento", use_container_width=True):
        if img_pag1 is None or img_pag2 is None:
            st.error("Por favor sube las fotos de ambas páginas de la encuesta para realizar la fusión.")
        else:
            final_key = st.secrets["GEMINI_API_KEY"] if api_key_disponible else api_key_manual
            if not final_key:
                st.error("Se requiere una API Key de Gemini para realizar la digitalización automática.")
            else:
                with st.spinner("La IA de Gemini está analizando ambas páginas de la encuesta..."):
                    try:
                        # Optimizar y comprimir imágenes antes del envío
                        p1_bytes = optimizar_imagen(img_pag1)
                        p2_bytes = optimizar_imagen(img_pag2)
                        
                        # Extraer datos dinámicamente con Gemini 2.0 Flash
                        resultado_ia = procesar_imagenes_con_gemini(p1_bytes, p2_bytes, final_key)
                        
                        # Guardar temporalmente en la cola de revisión
                        id_encuesta = f"SRPA-{int(pd.Timestamp.now().timestamp())}"
                        encabezado = resultado_ia.get("encabezado", {})
                        
                        sheet_cola = google_doc.worksheet("Cola_Revision")
                        sheet_cola.append_row([
                            id_encuesta,
                            pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                            resultado_ia.get("tipo_formulario", "PRETEST"),
                            encabezado.get("municipio", ""),
                            encabezado.get("institucion_educativa", ""),
                            encabezado.get("rol", "Estudiante"),
                            json.dumps(resultado_ia),
                            "Pendiente"
                        ])
                        
                        st.success("¡Encuesta leída con éxito! El nombre manuscrito de la Institución Educativa ha sido enviado a la Cola de Revisión para tu visto bueno.")
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error procesando la encuesta: {e}")

# TAB 2: BANCO DE VERIFICACIÓN
with tabs[1]:
    st.title("✏️ Banco de Verificación")
    st.write("Verifica y corrige la escritura de la Institución Educativa antes de enviarla a la base de datos definitiva.")
    
    df_cola = cargar_cola_revision()
    
    if df_cola.empty:
        st.success("🎉 ¡Excelente! No hay encuestas pendientes de verificación de caligrafía en este momento.")
    else:
        st.warning(f"Tienes {len(df_cola)} encuestas en cola pendientes de aprobación.")
        
        # Tomar la primera encuesta en cola para procesar
        fila = df_cola.iloc[0]
        id_encuesta = fila["ID_Encuesta"]
        ie_ia = fila["Institucion_Educativa_IA"]
        rol = fila["Rol"]
        tipo = fila["Tipo_Formulario"]
        fecha_carga = fila["Fecha_Carga"]
        
        st.info(f"Mostrando encuesta: **{id_encuesta}** | Tipo: **{tipo}** | Cargado el: **{fecha_carga}**")
        
        # Módulo editable de validación
        nueva_ie = st.text_input("📝 Confirmar / Corregir Institución Educativa:", value=ie_ia)
        
        c_verif1, c_verif2 = st.columns(2)
        with c_verif1:
            if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True):
                # Extraer respuestas completas almacenadas en el JSON temporal
                try:
                    datos_completos = json.loads(fila["JSON_Respuestas"])
                except Exception:
                    datos_completos = {}
                
                resp_c = datos_completos.get("respuestas_conocimiento", {})
                resp_s = datos_completos.get("evaluacion_satisfaccion", {})
                encabezado = datos_completos.get("encabezado", {})
                
                # Consolidar fila final en Respuestas_SRPA
                fila_final = [
                    id_encuesta,
                    tipo,
                    encabezado.get("fecha", ""),
                    encabezado.get("municipio", ""),
                    nueva_ie, # Nombre verificado por el humano
                    rol,
                    # Respuestas de conocimiento (p1 a p8)
                    resp_c.get("p1", ""), resp_c.get("p2", ""), resp_c.get("p3", ""), resp_c.get("p4", ""),
                    resp_c.get("p5", ""), resp_c.get("p6", ""), resp_c.get("p7", ""), resp_c.get("p8", ""),
                    # Respuestas de satisfacción (s1 a s9)
                    resp_s.get("s1", ""), resp_s.get("s2", ""), resp_s.get("s3", ""), resp_s.get("s4", ""),
                    resp_s.get("s5", ""), resp_s.get("s6", ""), resp_s.get("s7", ""), resp_s.get("s8", ""),
                    resp_s.get("s9", ""),
                    "Facilitador Campo",
                    pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                ]
                
                # Guardar en la hoja final
                sheet_resp = google_doc.worksheet("Respuestas_SRPA")
                sheet_resp.append_row(fila_final)
                
                # Marcar como Aprobado en la cola
                sheet_cola = google_doc.worksheet("Cola_Revision")
                # Buscar fila por ID_Encuesta para actualizar el Estado
                celda = sheet_cola.find(id_encuesta)
                if celda:
                    sheet_cola.update_cell(celda.row, 8, "Aprobado") # Columna 8 es 'Estado'
                
                st.success(f"Encuesta {id_encuesta} validada y archivada con éxito.")
                st.rerun()
                
        with c_verif2:
            if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                # Buscar y marcar como Rechazado en la hoja de cola
                sheet_cola = google_doc.worksheet("Cola_Revision")
                celda = sheet_cola.find(id_encuesta)
                if celda:
                    sheet_cola.update_cell(celda.row, 8, "Rechazado")
                st.warning(f"Entrada {id_encuesta} rechazada y eliminada de la cola.")
                st.rerun()

# TAB 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
with tabs[2]:
    st.title("📈 Dashboard Estadístico")
    st.write("Visualización agregada del impacto y cobertura del Proyecto 'Construyendo Futuro'.")
    
    df_resp = cargar_respuestas_validadas()
    
    if df_resp.empty:
        st.info("La base de datos se encuentra limpia y vacía. Los gráficos estadísticos se renderizarán tan pronto como comiences a aprobar encuestas en el Banco de Verificación.")
    else:
        # Filtros interactivos de barra lateral
        st.sidebar.header("Filtros del Proyecto")
        list_municipios = ["Todos"] + sorted(df_resp["Municipio"].unique().tolist())
        filtro_mun = st.sidebar.selectbox("Seleccionar Municipio", list_municipios)
        
        list_instituciones = ["Todas"] + sorted(df_resp["Institucion_Educativa_Verificada"].unique().tolist())
        filtro_ie = st.sidebar.selectbox("Seleccionar Institución Educativa", list_instituciones)
        
        # Aplicar filtros
        df_filtrado = df_resp.copy()
        if filtro_mun != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Municipio"] == filtro_mun]
        if filtro_ie != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Institucion_Educativa_Verificada"] == filtro_ie]
            
        if df_filtrado.empty:
            st.warning("No hay datos coincidentes con los filtros seleccionados.")
        else:
            # Tarjetas KPI principales
            kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
            with kpi_col1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">TOTAL REGISTROS</div>
                    <div class="kpi-value">{len(df_filtrado)}</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi_col2:
                num_pre = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"])
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #ffc107;">
                    <div class="kpi-title">TOTAL PRETEST</div>
                    <div class="kpi-value">{num_pre}</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi_col3:
                num_post = len(df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"])
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #28a745;">
                    <div class="kpi-title">TOTAL POSTEST</div>
                    <div class="kpi-value">{num_post}</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi_col4:
                num_cola = len(df_cola) if 'df_cola' in locals() else 0
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #dc3545;">
                    <div class="kpi-title">EN COLA REVISIÓN</div>
                    <div class="kpi-value">{num_cola}</div>
                </div>
                """, unsafe_allow_html=True)
                
            st.markdown("---")
            
            # Gráfico de Impacto Educativo (Preguntas comunes Pre vs Post)
            st.subheader("🎯 Medición de Impacto Educativo (Antes vs. Después)")
            
            # Agrupar respuestas correctas en las 4 preguntas de conocimiento comunes
            # P2/Sección A1: Finalidad del SRPA (Correcta: "b" o texto aproximado)
            # P3/Sección A2: Factor de riesgo (Correcta: "b")
            # P4/Sección A3: Factor protector (Correcta: "b")
            # P5/Sección A4: Responsable de prevención (Correcta: "d")
            
            pre_df = df_filtrado[df_filtrado["Tipo_Formulario"] == "PRETEST"]
            post_df = df_filtrado[df_filtrado["Tipo_Formulario"] == "POSTEST"]
            
            if pre_df.empty or post_df.empty:
                st.info("Para comparar el impacto académico del taller, se requiere contar con al menos un registro de Pretest y un registro de Postest en la base de datos.")
            else:
                preguntas_labels = [
                    "Finalidad SRPA (P2)", 
                    "Factores Riesgo (P3)", 
                    "Factores Protección (P4)", 
                    "Corresponsabilidad (P5)"
                ]
                
                # Calcular tasas de acierto aproximadas
                aciertos_pre = []
                aciertos_post = []
                
                # P2
                p2_pre_correct = (pre_df["Conocimientos_P2"].astype(str).str.lower().str.startswith("b") | pre_df["Conocimientos_P2"].astype(str).str.lower().str.contains("promover")).mean() * 100
                p2_post_correct = (post_df["Conocimientos_P1"].astype(str).str.lower().str.startswith("b") | post_df["Conocimientos_P1"].astype(str).str.lower().str.contains("promover")).mean() * 100
                aciertos_pre.append(p2_pre_correct)
                aciertos_post.append(p2_post_correct)
                
                # P3
                p3_pre_correct = (pre_df["Conocimientos_P3"].astype(str).str.lower().str.startswith("b") | pre_df["Conocimientos_P3"].astype(str).str.lower().str.contains("consumir")).mean() * 100
                p3_post_correct = (post_df["Conocimientos_P2"].astype(str).str.lower().str.startswith("b") | post_df["Conocimientos_P2"].astype(str).str.lower().str.contains("consumir")).mean() * 100
                aciertos_pre.append(p3_pre_correct)
                aciertos_post.append(p3_post_correct)
                
                # P4
                p4_pre_correct = (pre_df["Conocimientos_P4"].astype(str).str.lower().str.startswith("a") | pre_df["Conocimientos_P4"].astype(str).str.lower().str.contains("dialogar")).mean() * 100
                p4_post_correct = (post_df["Conocimientos_P3"].astype(str).str.lower().str.startswith("b") | post_df["Conocimientos_P3"].astype(str).str.lower().str.contains("participar")).mean() * 100
                aciertos_pre.append(p4_pre_correct)
                aciertos_post.append(p4_post_correct)
                
                # P5
                p5_pre_correct = (pre_df["Conocimientos_P5"].astype(str).str.lower().str.startswith("d") | pre_df["Conocimientos_P5"].astype(str).str.lower().str.contains("familia, la escuela")).mean() * 100
                p5_post_correct = (post_df["Conocimientos_P4"].astype(str).str.lower().str.startswith("d") | post_df["Conocimientos_P4"].astype(str).str.lower().str.contains("familia, la escuela")).mean() * 100
                aciertos_pre.append(p5_pre_correct)
                aciertos_post.append(p5_post_correct)
                
                fig_impacto = go.Figure(data=[
                    go.Bar(name="Pretest (Antes)", x=preguntas_labels, y=aciertos_pre, marker_color='#ffc107'),
                    go.Bar(name="Postest (Después)", x=preguntas_labels, y=aciertos_post, marker_color='#28a745')
                ])
                fig_impacto.update_layout(
                    barmode='group',
                    yaxis_title='Porcentaje de Respuestas Correctas (%)',
                    yaxis_range=[0, 100],
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_impacto, use_container_width=True)
                
            st.markdown("---")
            
            # Segunda Fila de Gráficos
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.subheader("👥 Distribución de Participantes")
                rol_counts = df_filtrado["Rol"].value_counts().reset_index()
                rol_counts.columns = ["Rol", "Cantidad"]
                fig_roles = px.pie(rol_counts, values="Cantidad", names="Rol", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_roles, use_container_width=True)
                
            with g_col2:
                st.subheader("⭐️ Nivel de Satisfacción (Postest)")
                # Filtrar solo postests que tengan respuestas de satisfacción
                post_sat = df_filtrado[(df_filtrado["Tipo_Formulario"] == "POSTEST") & (df_filtrado["Sat_P1"] != "")]
                
                if post_sat.empty:
                    st.info("No se han registrado respuestas de satisfacción en los Postests validados.")
                else:
                    satisfaccion_headers = [
                        "Claridad Info", "Dominio Tema", "Metodología", 
                        "Participación", "Utilidad Temas", "Organización", 
                        "Materiales", "Fortalecimiento", "Recomendaría"
                    ]
                    
                    # Convertir Excelente/Bueno/Regular/Deficiente a valores numéricos (4, 3, 2, 1) para promedio
                    map_sat = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1, "": None}
                    
                    promedios = []
                    for i in range(1, 10):
                        col_name = f"Sat_P{i}"
                        # Mapear valores y sacar promedio omitiendo vacíos
                        valores_numericos = post_sat[col_name].map(map_sat).dropna()
                        promedio = valores_numericos.mean() if not valores_numericos.empty else 0
                        promedios.append(promedio)
                        
                    fig_sat = go.Figure()
                    fig_sat.add_trace(go.Bar(
                        x=satisfaccion_headers,
                        y=promedios,
                        marker_color='#17a2b8',
                        name="Calificación Promedio (1-4)"
                    ))
                    # Añadir línea meta de satisfacción mínima (3.0 = Bueno)
                    fig_sat.add_shape(
                        type="line", line=dict(dash="dash", color="red", width=2),
                        x0=-0.5, x1=8.5, y0=3.0, y1=3.0
                    )
                    fig_sat.update_layout(
                        yaxis_title='Calificación Promedio',
                        yaxis_range=[0, 4],
                        title_text="Meta mínima recomendada: 3.0 (Bueno)",
                        title_font_size=12
                    )
                    st.plotly_chart(fig_sat, use_container_width=True)
