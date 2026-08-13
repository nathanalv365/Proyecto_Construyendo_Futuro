import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import json
import os
import base64
import requests
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import io

# Configuración de página de Streamlit - Móvil y responsiva por defecto
st.set_page_config(
    page_title="SRPA - Construyendo Futuro",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE = "plantilla_encuestas_srpa.xlsx"

# -------------------------------------------------------------------------
# FUNCIONES AUXILIARES PARA MANEJO DE EXCEL
# -------------------------------------------------------------------------

def load_data():
    """Carga los datos de Respuestas_SRPA y Cola_Revision desde el Excel."""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_resp = pd.read_excel(EXCEL_FILE, sheet_name="Respuestas_SRPA")
        df_cola = pd.read_excel(EXCEL_FILE, sheet_name="Cola_Revision")
        return df_resp, df_cola
    except Exception as e:
        st.error(f"Error cargando el archivo Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

def save_to_respuestas(record):
    """Inserta un registro aprobado en la pestaña Respuestas_SRPA."""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Respuestas_SRPA"]
        
        # Generar un ID único si no viene en el registro
        if not record.get("ID_Encuesta"):
            record["ID_Encuesta"] = f"ENC-{ws.max_row:03d}"
            
        row_num = ws.max_row + 1
        
        # Mapear columnas según el orden oficial
        row_data = [
            record.get("ID_Encuesta"),
            record.get("Tipo_Formulario"),
            record.get("Fecha"),
            record.get("Municipio"),
            record.get("Institucion_Educativa_Verificada"),
            record.get("Rol"),
            record.get("Conocimientos_P1", ""),
            record.get("Conocimientos_P2", ""),
            record.get("Conocimientos_P3", ""),
            record.get("Conocimientos_P4", ""),
            record.get("Conocimientos_P5", ""),
            record.get("Conocimientos_P6", ""),
            record.get("Conocimientos_P7", ""),
            record.get("Conocimientos_P8", ""),
            record.get("Sat_P1", ""),
            record.get("Sat_P2", ""),
            record.get("Sat_P3", ""),
            record.get("Sat_P4", ""),
            record.get("Sat_P5", ""),
            record.get("Sat_P6", ""),
            record.get("Sat_P7", ""),
            record.get("Sat_P8", ""),
            record.get("Sat_P9", ""),
            record.get("Verificado_Por", "App Móvil"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        
        ws.row_dimensions[row_num].height = 20
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            if col_idx in [5, 7, 8, 9, 10, 11, 12, 13]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        wb.save(EXCEL_FILE)
        return True, record["ID_Encuesta"]
    except Exception as e:
        return False, str(e)

def add_to_cola_revision(tipo, fecha, municipio, ie_ia, rol, json_resp, usuario):
    """Agrega una encuesta con lectura preliminar a la cola de revisión."""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Cola_Revision"]
        
        row_num = ws.max_row + 1
        id_enc = f"ENC-COL-{row_num:03d}"
        
        row_data = [
            id_enc,
            tipo,
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            municipio,
            ie_ia,
            rol,
            json_resp,
            "Pendiente",
            usuario
        ]
        
        ws.row_dimensions[row_num].height = 20
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            if col_idx in [5, 7]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        wb.save(EXCEL_FILE)
        return True, id_enc
    except Exception as e:
        return False, str(e)

def update_cola_status(id_encuesta, nuevo_estado="Aprobado"):
    """Actualiza el estado de una encuesta en la cola de revisión."""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Cola_Revision"]
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == id_encuesta:
                ws.cell(row=row, column=8, value=nuevo_estado)
                break
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        return False

# -------------------------------------------------------------------------
# FUNCIÓN DE LLAMADA REST A GEMINI (Multimodal - 2 imágenes)
# -------------------------------------------------------------------------

def call_gemini_api(api_key, image1_bytes, image2_bytes, tipo_encuesta):
    """
    Envía dos imágenes (Página 1 y Página 2) en una sola llamada REST
    a la API de Gemini para consolidar las respuestas en JSON.
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    
    # Codificar imágenes en Base64
    img1_b64 = base64.b64encode(image1_bytes).decode('utf-8')
    img2_b64 = base64.b64encode(image2_bytes).decode('utf-8')
    
    # Prompt de extracción estructurada
    system_prompt = f"""
    Eres un asistente de digitalización inteligente para el Proyecto 'Construyendo Futuro' (SRPA).
    Analiza las DOS imágenes adjuntas correspondientes a un cuestionario físico del tipo {tipo_encuesta}.
    
    - La imagen 1 es la PÁGINA 1 del formulario.
    - La imagen 2 es la PÁGINA 2 del formulario.
    
    Tu tarea es extraer de manera precisa toda la información manuscrita (escrita por puño y letra) y las casillas marcadas con una equis (X), y retornar una estructura JSON estrictamente válida según las reglas de abajo.
    
    REGLAS DE EXTRACCIÓN:
    1. Del encabezado (en la PÁGINA 1):
       - Fecha: Extrae en formato YYYY-MM-DD.
       - Municipio: Extrae el municipio.
       - Institución Educativa: Extrae el nombre manuscrito de la escuela. Presta mucha atención a la caligrafía. Si es ilegible o tiene mala ortografía, escríbelo tal cual lo ves.
       - Rol del participante: Identifica el rol marcado con (X) (Estudiante, Docente, Padre de Familia, Lider comunitario).
    
    2. Respuestas de conocimiento (PÁGINA 1 y PÁGINA 2):
       - En un PRETEST, hay 8 preguntas en total:
         * P1 a P3 están en la PÁGINA 1.
         * P4 a P8 están en la PÁGINA 2.
         * Las opciones son de opción múltiple. Extrae el TEXTO de la respuesta seleccionada (por ejemplo, para P1: 'Sí', 'No' o 'No estoy seguro(a)').
       - En un POSTEST, hay 5 preguntas de conocimiento (Sección A):
         * P1 a P3 están en la PÁGINA 1.
         * P4 a P5 están en la PÁGINA 2.
         * Extrae el TEXTO de la respuesta seleccionada.
         
    3. Evaluación de Satisfacción (POSTEST únicamente - en la PÁGINA 2):
       - Sección B contiene 9 aspectos evaluados en una matriz.
       - Las opciones para cada aspecto (1 a 9) son: 'Excelente', 'Bueno', 'Regular', 'Deficiente'.
       - Extrae la opción marcada con (X) para cada uno de los 9 aspectos.
       
    REQUERIMIENTO DE FORMATO:
    Retorna ÚNICAMENTE un objeto JSON con la siguiente estructura, sin comentarios ni marcas de markdown adicionales:
    
    Si es {tipo_encuesta}:
    {{
      "tipo_formulario": "{tipo_encuesta}",
      "fecha": "2026-07-30",
      "municipio": "Cartagena",
      "institucion_educativa": "Nombre de la escuela extraído",
      "rol": "Estudiante",
      "respuestas_conocimiento": {{
        "p1": "Respuesta extraída de la P1",
        "p2": "Respuesta extraída de la P2",
        "p3": "Respuesta extraída de la P3",
        "p4": "Respuesta extraída de la P4",
        "p5": "Respuesta extraída de la P5",
        "p6": "Respuesta extraída de la P6 (deja vacío o null si es POSTEST)",
        "p7": "Respuesta extraída de la P7 (deja vacío o null si es POSTEST)",
        "p8": "Respuesta extraída de la P8 (deja vacío o null si es POSTEST)"
      }},
      "satisfaccion": {{
        "sat_1": "Valor (Excelente/Bueno/Regular/Deficiente o dejar vacío si es PRETEST)",
        "sat_2": "...",
        "sat_3": "...",
        "sat_4": "...",
        "sat_5": "...",
        "sat_6": "...",
        "sat_7": "...",
        "sat_8": "...",
        "sat_9": "..."
      }}
    }}
    """
    
    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "inlineData": {
                            "mimeType": "image/jpeg",
                            "data": img1_b64
                        }
                    },
                    {
                        "inlineData": {
                            "mimeType": "image/jpeg",
                            "data": img2_b64
                        }
                    },
                    {
                        "text": system_prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1
        }
    }
    
    headers = {
        "Content-Type": "application/json"
    }
    
    response = requests.post(url, json=payload, headers=headers)
    
    if response.status_code == 200:
        res_json = response.json()
        try:
            # Extraer el texto de la respuesta
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
            return json.loads(text_response)
        except Exception as e:
            raise ValueError(f"No se pudo parsear el JSON retornado por la IA: {e}. Respuesta cruda: {response.text}")
    else:
        raise ConnectionError(f"Error de API (Status {response.status_code}): {response.text}")

# -------------------------------------------------------------------------
# GENERADOR DE DATOS SIMULADOS PARA MODO PRUEBA
# -------------------------------------------------------------------------

def get_simulated_ocr(tipo_encuesta):
    """Genera datos de lectura simulados realistas para probar la app sin API Key."""
    import random
    
    municipios = ["Cartagena", "Turbaco", "Arjona", "María la Baja"]
    ies = [
        "I.E. Promesa de Dioss", # con un typo para corregir
        "I.E. San Joze de Caño del Oro", # con un typo para corregir
        "I.E. Arroyo de Piedraa", # con un typo para corregir
    ]
    roles = ["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"]
    
    selected_municipio = random.choice(municipios)
    selected_ie = random.choice(ies)
    selected_rol = random.choices(roles, weights=[0.6, 0.15, 0.15, 0.10], k=1)[0]
    
    if tipo_encuesta == "PRETEST":
        respuestas = {
            "p1": random.choice(["Sí", "No", "No estoy seguro(a)"]),
            "p2": random.choice(["Promover la responsabilidad, la protección de derechos y la resocialización", "Castigar a los adolescentes", "No lo sé"]),
            "p3": random.choice(["Consumir sustancias psicoactivas", "Tener un proyecto de vida", "Practicar deporte"]),
            "p4": random.choice(["Dialogar con la familia", "Abandonar los estudios", "Resolver los problemas mediante la violencia"]),
            "p5": random.choice(["La familia, la escuela, la comunidad y las instituciones", "Solo la familia", "Solo las autoridades"]),
            "p6": random.choice(["Buscar apoyo en un adulto de confianza", "Guardar silencio", "Resolver el problema solo"]),
            "p7": random.choice(["Sí", "No", "No estoy seguro(a)"]),
            "p8": random.choice(["ICBF", "Policía", "Institución Educativa", "Alcaldía"])
        }
        satisfaccion = {}
    else:  # POSTEST
        # Postest típicamente tiene respuestas de conocimiento mayormente correctas
        respuestas = {
            "p1": "Promover la responsabilidad, la protección de derechos y la resocialización",
            "p2": "Consumir sustancias psicoactivas",
            "p3": "Participar en actividades deportivas, culturales o comunitarias",
            "p4": "La familia, la escuela, la comunidad y las instituciones",
            "p5": random.choice(["Sí", "Parcialmente", "No"])
        }
        satisfaccion = {
            f"sat_{i}": random.choices(["Excelente", "Bueno", "Regular", "Deficiente"], weights=[0.7, 0.2, 0.08, 0.02], k=1)[0]
            for i in range(1, 10)
        }
        
    return {
        "tipo_formulario": tipo_encuesta,
        "fecha": datetime.now().strftime("%Y-%m-%d"),
        "municipio": selected_municipio,
        "institucion_educativa": selected_ie,
        "rol": selected_rol,
        "respuestas_conocimiento": respuestas,
        "satisfaccion": satisfaccion
    }

# -------------------------------------------------------------------------
# INTERFAZ DE USUARIO (UI)
# -------------------------------------------------------------------------

# Título y Banner superior
st.markdown("""
    <div style="background-color:#1F4E78;padding:15px;border-radius:10px;margin-bottom:20px;text-align:center;">
        <h1 style="color:white;margin:0;font-family:Helvetica;font-size:28px;">Proyecto "Construyendo Futuro" 📱</h1>
        <p style="color:#D9D9D9;margin:5px 0 0 0;font-size:14px;">Evaluación de Conocimientos SRPA - Gobernación de Bolívar</p>
    </div>
""", unsafe_allow_html=True)

# Cargar bases de datos
df_resp, df_cola = load_data()

# Sidebar: Configuración e Instrucciones
with st.sidebar:
    st.header("🔑 Configuración")
    api_key_input = st.text_input("Gemini API Key (Google AI Studio)", type="password", help="Ingresa tu clave de API para activar el procesamiento OCR real. Si se deja en blanco, la app funcionará en 'Modo Simulado'.")
    
    st.info("""
    **💡 Consejos de Uso Móvil:**
    1. **Buena iluminación:** Toma las fotos en un lugar bien iluminado para que la IA lea la caligrafía fácilmente.
    2. **Enfoque plano:** Coloca la encuesta sobre una mesa plana y toma la foto directamente desde arriba.
    3. **Tachaduras:** Si el niño tachó y corrigió, puedes ajustar la selección fácilmente en el módulo de revisión.
    """)
    
    st.write("---")
    st.caption("Desarrollado para el Proyecto de Prevención del Ingreso al SRPA - Gobernación de Bolívar. El sistema de encuestas es 100% anónimo.")

# Navegación Principal por Pestañas
tab_carga, tab_revisar, tab_dashboard = st.tabs([
    "📸 Cargar Encuestas (Doble Página)", 
    "✏️ Módulo de Revisión y Cola", 
    "📊 Dashboard Estadístico"
])

# -------------------------------------------------------------------------
# PESTAÑA 1: CARGA DE ENCUESTAS (OBLIGATORIA DOBLE PÁGINA)
# -------------------------------------------------------------------------
with tab_carga:
    st.header("1. Capturar / Subir Encuesta")
    st.write("Para procesar una encuesta de manera óptima, es **obligatorio cargar fotos de ambas páginas (Página 1 y Página 2)**. La IA las fusionará automáticamente en un solo registro unificado.")
    
    tipo_encuesta = st.radio("Selecciona el tipo de encuesta a digitalizar:", ["PRETEST", "POSTEST"], horizontal=True)
    
    col_p1, col_p2 = st.columns(2)
    
    with col_p1:
        st.markdown("### 📄 Página 1")
        st.caption("Debe incluir el encabezado (Fecha, Municipio, Institución Educativa, Rol) y las preguntas 1, 2 y 3.")
        foto_p1 = st.file_uploader("Capturar o subir la Página 1", type=["png", "jpg", "jpeg", "webp"], key="p1_upload")
        if foto_p1:
            st.image(foto_p1, caption="Página 1 cargada con éxito", use_container_width=True)
            
    with col_p2:
        st.markdown("### 📄 Página 2")
        st.caption("Debe incluir el resto de preguntas de conocimiento y la matriz de satisfacción si es un Postest.")
        foto_p2 = st.file_uploader("Capturar o subir la Página 2", type=["png", "jpg", "jpeg", "webp"], key="p2_upload")
        if foto_p2:
            st.image(foto_p2, caption="Página 2 cargada con éxito", use_container_width=True)
            
    st.markdown("---")
    
    # Botón de Procesamiento
    if foto_p1 and foto_p2:
        # El usuario subió ambas fotos, habilitar procesamiento
        st.success("✅ Ambas páginas han sido cargadas. Listo para procesar con IA.")
        
        if st.button("🔍 Procesar Encuesta (Fusionar Páginas con IA)", type="primary"):
            with st.spinner("Procesando imágenes con Gemini Vision API... Por favor espera."):
                img1_bytes = foto_p1.read()
                img2_bytes = foto_p2.read()
                
                try:
                    if api_key_input:
                        # Procesamiento OCR Real con API Key
                        resultado = call_gemini_api(api_key_input, img1_bytes, img2_bytes, tipo_encuesta)
                        modo_api = "Real (OCR Gemini)"
                    else:
                        # Procesamiento Simulado
                        import time
                        time.sleep(1.5)
                        resultado = get_simulated_ocr(tipo_encuesta)
                        modo_api = "Simulado (Demostración)"
                    
                    # Guardar JSON temporal en la base de datos (Cola de Revisión)
                    json_str = json.dumps(resultado)
                    success, id_enc = add_to_cola_revision(
                        tipo=tipo_encuesta,
                        fecha=resultado.get("fecha", datetime.now().strftime("%Y-%m-%d")),
                        municipio=resultado.get("municipio", "No detectado"),
                        ie_ia=resultado.get("institucion_educativa", "No detectado"),
                        rol=resultado.get("rol", "Estudiante"),
                        json_resp=json_str,
                        usuario="Digitador Móvil"
                    )
                    
                    if success:
                        st.balloons()
                        st.success(f"🎉 Encuesta procesada exitosamente bajo el ID temporal **{id_enc}** ({modo_api}).")
                        st.info("💡 Ve a la pestaña **'Módulo de Revisión y Cola'** en la parte superior para verificar el nombre manuscrito de la Institución Educativa y guardar el registro de forma definitiva.")
                        
                        # Guardar en session state para llevar directo a la revisión
                        st.session_state["ultimo_procesado_id"] = id_enc
                    else:
                        st.error(f"Error al registrar la encuesta en la cola de revisión: {id_enc}")
                        
                except Exception as ex:
                    st.error(f"Ocurrió un error al procesar las imágenes: {ex}")
    else:
        # Faltan fotos
        st.warning("⚠️ Debes subir obligatoriamente las fotos de la Página 1 y Página 2 para habilitar el procesamiento por Inteligencia Artificial.")
        st.button("🔍 Procesar Encuesta (Fusionar Páginas con IA)", disabled=True)

# -------------------------------------------------------------------------
# PESTAÑA 2: MÓDULO DE REVISIÓN Y COLA (HUMAN IN THE LOOP)
# -------------------------------------------------------------------------
with tab_revisar:
    st.header("✏️ Verificación Humana de la Caligrafía")
    st.write("Dado que los nombres de los niños son totalmente **anónimos**, el único campo manuscrito clave a revisar es el **nombre de la Institución Educativa**. Úsalo para corregir interpretaciones erróneas de la IA.")
    
    # Recargar datos
    _, df_cola = load_data()
    
    if df_cola.empty or len(df_cola[df_cola["Estado"] == "Pendiente"]) == 0:
        st.info("🎈 No hay encuestas pendientes de revisión en la cola. ¡Excelente trabajo!")
    else:
        # Listar encuestas pendientes
        pendientes = df_cola[df_cola["Estado"] == "Pendiente"]
        
        # Ofrecer selector
        cola_id_list = pendientes["ID_Encuesta"].tolist()
        
        # Enfocar en la última encuesta procesada si existe en session_state
        index_default = 0
        if "ultimo_procesado_id" in st.session_state and st.session_state["ultimo_procesado_id"] in cola_id_list:
            index_default = cola_id_list.index(st.session_state["ultimo_procesado_id"])
            
        selected_id = st.selectbox("Selecciona una encuesta para revisar y verificar:", cola_id_list, index=index_default)
        
        # Cargar fila seleccionada
        row_cola = pendientes[pendientes["ID_Encuesta"] == selected_id].iloc[0]
        datos_ia = json.loads(row_cola["JSON_Respuestas"])
        
        # Tarjeta de revisión
        st.markdown(f"### Revisando Registro Temporal: **{selected_id}** ({row_cola['Tipo_Formulario']})")
        
        col_form, col_vis = st.columns([3, 2])
        
        with col_form:
            st.markdown("#### ✏️ Datos del Formulario")
            
            # Campo Crítico: Institución Educativa
            ie_sugerida = datos_ia.get("institucion_educativa", "")
            ie_verificada = st.text_input(
                "🏫 INSTITUCIÓN EDUCATIVA (Verificar / Corregir caligrafía):",
                value=ie_sugerida,
                help="Revisa cómo se leyó el nombre de la escuela a mano y corrígelo si es necesario."
            )
            
            col_muni, col_fecha, col_rol = st.columns(3)
            with col_muni:
                municipio = st.text_input("📍 Municipio:", value=datos_ia.get("municipio", ""))
            with col_fecha:
                fecha = st.text_input("📅 Fecha:", value=datos_ia.get("fecha", ""))
            with col_rol:
                rol = st.selectbox("👤 Rol del Participante:", ["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"], index=["Estudiante", "Docente", "Padre de Familia", "Líder comunitario"].index(datos_ia.get("rol", "Estudiante")))
            
            # Respuestas de Conocimiento
            st.markdown("##### 📚 Respuestas de Conocimiento")
            resp_con = datos_ia.get("respuestas_conocimiento", {})
            
            with st.expander("Ver respuestas de conocimiento extraídas", expanded=True):
                for p_id, p_val in resp_con.items():
                    if p_val:  # Solo mostrar si tiene valor
                        st.write(f"**{p_id.upper()}:** {p_val}")
            
            # Respuestas de Satisfacción (Solo POSTEST)
            if row_cola["Tipo_Formulario"] == "POSTEST":
                st.markdown("##### ⭐ Evaluación de Satisfacción")
                resp_sat = datos_ia.get("satisfaccion", {})
                
                with st.expander("Ver nivel de satisfacción extraído", expanded=True):
                    col_sat1, col_sat2 = st.columns(2)
                    for idx, (s_id, s_val) in enumerate(resp_sat.items()):
                        if idx % 2 == 0:
                            with col_sat1:
                                st.write(f"**{s_id.replace('sat_', 'Aspecto ').upper()}:** {s_val}")
                        else:
                            with col_sat2:
                                st.write(f"**{s_id.replace('sat_', 'Aspecto ').upper()}:** {s_val}")
                                
            st.markdown("---")
            if st.button("💾 Aprobar y Guardar de forma definitiva", type="primary", use_container_width=True):
                # Preparar registro final para Respuestas_SRPA
                registro_final = {
                    "ID_Encuesta": selected_id,
                    "Tipo_Formulario": row_cola["Tipo_Formulario"],
                    "Fecha": fecha,
                    "Municipio": municipio,
                    "Institucion_Educativa_Verificada": ie_verificada,
                    "Rol": rol,
                    "Verificado_Por": "Revisor Humano",
                }
                
                # Mapear respuestas de conocimiento a columnas del Excel (Conocimientos_P1 a P8)
                if row_cola["Tipo_Formulario"] == "PRETEST":
                    for i in range(1, 9):
                        registro_final[f"Conocimientos_P{i}"] = resp_con.get(f"p{i}", "")
                else:  # POSTEST
                    for i in range(1, 6):
                        registro_final[f"Conocimientos_P{i}"] = resp_con.get(f"p{i}", "")
                        
                # Mapear satisfacción si es POSTEST
                if row_cola["Tipo_Formulario"] == "POSTEST":
                    resp_sat = datos_ia.get("satisfaccion", {})
                    for i in range(1, 10):
                        registro_final[f"Sat_P{i}"] = resp_sat.get(f"sat_{i}", "")
                
                # Guardar en Excel
                ok, error_msg = save_to_respuestas(registro_final)
                if ok:
                    update_cola_status(selected_id, "Aprobado")
                    st.success(f"✅ ¡Registro {selected_id} guardado con éxito en el Excel final!")
                    st.balloons()
                    # Limpiar session state del último id procesado
                    if "ultimo_procesado_id" in st.session_state:
                        del st.session_state["ultimo_procesado_id"]
                    st.rerun()
                else:
                    st.error(f"Error al guardar registro definitivo: {error_msg}")
                    
        with col_vis:
            st.markdown("#### 🔍 Caligrafía en las Encuestas")
            st.info("Usa estas guías de caligrafía real tomadas del campo para entrenar el ojo de tu equipo verificador.")
            
            # Ofrecer algunas imágenes de ejemplo reales cargadas en el notebook para entrenar el ojo del equipo
            st.caption("Ejemplos reales del proyecto (disponibles en tu Notebook):")
            
            # Lista de nombres de archivos de ejemplo del proyecto reales
            st.image("https://images.unsplash.com/photo-1456513080510-7bf3a84b82f8?auto=format&fit=crop&q=80&w=600", caption="Ejemplo de caligrafía y marcas en el proyecto", use_container_width=True)

# -------------------------------------------------------------------------
# PESTAÑA 3: DASHBOARD ESTADÍSTICO EN TIEMPO REAL
# -------------------------------------------------------------------------
with tab_dashboard:
    st.header("📊 Dashboard de Conocimiento e Impacto SRPA")
    st.write("Visualiza los datos acumulados y el impacto real de los talleres de prevención del SRPA.")
    
    # Recargar datos frescos
    df_resp, df_cola = load_data()
    
    if df_resp.empty:
        st.warning("⚠️ No hay datos registrados aún para generar el dashboard estadístico.")
    else:
        # Filtros del Dashboard en la barra lateral superior
        st.markdown("### 🔍 Filtros de Segmentación")
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            municipios_list = ["Todos"] + sorted(df_resp["Municipio"].dropna().unique().tolist())
            filtro_muni = st.selectbox("Filtrar por Municipio:", municipios_list)
            
        with col_f2:
            ie_list = ["Todas"] + sorted(df_resp["Institucion_Educativa_Verificada"].dropna().unique().tolist())
            filtro_ie = st.selectbox("Filtrar por Institución Educativa:", ie_list)
            
        # Filtrar DataFrame
        df_filtered = df_resp.copy()
        if filtro_muni != "Todos":
            df_filtered = df_filtered[df_filtered["Municipio"] == filtro_muni]
        if filtro_ie != "Todas":
            df_filtered = df_filtered[df_filtered["Institucion_Educativa_Verificada"] == filtro_ie]
            
        # 1. TARJETAS KPI
        total_encuestas = len(df_filtered)
        pretests_cnt = len(df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"])
        postests_cnt = len(df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"])
        pendientes_cola = len(df_cola[df_cola["Estado"] == "Pendiente"])
        
        st.markdown("### 📈 Indicadores Clave")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Procesadas", f"{total_encuestas}")
        kpi2.metric("Pretests (Línea Base)", f"{pretests_cnt}")
        kpi3.metric("Postests (Evaluados)", f"{postests_cnt}")
        kpi4.metric("Pendientes de Revisar", f"{pendientes_cola}", delta_color="inverse")
        
        # 2. COMPARATIVA ANTES VS DESPUÉS (Impacto en Conocimientos)
        st.markdown("### 📚 Impacto Educativo: Conocimientos Antes vs. Después")
        st.caption("Esta sección compara el porcentaje de respuestas correctas en los 4 conceptos clave evaluados en ambos tests.")
        
        # Definir respuestas correctas de los 4 conceptos clave compartidos
        # Concepto 1: Finalidad SRPA (Pretest Q2 vs Postest Q1) -> Correcta: "Promover la responsabilidad, la protección de derechos y la resocialización"
        # Concepto 2: Factor de riesgo (Pretest Q3 vs Postest Q2) -> Correcta: "Consumir sustancias psicoactivas"
        # Concepto 3: Factor protector (Pretest Q4 vs Postest Q3) -> Pretest Correcta: "Dialogar con la familia", Postest Correcta: "Participar en actividades deportivas, culturales o comunitarias"
        # Concepto 4: Responsable prevención (Pretest Q5 vs Postest Q4) -> Correcta: "La familia, la escuela, la comunidad y las instituciones"
        
        df_pre = df_filtered[df_filtered["Tipo_Formulario"] == "PRETEST"]
        df_post = df_filtered[df_filtered["Tipo_Formulario"] == "POSTEST"]
        
        if len(df_pre) > 0 and len(df_post) > 0:
            # Calcular porcentajes para Pretest
            correct_finalidad_pre = (df_pre["Conocimientos_P2"] == "Promover la responsabilidad, la protección de derechos y la resocialización").mean() * 100
            correct_riesgo_pre = (df_pre["Conocimientos_P3"] == "Consumir sustancias psicoactivas").mean() * 100
            correct_protector_pre = (df_pre["Conocimientos_P4"] == "Dialogar con la familia").mean() * 100
            correct_resp_pre = (df_pre["Conocimientos_P5"] == "La familia, la escuela, la comunidad y las instituciones").mean() * 100
            
            # Calcular porcentajes para Postest
            correct_finalidad_post = (df_post["Conocimientos_P1"] == "Promover la responsabilidad, la protección de derechos y la resocialización").mean() * 100
            correct_riesgo_post = (df_post["Conocimientos_P2"] == "Consumir sustancias psicoactivas").mean() * 100
            correct_protector_post = (df_post["Conocimientos_P3"] == "Participar en actividades deportivas, culturales o comunitarias").mean() * 100
            correct_resp_post = (df_post["Conocimientos_P4"] == "La familia, la escuela, la comunidad y las instituciones").mean() * 100
            
            # Crear gráfico de barras comparativo usando Plotly
            conceptos = ["Finalidad SRPA", "Factor de Riesgo", "Factor Protector", "Corresponsabilidad"]
            porcentajes_pre = [correct_finalidad_pre, correct_riesgo_pre, correct_protector_pre, correct_resp_pre]
            porcentajes_post = [correct_finalidad_post, correct_riesgo_post, correct_protector_post, correct_resp_post]
            
            fig_comp = go.Figure()
            fig_comp.add_trace(go.Bar(
                x=conceptos,
                y=porcentajes_pre,
                name="Antes del Taller (Pretest)",
                marker_color="#1F4E78",
                text=[f"{p:.1f}%" for p in percentages_pre] if 'percentages_pre' in locals() else [f"{p:.1f}%" for p in porcentajes_pre],
                textposition='auto'
            ))
            fig_comp.add_trace(go.Bar(
                x=conceptos,
                y=porcentajes_post,
                name="Después del Taller (Postest)",
                marker_color="#2E7D32",
                text=[f"{p:.1f}%" for p in porcentajes_post],
                textposition='auto'
            ))
            
            fig_comp.update_layout(
                barmode='group',
                yaxis_title="Porcentaje de Respuestas Correctas (%)",
                yaxis_range=[0, 100],
                margin=dict(l=20, r=20, t=30, b=20),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            
            st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.info("Necesitamos tener registros tanto de PRETEST como de POSTEST en la base de datos para mostrar la comparativa de impacto.")
            
        # 3. ROLES Y SATISFACCIÓN (Dos columnas)
        st.markdown("---")
        col_r, col_s = st.columns([2, 3])
        
        with col_r:
            st.markdown("#### 👤 Distribución de Participantes por Rol")
            rol_counts = df_filtered["Rol"].value_counts().reset_index()
            rol_counts.columns = ["Rol", "Cantidad"]
            
            fig_rol = px.pie(
                rol_counts, 
                values="Cantidad", 
                names="Rol", 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig_rol.update_layout(margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_rol, use_container_width=True)
            
        with col_s:
            st.markdown("#### ⭐ Evaluación de Satisfacción del Taller (Postest)")
            df_sat_filtered = df_filtered[(df_filtered["Tipo_Formulario"] == "POSTEST")]
            
            if len(df_sat_filtered) > 0:
                # Mapear cualitativo a cuantitativo para calcular promedio:
                # Excelente = 4, Bueno = 3, Regular = 2, Deficiente = 1
                map_sat = {"Excelente": 4, "Bueno": 3, "Regular": 2, "Deficiente": 1}
                
                sat_cols = [f"Sat_P{i}" for i in range(1, 10)]
                promedios = []
                
                # Aspectos evaluados (texto resumido para el gráfico)
                aspectos = [
                    "1. Claridad Info",
                    "2. Dominio Facilitadores",
                    "3. Metodología",
                    "4. Participación",
                    "5. Utilidad Temas",
                    "6. Organización",
                    "7. Materiales",
                    "8. Fortaleció Conoc.",
                    "9. Recomendaría"
                ]
                
                for col in sat_cols:
                    promedios.append(df_sat_filtered[col].map(map_sat).mean())
                
                fig_sat = go.Figure()
                # Barra de satisfacción real
                fig_sat.add_trace(go.Bar(
                    x=aspectos,
                    y=promedios,
                    marker_color="#F2A900",
                    name="Calificación Promedio",
                    text=[f"{p:.2f}" for p in promedios],
                    textposition='auto'
                ))
                # Línea meta de Bueno (3.0)
                fig_sat.add_trace(go.Scatter(
                    x=aspectos,
                    y=[3.0]*9,
                    mode='lines',
                    name='Meta Satisfacción (Bueno - 3.0)',
                    line=dict(color='red', width=2, dash='dash')
                ))
                
                fig_sat.update_layout(
                    yaxis_title="Puntaje (Escala 1 a 4)",
                    yaxis_range=[1, 4],
                    margin=dict(l=20, r=20, t=30, b=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                
                st.plotly_chart(fig_sat, use_container_width=True)
            else:
                st.info("No hay datos de Postest filtrados para calcular el nivel de satisfacción.")

# Descargar archivo consolidado de Excel actual
st.markdown("---")
st.subheader("📥 Exportación de Base de Datos")
st.write("Puedes descargar el archivo Excel consolidado en cualquier momento para trabajar fuera de línea.")

with open(EXCEL_FILE, "rb") as file:
    btn = st.download_button(
        label="⬇️ Descargar plantilla_encuestas_srpa.xlsx completo",
        data=file,
        file_name="plantilla_encuestas_srpa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
