import streamlit as st
import pandas as pd
import json
import time
import base64
from io import BytesIO
from PIL import Image

# Configurar página
st.set_page_config(page_title="Evaluación SRPA - Proyecto Construyendo Futuro", layout="wide")

# Inicializar cliente de Gemini de forma segura
def obtener_cliente_gemini():
    api_key = st.secrets.get("GEMINI_API_KEY", None)
    if not api_key:
        return None
    
    # Intentar importar el nuevo SDK de Google GenAI
    try:
        from google import genai
        return genai.Client(api_key=api_key)
    except ImportError:
        try:
            import google.generativeai as google_genai
            google_genai.configure(api_key=api_key)
            return google_genai
        except ImportError:
            return None

# Conexión exclusiva a Google Sheets
def conectar_google_sheets():
    try:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open("Base_Encuestas_SRPA")
    except Exception as e:
        return f"Error de conexión: {e}"

# CSS Personalizado
st.markdown("""
<style>
    .main-header { text-align: center; color: #1E3A8A; font-weight: bold; margin-bottom: 5px; }
    .sub-header { text-align: center; color: #4B5563; font-size: 1.1rem; margin-bottom: 25px; }
    .metric-card { background-color: #F3F4F6; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid #E5E7EB; }
</style>
""", unsafe_allow_html=True)

# Encabezado Principal
st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Sistematización de Encuestas de Conocimiento y Satisfacción - SRPA Gobernación de Bolívar</p>", unsafe_allow_html=True)

# Verificar credenciales
client_gs = conectar_google_sheets()
client_gemini = obtener_cliente_gemini()

if isinstance(client_gs, str) or not client_gemini:
    st.warning("⚠️ Configuración Pendiente en Streamlit Cloud Secrets.")
    with st.expander("🛠️ Guía de Configuración Paso a Paso"):
        st.write("Asegúrate de agregar `GEMINI_API_KEY` y `[gcp_service_account]` en los Secrets de Streamlit.")
    st.stop()

# Inicializar hojas en Google Sheets si no existen
try:
    try:
        ws_cola = client_gs.worksheet("Cola_Revision")
    except gspread.exceptions.WorksheetNotFound:
        ws_cola = client_gs.add_worksheet(title="Cola_Revision", rows="1000", cols="15")
        ws_cola.append_row(["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Foto_P1", "Foto_P2"])
    
    try:
        ws_resp = client_gs.worksheet("Respuestas_SRPA")
    except gspread.exceptions.WorksheetNotFound:
        ws_resp = client_gs.add_worksheet(title="Respuestas_SRPA", rows="2000", cols="30")
        headers = ["ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
                   "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
                   "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                   "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                   "Verificado_Por", "Fecha_Aprobacion"]
        ws_resp.append_row(headers)
except Exception as e:
    st.error(f"Error inicializando Google Sheets: {e}")
    st.stop()

# Helper para cargar datos de Sheets de forma segura (evita errores de columnas extra o vacías)
def cargar_hoja_segura(ws, columnas_esperadas):
    try:
        rows = ws.get_all_values()
        if not rows or len(rows) <= 1:
            return pd.DataFrame(columns=columnas_esperadas)
        
        headers = [h.strip() for h in rows[0]]
        data_rows = rows[1:]
        
        # Filtrar columnas vacías al final
        valid_indices = [i for i, h in enumerate(headers) if h != ""]
        clean_headers = [headers[i] for i in valid_indices]
        
        clean_data = []
        for r in data_rows:
            row_dict = {}
            for idx, col_idx in enumerate(valid_indices):
                val = r[col_idx] if col_idx < len(r) else ""
                row_dict[clean_headers[idx]] = val
            clean_data.append(row_dict)
            
        df = pd.DataFrame(clean_data)
        for col in columnas_esperadas:
            if col not in df.columns:
                df[col] = ""
        return df[columnas_esperadas]
    except Exception as e:
        st.error(f"Error cargando hoja: {e}")
        return pd.DataFrame(columns=columnas_esperadas)

# Definir columnas oficiales
COLS_COLA = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Foto_P1", "Foto_P2"]
COLS_RESP = ["ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
             "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4",
             "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
             "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
             "Verificado_Por", "Fecha_Aprobacion"]

# Cargar bases de datos
df_cola = cargar_hoja_segura(ws_cola, COLS_COLA)
df_resp = cargar_hoja_segura(ws_resp, COLS_RESP)

# Filtrar cola para pendientes reales
df_pendientes = df_cola[df_cola["Estado"].str.strip().str.upper() == "PENDIENTE"]

# Redimensionar y comprimir imágenes para evitar sobrecostos y límites de celdas
def optimizar_imagen(img_file):
    try:
        img = Image.open(img_file)
        img.thumbnail((350, 450))
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=75)
        return base64.b64encode(buffer.getvalue()).decode("utf-8")
    except Exception as e:
        st.error(f"Error comprimiendo imagen: {e}")
        return ""

# Motor de validación de doble vía en Google Sheets (Inmune a delay de replicación)
def registrar_en_respuestas_y_verificar(fila_datos, id_encuesta, ws_destino, ws_origen, index_original_cola):
    try:
        # 1. Escritura directa
        resp = ws_destino.append_row(fila_datos, value_input_option="USER_ENTERED")
        
        # 2. Verificación matemática mediante el rango de celdas devuelto por la API de Google
        if resp and "updates" in resp and resp["updates"].get("updatedRows", 0) > 0:
            # 3. Actualizar estado en la Cola de Revisión
            ws_origen.update_cell(index_original_cola + 2, COLS_COLA.index("Estado") + 1, "Aprobado")
            return True
    except Exception as e:
        st.error(f"Fallo en la conexión de Google Sheets: {e}")
    return False

def rechazar_de_cola_y_verificar(id_encuesta, ws_origen, index_original_cola):
    try:
        ws_origen.update_cell(index_original_cola + 2, COLS_COLA.index("Estado") + 1, "Rechazado")
        return True
    except Exception as e:
        st.error(f"Fallo al rechazar en Google Sheets: {e}")
    return False

# Pestañas
tab1, tab2, tab3 = st.tabs(["📥 Cargar Encuestas", "✏️ Banco de Verificación", "📊 Dashboard Estadístico"])

# --- TAB 1: CARGAR ENCUESTAS ---
with tab1:
    st.subheader("Captura y Digitalización de Encuestas Físicas")
    st.write("Sube las dos páginas de la encuesta física para que el motor de IA extraiga las respuestas de forma automática.")
    
    col1, col2 = st.columns(2)
    with col1:
        img_p1 = st.file_uploader("📸 Subir Página 1 (Encabezado y Pretest/Postest)", type=["jpg", "png", "jpeg"])
    with col2:
        img_p2 = st.file_uploader("📸 Subir Página 2 (Satisfacción/Continuación)", type=["jpg", "png", "jpeg"])
        
    tipo_form = st.selectbox("Tipo de Formulario Físico", ["PRETEST", "POSTEST"])
    
    if st.button("🚀 Procesar Encuesta con IA", use_container_width=True):
        if not img_p1 or not img_p2:
            st.error("Por favor, sube ambas páginas de la encuesta.")
        else:
            with st.spinner("Analizando caligrafía y marcas con Gemini..."):
                # Comprimir y pasar a base64
                b64_p1 = optimizar_imagen(img_p1)
                b64_p2 = optimizar_imagen(img_p2)
                
                # Prompt estructurado para extraer JSON
                prompt = f"""
                Analiza las imágenes de la encuesta del Proyecto "Construyendo Futuro" (SRPA Bolívar).
                Tipo de Formulario: {tipo_form}
                
                Instrucciones:
                1. Extrae los metadatos manuscritos: Fecha, Municipio, Institución Educativa.
                2. Identifica el Rol marcado con una X.
                3. Lee las respuestas de selección múltiple (X). Si una pregunta no tiene marcas, déjala en blanco "".
                
                Responde EXCLUSIVAMENTE con este formato JSON:
                {{
                    "fecha": "YYYY-MM-DD o vacío",
                    "municipio": "Texto detectado",
                    "institucion_educativa": "Texto manuscrito detectado",
                    "rol": "Estudiante/Docente/Padre de Familia/Lider comunitario",
                    "conocimientos": {{
                        "p1": "Respuesta o vacio",
                        "p2": "Respuesta o vacio",
                        "p3": "Respuesta o vacio",
                        "p4": "Respuesta o vacio",
                        "p5": "Respuesta o vacio",
                        "p6": "Respuesta o vacio",
                        "p7": "Respuesta o vacio",
                        "p8": "Respuesta o vacio"
                    }},
                    "satisfaccion": {{
                        "sat_p1": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p2": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p3": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p4": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p5": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p6": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p7": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p8": "Excelente/Bueno/Regular/Deficiente o vacio",
                        "sat_p9": "Excelente/Bueno/Regular/Deficiente o vacio"
                    }}
                }}
                """
                
                exito_ocr = False
                resultado_json = None
                
                # Carrusel de Reintentos de Modelos (Garantía de API)
                modelos = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
                for modelo in modelos:
                    try:
                        # Re-inicializar imágenes para el llamado
                        img_p1.seek(0)
                        pil_img1 = Image.open(img_p1)
                        img_p2.seek(0)
                        pil_img2 = Image.open(img_p2)
                        
                        if hasattr(client_gemini, "models"):  # SDK Nuevo
                            response = client_gemini.models.generate_content(
                                model=modelo,
                                contents=[pil_img1, pil_img2, prompt]
                            )
                            texto_respuesta = response.text
                        else:  # SDK Clásico
                            model = client_gemini.GenerativeModel(modelo)
                            response = model.generate_content([pil_img1, pil_img2, prompt])
                            texto_respuesta = response.text
                        
                        # Limpiar y parsear JSON
                        raw_text = texto_respuesta.strip()
                        if "```json" in raw_text:
                            raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                        elif "```" in raw_text:
                            raw_text = raw_text.split("```")[1].split("```")[0].strip()
                        
                        resultado_json = json.loads(raw_text)
                        exito_ocr = True
                        break
                    except Exception:
                        continue
                
                if not exito_ocr:
                    st.error("Error procesando con Gemini. Verifica tu clave de API.")
                else:
                    # Registrar en cola de revisión
                    nuevo_id = f"ENC_{time.strftime('%Y%m%d_%H%M%S')}"
                    fecha_carga = time.strftime('%Y-%m-%d %H:%M:%S')
                    
                    fila_cola = [
                        nuevo_id,
                        fecha_carga,
                        tipo_form,
                        resultado_json.get("municipio", ""),
                        resultado_json.get("institucion_educativa", ""),
                        resultado_json.get("rol", ""),
                        json.dumps(resultado_json),
                        "Pendiente",
                        b64_p1,
                        b64_p2
                    ]
                    
                    try:
                        ws_cola.append_row(fila_cola)
                        st.success("🎉 ¡Encuesta digitalizada y enviada a la Cola de Revisión con éxito!")
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error guardando en Cola de Revisión: {e}")

# --- TAB 2: BANCO DE VERIFICACIÓN ---
with tab2:
    st.subheader("Verificación Humana de Caligrafía")
    st.write("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
    
    if len(df_pendientes) == 0:
        st.info("No hay encuestas pendientes de verificación. ¡Buen trabajo!")
    else:
        # Seleccionar encuesta a revisar
        encuesta_sel = df_pendientes.iloc[0]
        id_encuesta = encuesta_sel["ID_Encuesta"]
        
        # Encontrar índice original en la hoja completa para poder actualizarla
        index_original_cola = df_cola[df_cola["ID_Encuesta"] == id_encuesta].index[0]
        
        try:
            datos_json = json.loads(encuesta_sel["JSON_Respuestas"])
        except:
            datos_json = {}
            
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.write(f"**ID Encuesta:** {id_encuesta}")
            st.write(f"**Tipo:** {encuesta_sel['Tipo_Formulario']}")
            
            # Mostrar imágenes lado a lado
            subcol1, subcol2 = st.columns(2)
            with subcol1:
                if encuesta_sel["Foto_P1"]:
                    try:
                        st.image(BytesIO(base64.b64decode(encuesta_sel["Foto_P1"])), caption="Página 1", use_container_width=True)
                    except:
                        st.warning("Imagen 1 no disponible.")
            with subcol2:
                if encuesta_sel["Foto_P2"]:
                    try:
                        st.image(BytesIO(base64.b64decode(encuesta_sel["Foto_P2"])), caption="Página 2", use_container_width=True)
                    except:
                        st.warning("Imagen 2 no disponible.")
                        
        with col2:
            st.markdown("### Datos Extraídos por la IA")
            
            # Campos editables por el verificador
            ie_verificada = st.text_input("🏫 Institución Educativa (Corregir caligrafía aquí):", encuesta_sel["Institucion_Educativa_IA"])
            municipio_verificado = st.text_input("📍 Municipio:", encuesta_sel["Municipio"])
            rol_verificado = st.text_input("👥 Rol:", encuesta_sel["Rol"])
            
            # Revisor humano
            revisor = st.text_input("👤 Nombre del Revisor:", "Coordinador Bolívar")
            
            # Estructurar respuestas de Conocimiento
            conocimientos = datos_json.get("conocimientos", {})
            satisfaccion = datos_json.get("satisfaccion", {})
            
            # Acciones
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True, type="primary"):
                    # Preparar fila final
                    fila_final = [
                        id_encuesta,
                        encuesta_sel["Tipo_Formulario"],
                        datos_json.get("fecha", ""),
                        municipio_verificado,
                        ie_verificada,
                        rol_verificado,
                        conocimientos.get("p1", ""), conocimientos.get("p2", ""), conocimientos.get("p3", ""), conocimientos.get("p4", ""),
                        conocimientos.get("p5", ""), conocimientos.get("p6", ""), conocimientos.get("p7", ""), conocimientos.get("p8", ""),
                        satisfaccion.get("sat_p1", ""), satisfaccion.get("sat_p2", ""), satisfaccion.get("sat_p3", ""), satisfaccion.get("sat_p4", ""),
                        satisfaccion.get("sat_p5", ""), satisfaccion.get("sat_p6", ""), satisfaccion.get("sat_p7", ""), satisfaccion.get("sat_p8", ""), satisfaccion.get("sat_p9", ""),
                        revisor,
                        time.strftime('%Y-%m-%d %H:%M:%S')
                    ]
                    
                    # Guardar con validación estricta de doble vía
                    if registrar_en_respuestas_y_verificar(fila_final, id_encuesta, ws_resp, ws_cola, index_original_cola):
                        st.success("¡Encuesta verificada e ingresada correctamente!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("Error al registrar la encuesta en Google Sheets. Por favor, reintenta.")
                        
            with col_b2:
                if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                    if rechazar_de_cola_y_verificar(id_encuesta, ws_cola, index_original_cola):
                        st.warning("La encuesta ha sido descartada.")
                        time.sleep(1)
                        st.rerun()

# --- TAB 3: DASHBOARD ESTADÍSTICO ---
with tab3:
    st.subheader("Indicadores de Conocimiento y Satisfacción Real")
    
    if len(df_resp) == 0:
        st.info("Sube y verifica encuestas para ver el Dashboard interactivo.")
    else:
        # Métricas generales
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"<div class='metric-card'><h3>{len(df_resp)}</h3><p>Total Encuestas Procesadas</p></div>", unsafe_allow_html=True)
        with m2:
            ie_counts = df_resp["Institucion_Educativa_Verificada"].nunique()
            st.markdown(f"<div class='metric-card'><h3>{ie_counts}</h3><p>Instituciones Impactadas</p></div>", unsafe_allow_html=True)
        with m3:
            mun_counts = df_resp["Municipio"].nunique()
            st.markdown(f"<div class='metric-card'><h3>{mun_counts}</h3><p>Municipios de Bolívar</p></div>", unsafe_allow_html=True)
            
        st.markdown("---")
        
        # Gráfico por Municipios
        import plotly.express as px
        df_mun = df_resp["Municipio"].value_counts().reset_index()
        df_mun.columns = ["Municipio", "Cantidad"]
        fig_mun = px.bar(df_mun, x="Municipio", y="Cantidad", title="Participantes por Municipio", color="Municipio")
        st.plotly_chart(fig_mun, use_container_width=True)
