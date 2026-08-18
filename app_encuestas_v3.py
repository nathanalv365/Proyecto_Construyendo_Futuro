import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
import io
import base64
import time
from datetime import datetime
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Intento de importación dinámica de Google GenAI / GenerativeAI
use_new_sdk = False
try:
    from google import genai
    from google.genai import types
    use_new_sdk = True
except ImportError:
    try:
        import google.generativeai as genai_old
        use_new_sdk = False
    except ImportError:
        pass

# Configuración de página
st.set_page_config(page_title="Proyecto Construyendo Futuro", page_icon="📊", layout="wide")

# CSS personalizado responsivo para móviles y escritorio con unsafe_allow_html=True
st.markdown("""
<style>
.main-header {
    color: #1E3A8A;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-weight: 700;
    text-align: center;
    margin-bottom: 2px;
}
.sub-header {
    color: #4B5563;
    font-family: 'Segoe UI', sans-serif;
    text-align: center;
    font-size: 14px;
    margin-bottom: 25px;
}
.card {
    background-color: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    margin-bottom: 20px;
}
.metric-box {
    background-color: #F3F4F6;
    border-radius: 8px;
    padding: 15px;
    text-align: center;
    border-left: 5px solid #3B82F6;
}
.metric-val {
    font-size: 24px;
    font-weight: bold;
    color: #1E3A8A;
}
.metric-lbl {
    font-size: 12px;
    color: #6B7280;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)

# Helper para redimensionar y comprimir imágenes evitando el error de 'Resampling' y 'LANCELET'
def comprimir_imagen(image_file, max_width=350):
    try:
        img = Image.open(image_file)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        width, height = img.size
        if width > max_width:
            ratio = max_width / float(width)
            new_height = int(float(height) * float(ratio))
            
            # Intento dinámico de obtener el método de remuestreo (resampling) correcto
            resample_method = Image.NEAREST
            try:
                # Pillow >= 9.1.0 usa Image.Resampling
                resample_method = Image.Resampling.LANCZOS
            except AttributeError:
                try:
                    # Versiones más antiguas usan Image.LANCZOS
                    resample_method = Image.LANCZOS
                except AttributeError:
                    # Fallback definitivo a Image.ANTIALIAS
                    try:
                        resample_method = Image.ANTIALIAS
                    except AttributeError:
                        pass
            
            img = img.resize((max_width, new_height), resample_method)
            
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=75)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Error comprimiendo imagen: {e}")
        return None

# Convertir bytes de imagen comprimida a base64 para guardado en Sheets
def bytes_a_b64(img_bytes):
    if img_bytes:
        return base64.b64encode(img_bytes).decode("utf-8")
    return ""

def b64_a_bytes(b64_str):
    if b64_str:
        return base64.b64decode(b64_str.encode("utf-8"))
    return None

# Conexión con Google Sheets
def obtener_credenciales_gcp():
    if "gcp_service_account" in st.secrets:
        return dict(st.secrets["gcp_service_account"])
    return None

def conectar_google_sheets():
    creds_dict = obtener_credenciales_gcp()
    if not creds_dict:
        return None
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        # Corregir saltos de línea de la clave privada
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets API: {e}")
        return None

def inicializar_base_datos(client):
    try:
        nombre_doc = "Base_Encuestas_SRPA"
        try:
            doc = client.open(nombre_doc)
        except gspread.SpreadsheetNotFound:
            doc = client.create(nombre_doc)
            st.info(f"Se ha creado un nuevo libro de Google Sheets llamado '{nombre_doc}' en tu cuenta.")
        
        # Inicializar Cola_Revision
        try:
            ws_cola = doc.worksheet("Cola_Revision")
        except gspread.WorksheetNotFound:
            ws_cola = doc.add_worksheet("Cola_Revision", rows=1000, cols=12)
            headers = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Imagen_B64_P1", "Imagen_B64_P2"]
            ws_cola.append_row(headers)
        
        # Inicializar Respuestas_SRPA
        try:
            ws_resp = doc.worksheet("Respuestas_SRPA")
        except gspread.WorksheetNotFound:
            ws_resp = doc.add_worksheet("Respuestas_SRPA", rows=1000, cols=30)
            headers = [
                "ID_Encuesta", "Tipo_Formulario", "Fecha", "Municipio", "Institucion_Educativa_Verificada", "Rol",
                "Conocimientos_P1", "Conocimientos_P2", "Conocimientos_P3", "Conocimientos_P4", 
                "Conocimientos_P5", "Conocimientos_P6", "Conocimientos_P7", "Conocimientos_P8",
                "Sat_P1", "Sat_P2", "Sat_P3", "Sat_P4", "Sat_P5", "Sat_P6", "Sat_P7", "Sat_P8", "Sat_P9",
                "Verificado_Por", "Fecha_Aprobacion"
            ]
            ws_resp.append_row(headers)
            
        return doc
    except Exception as e:
        st.error(f"Error inicializando las hojas de cálculo: {e}")
        return None

def cargar_cola_revision(doc):
    try:
        ws = doc.worksheet("Cola_Revision")
        rows = ws.get_all_values()
        if len(rows) <= 1:
            return pd.DataFrame(columns=["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Imagen_B64_P1", "Imagen_B64_P2"])
        headers = rows[0]
        data = rows[1:]
        df = pd.DataFrame(data, columns=headers)
        columnas_criticas = ["ID_Encuesta", "Fecha_Carga", "Tipo_Formulario", "Municipio", "Institucion_Educativa_IA", "Rol", "JSON_Respuestas", "Estado", "Imagen_B64_P1", "Imagen_B64_P2"]
        for col in columnas_criticas:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e:
        st.error(f"Error cargando Cola de Revisión: {e}")
        return pd.DataFrame()

def cargar_respuestas_validadas(doc):
    try:
        ws = doc.worksheet("Respuestas_SRPA")
        rows = ws.get_all_values()
        if len(rows) <= 1:
            return pd.DataFrame()
        headers = rows[0]
        data = rows[1:]
        df = pd.DataFrame(data, columns=headers)
        return df
    except Exception as e:
        st.error(f"Error cargando respuestas validadas: {e}")
        return pd.DataFrame()

# Escritura y validación de doble vía en Sheets
def append_row_and_verify(ws, row_data, id_to_check):
    try:
        ws.append_row(row_data)
        time.sleep(1.5)
        cells = ws.findall(id_to_check)
        if cells:
            return True
        return False
    except Exception as e:
        st.error(f"Error de doble vía al escribir registro: {e}")
        return False

def update_row_and_verify(ws, row_index, row_data):
    try:
        num_cols = len(row_data)
        rango_celdas = f"A{row_index}:{gspread.utils.rowcol_to_a1(row_index, num_cols)}"
        ws.update(rango_celdas, [row_data])
        time.sleep(1.5)
        val = ws.cell(row_index, 1).value
        if val == row_data[0]:
            return True
        return False
    except Exception as e:
        st.error(f"Error de doble vía al actualizar registro: {e}")
        return False

# Motor OCR con resiliencia de carrusel de modelos de Gemini
def consultar_gemini_vision(p1_bytes, p2_bytes, content_type):
    prompt_ocr = """
    Analiza estas dos páginas de una encuesta del "Proyecto Construyendo Futuro" (SRPA).
    Identifica el tipo de formulario ("PRETEST" o "POSTEST").
    Extrae con extrema precisión:
    - Fecha (en la parte superior de la página 1)
    - Municipio (en la parte superior de la página 1)
    - Institución Educativa (en la parte superior de la página 1). NOTA: Si está escrito a mano y es inteligible, extráelo textualmente.
    - Rol del participante (marcado con X en Estudiante, Docente, Padre de Familia, Lider comunitario).
    - Respuestas de conocimiento (1 a 8 marcadas con X en el Pretest, o Sección A 1 a 5 en el Postest):
      Para cada pregunta, extrae el literal marcado (ej. "a", "b", "c", "d" o vacío "" si no está marcado). NO asumas ni auto-rellenes ninguna respuesta si la casilla está vacía.
    - Evaluación de satisfacción (Solo en el Postest, Sección B preguntas 1 a 9):
      Extrae la opción marcada (Excelente, Bueno, Regular, Deficiente). Si está vacía, devuelve "". NO auto-rellenes.
    
    Responde estrictamente en formato JSON válido con la siguiente estructura:
    {
      "tipo_formulario": "PRETEST",
      "encabezado": {
        "fecha": "30/07/2026",
        "municipio": "Cartagena",
        "institucion_educativa": "Promesa de Dios",
        "rol": "Estudiante"
      },
      "respuestas_conocimiento": {
        "p1": "b",
        "p2": "b",
        "p3": "b",
        "p4": "a",
        "p5": "b",
        "p6": "a",
        "p7": "c",
        "p8": "c"
      },
      "evaluacion_satisfaccion": {
        "sat_p1": "",
        "sat_p2": "",
        "sat_p3": "",
        "sat_p4": "",
        "sat_p5": "",
        "sat_p6": "",
        "sat_p7": "",
        "sat_p8": "",
        "sat_p9": ""
      }
    }
    """
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Clave API de Gemini (GEMINI_API_KEY) no configurada en Secrets.")

    modelos_carrusel = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

    # Intento 1: Nuevo SDK (genai)
    if use_new_sdk:
        for model_name in modelos_carrusel:
            try:
                client = genai.Client(api_key=api_key)
                parts = [
                    types.Part.from_bytes(data=p1_bytes, mime_type=content_type),
                    types.Part.from_bytes(data=p2_bytes, mime_type=content_type),
                    prompt_ocr
                ]
                response = client.models.generate_content(
                    model=model_name,
                    contents=parts,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                return json.loads(response.text)
            except Exception as e:
                continue

    # Intento 2: Fallback SDK Antiguo (google.generativeai)
    else:
        for model_name in modelos_carrusel:
            try:
                genai_old.configure(api_key=api_key)
                img1 = Image.open(io.BytesIO(p1_bytes))
                img2 = Image.open(io.BytesIO(p2_bytes))
                model = genai_old.GenerativeModel(model_name)
                response = model.generate_content([
                    prompt_ocr,
                    img1,
                    img2
                ])
                text_cleaned = response.text.strip()
                if text_cleaned.startswith("```json"):
                    text_cleaned = text_cleaned.split("```json")[1].split("```")[0].strip()
                elif text_cleaned.startswith("```"):
                    text_cleaned = text_cleaned.split("```")[1].split("```")[0].strip()
                return json.loads(text_cleaned)
            except Exception as e:
                continue

    raise RuntimeError("Ninguno de los modelos de Gemini pudo procesar la solicitud.")

# Flujo de Interfaz
st.markdown("<h1 class='main-header'>Proyecto \"Construyendo Futuro\"</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Evaluación de Conocimientos SRPA - Gobernación de Bolívar</p>", unsafe_allow_html=True)

client = conectar_google_sheets()

if not client:
    st.warning("⚠️ Conexión a Google Sheets no configurada.")
    st.markdown("""
    ### 💻 Cómo conectar tu Google Sheet:
    1. Genera tu archivo JSON de cuenta de servicio en **Google Cloud Console**.
    2. Comparte tu documento de Google Sheet llamado **`Base_Encuestas_SRPA`** con el correo de la cuenta de servicio como **Editor**.
    3. Pega tus credenciales en el archivo de secretos de Streamlit Cloud (`Secrets`) con el siguiente formato:
    
    ```toml
    GEMINI_API_KEY = "tu_api_key_aquí"
    
    [gcp_service_account]
    type = "service_account"
    project_id = "tu-proyecto-id"
    private_key_id = "tu-private-key-id"
    private_key = "-----BEGIN PRIVATE KEY-----\\nTU_LLAVE_PRIVADA\\n-----END PRIVATE KEY-----\\n"
    client_email = "tu-cuenta-de-servicio@proyecto.iam.gserviceaccount.com"
    client_id = "tu-client-id"
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.google.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/tu-cuenta-de-servicio"
    ```
    """)
else:
    doc = inicializar_base_datos(client)
    if doc:
        pestana1, pestana2, pestana3 = st.tabs(["📤 Cargar Encuestas", "✅ Banco de Verificación", "📊 Dashboard de Impacto"])
        
        # PESTAÑA 1: Carga Móvil
        with pestana1:
            st.subheader("Carga y Procesamiento de Encuestas Físicas")
            st.markdown("Carga las fotos de la **Página 1** y la **Página 2** de una encuesta para digitalizarla de forma automática.")
            
            p1_file = st.file_uploader("Subir Página 1 (Encabezado y Preguntas 1-3)", type=["jpg", "jpeg", "png"], key="p1")
            p2_file = st.file_uploader("Subir Página 2 (Preguntas 4-8 o Satisfacción)", type=["jpg", "jpeg", "png"], key="p2")
            
            if p1_file and p2_file:
                if st.button("🔍 Procesar Encuesta con IA", use_container_width=True):
                    with st.spinner("Leyendo caligrafía y marcas con Gemini..."):
                        p1_comp = comprimir_imagen(p1_file)
                        p2_comp = comprimir_imagen(p2_file)
                        
                        if p1_comp and p2_comp:
                            try:
                                json_res = consultar_gemini_vision(p1_comp, p2_comp, p1_file.type)
                                
                                id_encuesta = f"ENC_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                                fecha_carga = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                tipo_form = json_res.get("tipo_formulario", "PRETEST")
                                municipio = json_res.get("encabezado", {}).get("municipio", "")
                                ie_ia = json_res.get("encabezado", {}).get("institucion_educativa", "")
                                rol = json_res.get("encabezado", {}).get("rol", "")
                                
                                img_b64_p1 = bytes_a_b64(p1_comp)
                                img_b64_p2 = bytes_a_b64(p2_comp)
                                
                                # Insertar en la Cola_Revision y verificar
                                ws_cola = doc.worksheet("Cola_Revision")
                                headers = ws_cola.row_values(1)
                                
                                row_map = {
                                    "ID_Encuesta": id_encuesta,
                                    "Fecha_Carga": fecha_carga,
                                    "Tipo_Formulario": tipo_form,
                                    "Municipio": municipio,
                                    "Institucion_Educativa_IA": ie_ia,
                                    "Rol": rol,
                                    "JSON_Respuestas": json.dumps(json_res),
                                    "Estado": "Pendiente",
                                    "Imagen_B64_P1": img_b64_p1,
                                    "Imagen_B64_P2": img_b64_p2
                                }
                                
                                new_row = [row_map.get(h, "") for h in headers]
                                
                                if append_row_and_verify(ws_cola, new_row, id_encuesta):
                                    st.success(f"🎉 ¡Encuesta cargada con éxito! ID: {id_encuesta}. Por favor, ve al **Banco de Verificación** para confirmar los datos manuscritos.")
                                    time.sleep(2)
                                    st.rerun()
                                else:
                                    st.error("Error al registrar la encuesta en Google Sheets. Por favor, reintenta.")
                                    
                            except Exception as e:
                                st.error(f"Error procesando la encuesta con Gemini (Error en API): {e}")
            
        # PESTAÑA 2: Banco de Verificación Lado a Lado
        with pestana2:
            st.subheader("Banco de Verificación de Escritura a Mano")
            st.markdown("Verifica y corrige el nombre de la Institución Educativa manuscrita antes de guardarla permanentemente.")
            
            df_cola = cargar_cola_revision(doc)
            
            if not df_cola.empty:
                df_filtrado = df_cola[df_cola["Estado"] == "Pendiente"]
                
                if not df_filtrado.empty:
                    registro = df_filtrado.iloc[0]
                    idx_original = df_cola[df_cola["ID_Encuesta"] == registro["ID_Encuesta"]].index[0] + 2 # +1 encabezado +1 base 1
                    
                    st.warning(f"Revisando Registro {registro['ID_Encuesta']} - Cargado el {registro['Fecha_Carga']}")
                    
                    # Cargar imágenes guardadas
                    p1_bytes = b64_a_bytes(registro["Imagen_B64_P1"])
                    p2_bytes = b64_a_bytes(registro["Imagen_B64_P2"])
                    
                    col_img1, col_img2 = st.columns(2)
                    with col_img1:
                        if p1_bytes:
                            st.image(p1_bytes, caption="Página 1 (Encabezado manuscrito)", use_container_width=True)
                        else:
                            st.info("Imagen de la página 1 no disponible.")
                    with col_img2:
                        if p2_bytes:
                            st.image(p2_bytes, caption="Página 2 (Respuestas y satisfacción)", use_container_width=True)
                        else:
                            st.info("Imagen de la página 2 no disponible.")
                    
                    # Formulario de verificación
                    st.markdown("### 📝 Corregir Datos de Encabezado")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        tipo_verif = st.selectbox("Tipo de Formulario", ["PRETEST", "POSTEST"], index=0 if registro["Tipo_Formulario"] == "PRETEST" else 1)
                    with col2:
                        mun_verif = st.text_input("Municipio", value=registro["Municipio"])
                    with col3:
                        ie_verif = st.text_input("Institución Educativa (MANUSCRITA)", value=registro["Institucion_Educativa_IA"])
                    with col4:
                        rol_verif = st.text_input("Rol del Participante", value=registro["Rol"])
                        
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("✅ Aprobar e Ingresar a Base de Datos", use_container_width=True):
                            with st.spinner("Guardando en la base de datos definitiva..."):
                                try:
                                    raw_json = json.loads(registro["JSON_Respuestas"])
                                    
                                    # Preparar fila de Respuestas_SRPA
                                    ws_resp = doc.worksheet("Respuestas_SRPA")
                                    headers_resp = ws_resp.row_values(1)
                                    
                                    # Mapear respuestas
                                    con_p = raw_json.get("respuestas_conocimiento", {})
                                    sat_p = raw_json.get("evaluacion_satisfaccion", {})
                                    
                                    resp_map = {
                                        "ID_Encuesta": registro["ID_Encuesta"],
                                        "Tipo_Formulario": tipo_verif,
                                        "Fecha": raw_json.get("encabezado", {}).get("fecha", ""),
                                        "Municipio": mun_verif,
                                        "Institucion_Educativa_Verificada": ie_verif,
                                        "Rol": rol_verif,
                                        "Conocimientos_P1": con_p.get("p1", ""),
                                        "Conocimientos_P2": con_p.get("p2", ""),
                                        "Conocimientos_P3": con_p.get("p3", ""),
                                        "Conocimientos_P4": con_p.get("p4", ""),
                                        "Conocimientos_P5": con_p.get("p5", ""),
                                        "Conocimientos_P6": con_p.get("p6", ""),
                                        "Conocimientos_P7": con_p.get("p7", ""),
                                        "Conocimientos_P8": con_p.get("p8", ""),
                                        "Sat_P1": sat_p.get("sat_p1", ""),
                                        "Sat_P2": sat_p.get("sat_p2", ""),
                                        "Sat_P3": sat_p.get("sat_p3", ""),
                                        "Sat_P4": sat_p.get("sat_p4", ""),
                                        "Sat_P5": sat_p.get("sat_p5", ""),
                                        "Sat_P6": sat_p.get("sat_p6", ""),
                                        "Sat_P7": sat_p.get("sat_p7", ""),
                                        "Sat_P8": sat_p.get("sat_p8", ""),
                                        "Sat_P9": sat_p.get("sat_p9", ""),
                                        "Verificado_Por": "Coordinador de Campo",
                                        "Fecha_Aprobacion": datetime.now().strftime('%Y-%m-%d')
                                    }
                                    
                                    row_resp_final = [resp_map.get(h, "") for h in headers_resp]
                                    
                                    # Insertar y verificar
                                    if append_row_and_verify(ws_resp, row_resp_final, registro["ID_Encuesta"]):
                                        # Actualizar el registro en Cola_Revision y verificar
                                        ws_cola = doc.worksheet("Cola_Revision")
                                        headers_cola = ws_cola.row_values(1)
                                        
                                        cola_row_map = dict(registro)
                                        cola_row_map["Estado"] = "Aprobado"
                                        cola_row_map["Institucion_Educativa_IA"] = ie_verif
                                        cola_row_map["Municipio"] = mun_verif
                                        cola_row_map["Tipo_Formulario"] = tipo_verif
                                        cola_row_map["Rol"] = rol_verif
                                        
                                        row_cola_final = [cola_row_map.get(h, "") for h in headers_cola]
                                        
                                        if update_row_and_verify(ws_cola, idx_original, row_cola_final):
                                            st.success("🎉 ¡Registro verificado e insertado correctamente!")
                                            time.sleep(1.5)
                                            st.rerun()
                                        else:
                                            st.error("Error al actualizar la Cola de Revisión en Google Sheets.")
                                    else:
                                        st.error("Error al insertar el registro en Respuestas_SRPA en Google Sheets.")
                                except Exception as e:
                                    st.error(f"Error al procesar y guardar: {e}")
                                    
                    with col_b2:
                        if st.button("❌ Rechazar Entrada (Eliminar de la Cola)", use_container_width=True):
                            with st.spinner("Rechazando y eliminando registro..."):
                                try:
                                    ws_cola = doc.worksheet("Cola_Revision")
                                    headers_cola = ws_cola.row_values(1)
                                    
                                    cola_row_map = dict(registro)
                                    cola_row_map["Estado"] = "Rechazado"
                                    
                                    row_cola_final = [cola_row_map.get(h, "") for h in headers_cola]
                                    
                                    if update_row_and_verify(ws_cola, idx_original, row_cola_final):
                                        st.warning("⚠️ Entrada rechazada y eliminada de la cola con éxito.")
                                        time.sleep(1.5)
                                        st.rerun()
                                    else:
                                        st.error("Error al actualizar el estado de rechazo en Google Sheets.")
                                except Exception as e:
                                    st.error(f"Error al rechazar registro: {e}")
                else:
                    st.info("🙌 No hay encuestas pendientes de verificación en este momento. ¡Buen trabajo!")
            else:
                st.info("🙌 La base de datos temporal está limpia. Registra encuestas en la pestaña de Carga para empezar.")
                
        # PESTAÑA 3: Dashboard de Impacto
        with pestana3:
            st.subheader("Dashboard Estadístico del Proyecto")
            df_resp = cargar_respuestas_validadas(doc)
            
            if not df_resp.empty:
                # Métricas generales en cards responsive
                num_encuestas = len(df_resp)
                colegios = df_resp["Institucion_Educativa_Verificada"].nunique()
                municipios = df_resp["Municipio"].nunique()
                
                m_col1, m_col2, m_col3 = st.columns(3)
                with m_col1:
                    st.markdown(f"<div class='metric-box'><div class='metric-val'>{num_encuestas}</div><div class='metric-lbl'>Encuestas Validadas</div></div>", unsafe_allow_html=True)
                with m_col2:
                    st.markdown(f"<div class='metric-box'><div class='metric-val'>{colegios}</div><div class='metric-lbl'>Colegios Participantes</div></div>", unsafe_allow_html=True)
                with m_col3:
                    st.markdown(f"<div class='metric-box'><div class='metric-val'>{municipios}</div><div class='metric-lbl'>Municipios Impactados</div></div>", unsafe_allow_html=True)
                
                st.write("")
                
                # Gráfico de tipo de formulario
                col_chart1, col_chart2 = st.columns(2)
                with col_chart1:
                    fig_tipo = px.pie(df_resp, names="Tipo_Formulario", title="Distribución de Participación (Pretest vs Postest)", color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig_tipo, use_container_width=True)
                with col_chart2:
                    fig_rol = px.bar(df_resp, x="Rol", title="Participantes por Rol", color="Rol", color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig_rol, use_container_width=True)
                
                # Gráfico de Satisfacción si es Postest
                st.markdown("### 🌟 Evaluación de Satisfacción de las Jornadas (Postest)")
                df_post = df_resp[df_resp["Tipo_Formulario"] == "POSTEST"]
                if not df_post.empty:
                    satisfaccion_cols = [c for c in df_post.columns if c.startswith("Sat_P")]
                    if satisfaccion_cols:
                        st.markdown("Resultados agregados para el aspecto: **La jornada fortaleció mis conocimientos sobre la prevención del SRPA (Sat_P8)**")
                        sat_counts = df_post["Sat_P8"].value_counts().reset_index()
                        sat_counts.columns = ["Calificación", "Cantidad"]
                        
                        fig_sat = px.bar(sat_counts, x="Calificación", y="Cantidad", color="Calificación", title="Fortalecimiento de Conocimiento (Pregunta 8)", color_discrete_sequence=px.colors.qualitative.Safe)
                        st.plotly_chart(fig_sat, use_container_width=True)
                    else:
                        st.info("No se han cargado respuestas de satisfacción en el Postest.")
                else:
                    st.info("Se requiere digitalizar formularios tipo POSTEST para visualizar la evaluación de satisfacción.")
            else:
                st.info("No hay datos reales almacenados en la base de datos de Google Sheets. Sube y valida encuestas en las pestañas anteriores para habilitar las estadísticas.")
