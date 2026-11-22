import base64
import difflib
import os
import re
import zipfile
from io import BytesIO
from typing import List, Literal

import openpyxl
import pandas as pd
import pymupdf
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image
from pydantic import BaseModel, Field

# ---------------------------------------------------------
# 1. CONFIGURACIÓN INICIAL Y ESTILOS
# ---------------------------------------------------------
load_dotenv()

st.set_page_config(
    page_title="Calificador Automático de Hojas de Trabajo",
    page_icon="📝",
    layout="wide",
)

MODELOS_RESPALDO = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
]

TIPOS_EJERCICIO = ["elegir_palabra", "escribir_palabra", "encerrar_opcion"]
UBICACIONES = ["Izquierda", "Derecha", "Completa"]


def obtener_cliente_genai():
    api_key = None
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
    except Exception:
        pass
    if not api_key:
        api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.error(
            "No se encontró GEMINI_API_KEY. Configúrala en Secrets de Streamlit Cloud o en tu archivo .env."
        )
        st.stop()
    return genai.Client(api_key=api_key)


def cargar_fondo_base64(ruta_imagen: str) -> str:
    if not os.path.exists(ruta_imagen):
        return ""
    with open(ruta_imagen, "rb") as f:
        return base64.b64encode(f.read()).decode()


def aplicar_estilos():
    fondo_b64 = cargar_fondo_base64("fondo.jpg")
    bg_css = (
        f"""
        .stApp {{
            background-image: url("data:image/jpeg;base64,{fondo_b64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        """
        if fondo_b64
        else ""
    )

    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');

        {bg_css}

        html, body, [class*="css"], .stMarkdown, .stTextInput, .stSelectbox, .stNumberInput, button {{
            font-family: 'DM Sans', sans-serif !important;
        }}

        h1, h2, h3, h4, h5, h6 {{
            font-family: 'DM Sans', sans-serif !important;
            font-weight: 700 !important;
        }}

        p, label, .stMarkdown p {{
            font-size: calc(1rem + 2px) !important;
            font-weight: 500 !important;
        }}

        /* Botones primarios negros */
        div.stButton > button[kind="primary"],
        div.stDownloadButton > button[kind="primary"],
        div.stButton > button {{
            background-color: #111111 !important;
            color: #FFFFFF !important;
            border: 1px solid #111111 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
        }}

        div.stButton > button:hover,
        div.stDownloadButton > button:hover {{
            background-color: #333333 !important;
            border-color: #333333 !important;
            color: #FFFFFF !important;
        }}

        /* Tarjetas con fondo blanco semitransparente */
        .bloque-card {{
            background-color: rgba(255, 255, 255, 0.94);
            padding: 1.2rem;
            border-radius: 12px;
            margin-bottom: 1rem;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


aplicar_estilos()


# ---------------------------------------------------------
# 2. MODELOS PYDANTIC Y PROMPTS
# ---------------------------------------------------------
class BloqueClaveIA(BaseModel):
    nombre: str = Field(description="Nombre del bloque, ej: Ejercicio 1")
    tipo: Literal["elegir_palabra", "escribir_palabra", "encerrar_opcion"]
    inicia_en: int = Field(default=1, description="Número del primer ítem real del ejercicio")
    ubicacion: Literal["Izquierda", "Derecha", "Completa"] = "Completa"
    respuestas: List[str] = Field(
        description="Lista ordenada de respuestas correctas. Si hay alternativas válidas, separarlas con /"
    )


class ClaveCompletaIA(BaseModel):
    bloques: List[BloqueClaveIA]


class RespuestaItem(BaseModel):
    id_item: str = Field(description="Identificador exacto del ítem, ej: Ejercicio 1 - ítem 1")
    respuesta_estudiante: str = Field(
        description="Texto exacto transcrito o código de anomalía (EN_BLANCO, DOBLE_MARCA, ILEGIBLE)"
    )


class HojaEstudianteIA(BaseModel):
    nombre_estudiante: str = Field(
        description="Nombre del estudiante escrito en la hoja, o exactamente 'SIN NOMBRE' si no aparece"
    )
    respuestas: List[RespuestaItem]
    items_con_tachon: List[str] = Field(
        default_factory=list,
        description="Lista de id_item donde el estudiante tachó una respuesta antes de escribir/marcar otra",
    )


PROMPT_ANTI_SESGO = """
Eres un escáner OCR puramente visual y estricto. NO eres profesor.
ADVERTENCIA CRÍTICA: Esta hoja contiene errores, tachones y marcas imprecisas hechas por estudiantes.
- PROHIBIDO deducir la respuesta correcta leyendo el texto de la hoja.
- Fíjate ÚNICAMENTE en la forma visual de las letras escritas a mano y en qué texto exacto está dentro del círculo o sobre el subrayado.
- REGLA DE MÚLTIPLES RENGLONES: Solo combina dos líneas de texto si ambas pertenecen claramente al espacio del MISMO ítem (sin cruzar hacia la línea del número siguiente). Nunca juntes la respuesta de un ítem con la del ítem de abajo.
- REGLA ESTRICTA DE LÍMITES: Transcribe TODAS las palabras que queden dentro del círculo o subrayado, aunque se pase hacia la derecha o izquierda (por ejemplo, si el círculo abarca 'the park', escribe exactamente 'the park' y NO lo recortes a 'the').
- Si el alumno encerró una palabra que NO es parte de las opciones (ej. 'gallery'), transcribe esa palabra exacta ('gallery'). NO adivines entre las opciones.
- NO corrijas faltas de ortografía ni tildes (transcribe exactamente como está escrito).
- ESCANEO DE DOBLE MARCA: En ejercicios de opción múltiple (A, B, C, etc.), revisa TODA la fila horizontal completa de izquierda a derecha antes de responder. Si ves dos o más letras encerradas, subrayadas o marcadas en la misma fila y NINGUNA está tachada, debes escribir obligatoriamente DOBLE_MARCA (nunca elijas solo la primera que veas).
- Si un trazo manuscrito es confuso o no se distingue con claridad: escribe exactamente ILEGIBLE.
- Códigos de anomalía obligatorios:
  * Espacio vacío: EN_BLANCO
  * Dos o más opciones marcadas/encerradas a la vez sin tachar ninguna: DOBLE_MARCA
  * Tachón o letra imposible de leer con seguridad: ILEGIBLE
  * Si tachó claramente una opción y marcó otra limpia al lado, transcribe la opción limpia Y agrega ese ID a 'items_con_tachon'.
"""


# ---------------------------------------------------------
# 3. FUNCIONES DE PROCESAMIENTO E IA
# ---------------------------------------------------------
def archivo_a_imagenes_pil(nombre_archivo: str, datos_bytes: bytes, dpi: int = 180) -> List[Image.Image]:
    if nombre_archivo.lower().endswith(".pdf"):
        imagenes = []
        with pymupdf.open(stream=datos_bytes, filetype="pdf") as doc:
            for pag in doc:
                pix = pag.get_pixmap(dpi=dpi)
                imagenes.append(Image.open(BytesIO(pix.tobytes("png"))).convert("RGB"))
        return imagenes
    return [Image.open(BytesIO(datos_bytes)).convert("RGB")]


def extraer_clave_desde_imagen(client, imagenes_pil) -> List[dict]:
    if not isinstance(imagenes_pil, list):
        imagenes_pil = [imagenes_pil]

    prompt_clave = (
        "Extrae todos los ejercicios y sus respuestas correctas de este Answer Key. "
        "Si el documento incluye la hoja de ejercicios y aparte la página de respuestas, "
        "usa únicamente las respuestas del Answer Key. "
        "Respeta el número exacto de cada ítem (si empieza en 2 porque el 1 es ejemplo, pon inicia_en=2). "
        "Si un ítem tiene varias respuestas válidas, sepáralas con '/'. No inventes respuestas."
    )

    ultimo_error = None
    for modelo in MODELOS_RESPALDO:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=[prompt_clave, *imagenes_pil],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ClaveCompletaIA,
                    temperature=0.0,
                ),
            )
            data = ClaveCompletaIA.model_validate_json(response.text)
            bloques_out = []
            for b in data.bloques:
                bloques_out.append(
                    {
                        "nombre": b.nombre,
                        "tipo": b.tipo if b.tipo in TIPOS_EJERCICIO else "escribir_palabra",
                        "inicia_en": max(1, int(b.inicia_en)),
                        "cantidad": max(1, len(b.respuestas)),
                        "pts_item": 1.0,
                        "ubicacion": b.ubicacion if b.ubicacion in UBICACIONES else "Completa",
                        "respuestas": b.respuestas if b.respuestas else [""],
                    }
                )
            return bloques_out
        except Exception as e:
            ultimo_error = e
    raise RuntimeError(f"No se pudo leer el Answer Key con ningún modelo: {ultimo_error}")


def leer_hoja_estudiante(
    client,
    imagenes_pil: List[Image.Image],
    bloques: List[dict],
    dos_columnas: bool,
    nota_extra: str,
) -> HojaEstudianteIA:
    estructura_items = []
    for b in bloques:
        inicio = int(b["inicia_en"])
        ubic = b.get("ubicacion", "Completa") if dos_columnas else "Completa"
        for idx in range(int(b["cantidad"])):
            num_item = inicio + idx
            id_item = f"{b['nombre']} - ítem {num_item}"
            estructura_items.append(
                f"- {id_item} | tipo: {b['tipo']} | columna: {ubic}"
            )

    lista_texto = "\n".join(estructura_items)
    nota_usuario = f"\nNOTA ADICIONAL DEL DOCENTE:\n{nota_extra.strip()}\n" if nota_extra.strip() else ""

    prompt_final = (
        f"{PROMPT_ANTI_SESGO}\n"
        f"{nota_usuario}\n"
        f"Debes identificar el nombre del estudiante (si no hay nombre escrito, devuelve exactamente 'SIN NOMBRE') "
        f"y transcribir EXCLUSIVAMENTE los siguientes ítems usando exactamente el mismo id_item:\n"
        f"{lista_texto}"
    )

    ultimo_error = None
    for modelo in MODELOS_RESPALDO:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=[prompt_final, *imagenes_pil],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=HojaEstudianteIA,
                    temperature=0.0,
                ),
            )
            return HojaEstudianteIA.model_validate_json(response.text)
        except Exception as e:
            ultimo_error = e
    raise RuntimeError(f"Error al procesar la hoja con IA: {ultimo_error}")


# ---------------------------------------------------------
# 4. LÓGICA DE CALIFICACIÓN Y EXCEL
# ---------------------------------------------------------
def normalizar(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "").strip()).lower()


def evaluar_respuesta(
    resp_est: str,
    clave_raw: str,
    tipo: str,
    pts_max: float,
    tiene_tachon: bool,
    permitir_ortografia: bool,
    umbral_ortografia: float,
    factor_ortografia: float,
):
    resp_limpia = str(resp_est or "").strip()
    sufijo_tachon = " (con tachón)" if tiene_tachon else ""

    if resp_limpia in ("EN_BLANCO", "DOBLE_MARCA", "ILEGIBLE", ""):
        obs = f"Anomalía: {resp_limpia or 'EN_BLANCO'}{sufijo_tachon}"
        return 0.0, "Incorrecto", obs

    alternativas = [a.strip() for a in str(clave_raw or "").split("/") if a.strip()]
    resp_norm = normalizar(resp_limpia)

    # Coincidencia exacta
    for alt in alternativas:
        if resp_norm == normalizar(alt):
            return pts_max, "Correcto", f"OK{sufijo_tachon}"

    # Tolerancia ortográfica solo para escribir_palabra
    if permitir_ortografia and tipo == "escribir_palabra" and alternativas:
        mejor_ratio = max(
            difflib.SequenceMatcher(None, resp_norm, normalizar(alt)).ratio()
            for alt in alternativas
        )
        if mejor_ratio >= umbral_ortografia:
            pts = round(pts_max * factor_ortografia, 2)
            porcentaje = int(round(mejor_ratio * 100))
            return pts, "Ortografía", f"Posible falta ortográfica ({porcentaje}% similar){sufijo_tachon}"

    return 0.0, "Incorrecto", f"Respuesta incorrecta{sufijo_tachon}"


def exportar_clave_excel(bloques: List[dict]) -> bytes:
    filas = []
    for b in bloques:
        inicio = int(b["inicia_en"])
        for idx, resp in enumerate(b["respuestas"]):
            filas.append(
                {
                    "Bloque": b["nombre"],
                    "Tipo": b["tipo"],
                    "Inicia_en": inicio,
                    "Item": inicio + idx,
                    "Pts_item": b["pts_item"],
                    "Ubicacion": b["ubicacion"],
                    "Respuesta": resp,
                }
            )
    df = pd.DataFrame(filas)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Clave")
    return buf.getvalue()


def cargar_clave_excel(archivo_bytes: bytes) -> List[dict]:
    df = pd.read_excel(BytesIO(archivo_bytes))
    bloques = []
    for nombre, grupo in df.groupby("Bloque", sort=False):
        grupo = grupo.sort_values("Item")
        bloques.append(
            {
                "nombre": str(nombre),
                "tipo": str(grupo["Tipo"].iloc[0]),
                "inicia_en": int(grupo["Inicia_en"].iloc[0]),
                "cantidad": len(grupo),
                "pts_item": float(grupo["Pts_item"].iloc[0]),
                "ubicacion": str(grupo["Ubicacion"].iloc[0]),
                "respuestas": [str(x) if pd.notna(x) else "" for x in grupo["Respuesta"].tolist()],
            }
        )
    return bloques


def generar_excel_reporte(df_resumen: pd.DataFrame, df_detalle: pd.DataFrame) -> bytes:
    df_errores = df_detalle[df_detalle["Estado"] != "Correcto"].copy()

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, index=False, sheet_name="Resumen")
        df_detalle.to_excel(writer, index=False, sheet_name="Transcripciones")
        df_errores.to_excel(writer, index=False, sheet_name="Errores")

    buf.seek(0)
    wb = openpyxl.load_workbook(buf)

    fill_header = PatternFill(start_color="111111", end_color="111111", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True)
    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_ok = Font(color="006100", bold=True)
    fill_err = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_err = Font(color="9C0006", bold=True)
    fill_ort = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_ort = Font(color="9C6500", bold=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center")

        if sheet_name in ("Transcripciones", "Errores"):
            headers = [c.value for c in ws[1]]
            if "Estado" in headers:
                col_estado = headers.index("Estado") + 1
                for row in range(2, ws.max_row + 1):
                    estado = ws.cell(row=row, column=col_estado).value
                    if estado == "Correcto":
                        ws.cell(row=row, column=col_estado).fill = fill_ok
                        ws.cell(row=row, column=col_estado).font = font_ok
                    elif estado == "Ortografía":
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=c).fill = fill_ort
                        ws.cell(row=row, column=col_estado).font = font_ort
                    elif estado == "Incorrecto":
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=c).fill = fill_err
                        ws.cell(row=row, column=col_estado).font = font_err

        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------------------------------------------------------
# 5. ESTADO DE SESIÓN
# ---------------------------------------------------------
if "bloques" not in st.session_state:
    st.session_state.bloques = [
        {
            "nombre": "Ejercicio 1",
            "tipo": "elegir_palabra",
            "inicia_en": 1,
            "cantidad": 1,
            "pts_item": 1.0,
            "ubicacion": "Izquierda",
            "respuestas": [""],
        }
    ]


# ---------------------------------------------------------
# 6. BARRA LATERAL (CONFIGURACIÓN)
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    dos_columnas = st.toggle("Hoja dividida en 2 columnas", value=True)
    escala = st.selectbox(
        "Escala de calificación",
        ["Sobre 10", "Sobre 20", "Sobre 100", "Puntaje directo"],
        index=0,
    )

    st.divider()
    st.markdown("## 📏 Reglas de calificación")
    permitir_ortografia = st.checkbox("Dar puntaje parcial por ortografía", value=True)
    umbral_ortografia = st.slider(
        "Similitud mínima (%)", min_value=70, max_value=95, value=80, step=5
    ) / 100.0
    factor_ortografia = st.slider(
        "Fracción del puntaje por ortografía", min_value=0.0, max_value=1.0, value=0.5, step=0.25
    )


# ---------------------------------------------------------
# 7. INTERFAZ PRINCIPAL
# ---------------------------------------------------------
st.markdown("# 📝 Calificador Automático de Hojas de Trabajo")
st.markdown(
    "Arma tu clave por bloques de ejercicios, sube las fotos o PDFs de los estudiantes y descarga todos los reportes listos."
)

st.info(
    "📌 **Indicaciones para que la lectura sea exacta:** En ejercicios de opción múltiple o subrayado, "
    "asegúrate de que las fotos estén derechas y bien iluminadas. Si el ítem 1 es un ejemplo impreso, "
    "ajusta **Inicia en #** a `2`."
)

# --- SECCIÓN 1: ACTIVIDAD Y CLAVE ---
st.markdown("## 1. Actividad y Clave por Bloques")

col_act, col_xlsx, col_ak = st.columns([1.2, 1.2, 1.4])

with col_act:
    nombre_actividad = st.text_input("Nombre de la actividad", value="CL_22_11_2026")

with col_xlsx:
    archivo_clave_xlsx = st.file_uploader(
        "Cargar clave.xlsx guardada (opcional)", type=["xlsx"], key="up_clave_xlsx"
    )
    if archivo_clave_xlsx is not None:
        if st.button("Cargar desde Excel", use_container_width=True):
            try:
                st.session_state.bloques = cargar_clave_excel(archivo_clave_xlsx.getvalue())
                st.success("Clave cargada desde Excel.")
                st.rerun()
            except Exception as e:
                st.error(f"Error leyendo clave.xlsx: {e}")

with col_ak:
    archivo_answer_key = st.file_uploader(
        "Subir foto/PDF del Answer Key (IA)",
        type=["png", "jpg", "jpeg", "pdf"],
        key="up_ak",
    )
    if st.button("✨ Leer Answer Key con IA", use_container_width=True):
        if archivo_answer_key is None:
            st.warning("Primero sube una foto o PDF del Answer Key.")
        else:
            with st.spinner("Leyendo Answer Key con Gemini..."):
                try:
                    client = obtener_cliente_genai()
                    datos_ak = archivo_answer_key.getvalue()
                    imgs_ak = archivo_a_imagenes_pil(archivo_answer_key.name, datos_ak, dpi=150)
                    st.session_state.bloques = extraer_clave_desde_imagen(client, imgs_ak)
                    st.success(f"Se extrajeron {len(st.session_state.bloques)} bloques correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al leer Answer Key: {e}")

nota_ia = st.text_area(
    "💡 Nota para la IA en esta tarea (opcional)",
    placeholder="Ej: El nombre del alumno está abajo a la derecha. Ignora el ítem 1 porque es ejemplo.",
    height=70,
)

# Editor de bloques
bloques_a_eliminar = []
for i, bloque in enumerate(st.session_state.bloques):
    with st.container(border=True):
        c1, c2, c3, c4, c5, c6, c7 = st.columns([1.6, 1.5, 0.9, 0.9, 1.0, 1.2, 0.5])
        bloque["nombre"] = c1.text_input("Nombre del bloque", value=bloque["nombre"], key=f"b_nom_{i}")
        bloque["tipo"] = c2.selectbox(
            "Tipo de ejercicio",
            TIPOS_EJERCICIO,
            index=TIPOS_EJERCICIO.index(bloque["tipo"]) if bloque["tipo"] in TIPOS_EJERCICIO else 0,
            key=f"b_tipo_{i}",
        )
        bloque["inicia_en"] = c3.number_input(
            "Inicia en #", min_value=1, value=int(bloque["inicia_en"]), step=1, key=f"b_ini_{i}"
        )
        nueva_cant = c4.number_input(
            "Cantidad", min_value=1, max_value=100, value=int(bloque["cantidad"]), step=1, key=f"b_cant_{i}"
        )
        bloque["pts_item"] = c5.number_input(
            "Pts / ítem", min_value=0.1, value=float(bloque["pts_item"]), step=0.25, key=f"b_pts_{i}"
        )
        bloque["ubicacion"] = c6.selectbox(
            "Ubicación",
            UBICACIONES,
            index=UBICACIONES.index(bloque["ubicacion"]) if bloque["ubicacion"] in UBICACIONES else 0,
            key=f"b_ubi_{i}",
        )
        if c7.button("🗑️", key=f"b_del_{i}"):
            bloques_a_eliminar.append(i)

        # Ajustar tamaño de lista de respuestas
        if nueva_cant != len(bloque["respuestas"]):
            if nueva_cant > len(bloque["respuestas"]):
                bloque["respuestas"].extend([""] * (nueva_cant - len(bloque["respuestas"])))
            else:
                bloque["respuestas"] = bloque["respuestas"][:nueva_cant]
            bloque["cantidad"] = nueva_cant

        st.markdown("**RESPUESTAS:** *(separa alternativas válidas con `/`)*")
        cols_resp = st.columns(4)
        for j in range(int(bloque["cantidad"])):
            num_item = int(bloque["inicia_en"]) + j
            with cols_resp[j % 4]:
                bloque["respuestas"][j] = st.text_input(
                    f"Ítem {num_item}",
                    value=bloque["respuestas"][j],
                    key=f"b_{i}_r_{j}",
                )

if bloques_a_eliminar:
    for idx in sorted(bloques_a_eliminar, reverse=True):
        st.session_state.bloques.pop(idx)
    st.rerun()

col_add, col_dl_clave = st.columns([1, 1])
with col_add:
    if st.button("➕ Agregar bloque de ejercicio", use_container_width=True):
        nuevo_num = len(st.session_state.bloques) + 1
        st.session_state.bloques.append(
            {
                "nombre": f"Ejercicio {nuevo_num}",
                "tipo": "escribir_palabra",
                "inicia_en": 1,
                "cantidad": 5,
                "pts_item": 1.0,
                "ubicacion": "Completa",
                "respuestas": [""] * 5,
            }
        )
        st.rerun()

with col_dl_clave:
    st.download_button(
        "💾 Descargar clave.xlsx",
        data=exportar_clave_excel(st.session_state.bloques),
        file_name=f"clave_{nombre_actividad}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# --- SECCIÓN 2: SUBIR HOJAS Y CALIFICAR ---
st.divider()
st.markdown("## 2. Subir Hojas de Estudiantes y Calificar")

archivos_estudiantes = st.file_uploader(
    "Sube las fotos o PDFs de los estudiantes (puedes seleccionar varios a la vez)",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
)

if st.button("🚀 Calificar Hojas", type="primary", use_container_width=True):
    if not archivos_estudiantes:
        st.warning("Sube al menos una hoja de estudiante para calificar.")
    else:
        client = obtener_cliente_genai()
        filas_resumen = []
        filas_detalle_global = []
        reportes_individuales = {}

        barra = st.progress(0, text="Iniciando calificación...")
        total_archivos = len(archivos_estudiantes)

        # Mapa de clave
        mapa_clave = {}
        puntaje_max_total = 0.0
        for b in st.session_state.bloques:
            inicio = int(b["inicia_en"])
            for idx, resp in enumerate(b["respuestas"]):
                id_item = f"{b['nombre']} - ítem {inicio + idx}"
                mapa_clave[id_item] = {
                    "tipo": b["tipo"],
                    "clave": resp,
                    "pts_max": float(b["pts_item"]),
                }
                puntaje_max_total += float(b["pts_item"])

        for idx_arch, archivo in enumerate(archivos_estudiantes):
            barra.progress(
                idx_arch / total_archivos,
                text=f"Calificando {archivo.name} ({idx_arch + 1}/{total_archivos})...",
            )
            try:
                imgs = archivo_a_imagenes_pil(archivo.name, archivo.getvalue(), dpi=180)
                resultado = leer_hoja_estudiante(
                    client, imgs, st.session_state.bloques, dos_columnas, nota_ia
                )
                nombre_est = (resultado.nombre_estudiante or "").strip() or "SIN NOMBRE"
                mapa_resp_est = {r.id_item: r.respuesta_estudiante for r in resultado.respuestas}
                set_tachones = set(resultado.items_con_tachon or [])

                filas_est = []
                pts_obtenidos = 0.0
                correctas = 0
                ortografia = 0
                incorrectas = 0

                for id_item, info in mapa_clave.items():
                    resp_est = mapa_resp_est.get(id_item, "EN_BLANCO")
                    tiene_tachon = id_item in set_tachones
                    pts, estado, obs = evaluar_respuesta(
                        resp_est=resp_est,
                        clave_raw=info["clave"],
                        tipo=info["tipo"],
                        pts_max=info["pts_max"],
                        tiene_tachon=tiene_tachon,
                        permitir_ortografia=permitir_ortografia,
                        umbral_ortografia=umbral_ortografia,
                        factor_ortografia=factor_ortografia,
                    )
                    pts_obtenidos += pts
                    if estado == "Correcto":
                        correctas += 1
                    elif estado == "Ortografía":
                        ortografia += 1
                    else:
                        incorrectas += 1

                    fila = {
                        "Estudiante": nombre_est,
                        "Ítem": id_item,
                        "Tipo": info["tipo"],
                        "Respuesta Estudiante": resp_est,
                        "Clave Oficial": info["clave"],
                        "Puntaje": pts,
                        "Máximo": info["pts_max"],
                        "Estado": estado,
                        "Observación": obs,
                    }
                    filas_est.append(fila)
                    filas_detalle_global.append(fila)

                # Calcular nota según escala
                if escala == "Sobre 10" and puntaje_max_total > 0:
                    nota_final = round((pts_obtenidos / puntaje_max_total) * 10, 2)
                elif escala == "Sobre 20" and puntaje_max_total > 0:
                    nota_final = round((pts_obtenidos / puntaje_max_total) * 20, 2)
                elif escala == "Sobre 100" and puntaje_max_total > 0:
                    nota_final = round((pts_obtenidos / puntaje_max_total) * 100, 2)
                else:
                    nota_final = round(pts_obtenidos, 2)

                resumen_est = {
                    "Estudiante": nombre_est,
                    "Archivo": archivo.name,
                    "Correctas": correctas,
                    "Ortografía": ortografia,
                    "Incorrectas": incorrectas,
                    "Puntaje": round(pts_obtenidos, 2),
                    "Puntaje Máximo": round(puntaje_max_total, 2),
                    f"Nota ({escala})": nota_final,
                }
                filas_resumen.append(resumen_est)

                # Excel individual
                df_res_ind = pd.DataFrame([resumen_est])
                df_det_ind = pd.DataFrame(filas_est)
                nombre_limpio = re.sub(r"[^A-Za-z0-9_\- ]+", "", nombre_est).strip().replace(" ", "_")
                reportes_individuales[f"{nombre_actividad}_{nombre_limpio or 'SIN_NOMBRE'}.xlsx"] = (
                    generar_excel_reporte(df_res_ind, df_det_ind)
                )

            except Exception as e:
                st.error(f"Error procesando {archivo.name}: {e}")

        barra.progress(1.0, text="¡Calificación completada!")

        if filas_resumen:
            df_resumen_total = pd.DataFrame(filas_resumen)
            df_detalle_total = pd.DataFrame(filas_detalle_global)

            st.success(f"Se calificaron {len(filas_resumen)} hojas correctamente.")
            st.dataframe(df_resumen_total, use_container_width=True)

            excel_consolidado = generar_excel_reporte(df_resumen_total, df_detalle_total)

            # Crear ZIP con consolidado + individuales
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"Consolidado_{nombre_actividad}.xlsx", excel_consolidado)
                for fname, fbytes in reportes_individuales.items():
                    zf.writestr(f"Individuales/{fname}", fbytes)

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    "📊 Descargar Excel Consolidado",
                    data=excel_consolidado,
                    file_name=f"Consolidado_{nombre_actividad}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_d2:
                st.download_button(
                    "📦 Descargar ZIP Completo (Consolidado + Individuales)",
                    data=zip_buf.getvalue(),
                    file_name=f"Reportes_{nombre_actividad}.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
