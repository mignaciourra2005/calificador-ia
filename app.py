import os
import re
import json
import time
import uuid
import base64
import zipfile
import difflib
import unicodedata
import pymupdf
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
from dotenv import load_dotenv
from pydantic import BaseModel, Field, create_model
from google import genai
from google.genai import types
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# CONFIGURACIÓN DE PÁGINA (OPTIMIZADA PARA PC)
# ==========================================
st.set_page_config(
    page_title="Calificador Automático con IA",
    page_icon="📝",
    layout="wide",
)

# ==========================================
# FONDO Y ESTILOS CSS (DM SANS 21.5PX + BOTONES NEGROS + FIX UPLOAD)
# ==========================================
def cargar_fondo_base64(ruta_imagen):
    if os.path.exists(ruta_imagen):
        with open(ruta_imagen, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

fondo_b64 = cargar_fondo_base64("fondo.jpg")
css_fondo = f"""
    .stApp {{
        background-image: url("data:image/png;base64,{fondo_b64}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}
    .main .block-container {{
        background-color: rgba(255, 255, 255, 0.94);
        border-radius: 16px;
        padding: 2rem 2.5rem !important;
        margin-top: 1.5rem;
        margin-bottom: 2rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }}
""" if fondo_b64 else ""

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@500;700;800&display=swap');
    {css_fondo}

    /* Tipografía DM Sans sin romper íconos internos de Streamlit */
    :root {{
        --font: 'DM Sans', sans-serif !important;
    }}

    html, body, .stApp, [data-testid="stSidebar"],
    input, textarea, select, p, label, h1, h2, h3 {{
        font-family: 'DM Sans', sans-serif !important;
    }}

    .stButton > button, .stDownloadButton > button {{
        font-family: 'DM Sans', sans-serif !important;
    }}

    /* Tamaño base +2px (21.5px) y negrita */
    html, body, .stApp, [data-testid="stSidebar"] {{
        font-size: 21.5px !important;
        font-weight: 700 !important;
    }}

    h1, h2, h3, label,
    [data-testid="stWidgetLabel"] p,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label {{
        font-size: 21.5px !important;
        font-weight: 700 !important;
        margin-bottom: 4px !important;
    }}

    h1 {{ font-size: 3.0rem !important; }}
    h2 {{ font-size: 2.3rem !important; }}
    h3 {{ font-size: 1.85rem !important; }}

    /* Inputs, textareas y selectores */
    div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"] > input,
    div[data-baseweb="textarea"] > textarea,
    div[data-baseweb="select"] > div {{
        font-size: 21.5px !important;
        font-weight: 700 !important;
        min-height: 54px !important;
        padding: 8px 14px !important;
        border-radius: 8px !important;
    }}

    /* Botones generales */
    .stButton > button, .stDownloadButton > button {{
        font-size: 21.5px !important;
        font-weight: 700 !important;
        min-height: 54px !important;
        padding: 10px 18px !important;
        border-radius: 8px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        transition: all 0.15s ease-in-out !important;
    }}

    /* Botones de flechas (◀ / ▶) ACTIVOS en NEGRO */
    [class*="st-key-btn_izq_"] button[data-testid="stBaseButton-primary"],
    [class*="st-key-btn_der_"] button[data-testid="stBaseButton-primary"],
    [class*="st-key-btn_izq_"] button[kind="primary"],
    [class*="st-key-btn_der_"] button[kind="primary"] {{
        background-color: #000000 !important;
        color: #FFFFFF !important;
        border: 1.5px solid #000000 !important;
    }}

    [class*="st-key-btn_izq_"] button[data-testid="stBaseButton-primary"]:hover,
    [class*="st-key-btn_der_"] button[data-testid="stBaseButton-primary"]:hover {{
        background-color: #262626 !important;
        border-color: #262626 !important;
        color: #FFFFFF !important;
    }}

    /* Botones de flechas (◀ / ▶) INACTIVOS en blanco con icono negro */
    [class*="st-key-btn_izq_"] button[data-testid="stBaseButton-secondary"],
    [class*="st-key-btn_der_"] button[data-testid="stBaseButton-secondary"],
    [class*="st-key-btn_izq_"] button[kind="secondary"],
    [class*="st-key-btn_der_"] button[kind="secondary"] {{
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1.5px solid #CBD5E1 !important;
    }}

    [class*="st-key-btn_izq_"] button[data-testid="stBaseButton-secondary"]:hover,
    [class*="st-key-btn_der_"] button[data-testid="stBaseButton-secondary"]:hover {{
        background-color: #F3F4F6 !important;
        border-color: #94A3B8 !important;
    }}

    /* Tamaño y centrado de las flechas ◀ / ▶ */
    [class*="st-key-btn_izq_"] button p,
    [class*="st-key-btn_der_"] button p {{
        font-size: 26px !important;
        line-height: 1 !important;
        margin: 0 !important;
        padding: 0 !important;
    }}

    [class*="st-key-btn_izq_"] button,
    [class*="st-key-btn_der_"] button {{
        padding: 0 !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

load_dotenv()

def obtener_cliente():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key and "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    if not api_key:
        st.error("No se encontró GEMINI_API_KEY. Configúrala en .env (local) o en Secrets (Streamlit Cloud).")
        st.stop()
    return genai.Client(api_key=api_key)

MODELOS_RESPALDO = [
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.8-flash",
]

PROMPT_ANTI_SESGO = """
Eres un escáner OCR puramente visual y estricto. NO eres profesor.
ADVERTENCIA CRÍTICA: Esta hoja contiene errores, tachones y marcas imprecisas hechas por estudiantes.
- PROHIBIDO deducir la respuesta correcta leyendo el texto de la hoja.
- Fíjate ÚNICAMENTE en la forma visual de las letras escritas a mano y en qué texto exacto está dentro del círculo o sobre el subrayado.
- REGLA ESTRICTA DE LÍMITES: Transcribe TODAS las palabras que queden dentro del círculo o subrayado, aunque se pase hacia la derecha o izquierda (por ejemplo, si el círculo abarca 'the park', escribe exactamente 'the park' y NO lo recortes a 'the').
- Si el alumno encerró una palabra que NO es parte de las opciones (ej. 'gallery'), transcribe esa palabra exacta ('gallery'). NO adivines entre las opciones.
- NO corrijas faltas de ortografía ni tildes (transcribe exactamente como está escrito).
- Si un trazo manuscrito es confuso o no se distingue con claridad: escribe exactamente ILEGIBLE.
- Códigos de anomalía obligatorios:
  * Espacio vacío: EN_BLANCO
  * Dos opciones marcadas/encerradas a la vez sin tachar ninguna: DOBLE_MARCA
  * Tachón o letra imposible de leer con seguridad: ILEGIBLE
  * Si tachó claramente una opción y marcó otra limpia al lado, transcribe la opción limpia Y agrega ese ID a 'items_con_tachon'.
"""

TIPOS_EJERCICIO = ["elegir_palabra", "escribir_palabra", "encerrar_opcion"]
UBICACIONES = ["Izquierda", "Derecha", "Mixta"]
ESCALAS_NOTA = {
    "Sobre 10": (10.0, "Nota / 10"),
    "Sobre 20": (20.0, "Nota / 20"),
    "Sobre 100": (100.0, "Nota / 100"),
    "Puntaje directo": (None, "Puntaje Total"),
}

COLORES_PASTEL = [
    ("#EEF4FF", "#C3D7FA"),  # Azul suave
    ("#F0FAF4", "#BDE6CE"),  # Verde menta suave
    ("#FFF7EB", "#FADBB3"),  # Durazno suave
    ("#F7F0FF", "#DEC7FA"),  # Lavanda suave
    ("#FFF0F5", "#FAC8DA"),  # Rosa suave
    ("#F0FAFC", "#BDE5EE"),  # Celeste suave
]

# ==========================================
# MODELOS PARA LEER ANSWER KEY CON IA
# ==========================================
class ItemClave(BaseModel):
    numero: int = Field(description="Número del ítem (ej: 2, 3, 4)")
    respuesta: str = Field(description="Respuesta correcta exacta (si hay varias válidas, sepáralas con /)")

class EjercicioClave(BaseModel):
    titulo: str = Field(description="Nombre del ejercicio (ej: 'Ejercicio 1', 'Ejercicio 2')")
    tipo_sugerido: str = Field(description="Sugiere 'encerrar_opcion' si son letras A/B/C/D, 'elegir_palabra' si son artículos cortos, o 'escribir_palabra' si son palabras o frases")
    items: list[ItemClave]

class ClaveExtraida(BaseModel):
    ejercicios: list[EjercicioClave]

def extraer_clave_desde_imagen(client, imagen_pil) -> list[dict]:
    prompt_clave = (
        "Extrae todos los ejercicios y sus respuestas correctas de esta hoja de Answer Key. "
        "Respeta el número exacto de cada ítem (si empieza en 2, pon 2). No inventes respuestas."
    )
    for modelo in MODELOS_RESPALDO:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=[prompt_clave, imagen_pil],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ClaveExtraida,
                    temperature=0.0,
                ),
            )
            datos = json.loads(response.text)
            bloques_nuevos = []
            for ej in datos.get("ejercicios", []):
                items = sorted(ej.get("items", []), key=lambda x: x["numero"])
                if not items:
                    continue
                inicio = items[0]["numero"]
                cantidad = items[-1]["numero"] - inicio + 1
                tipo = ej.get("tipo_sugerido", "escribir_palabra")
                if tipo not in TIPOS_EJERCICIO:
                    tipo = "escribir_palabra"

                bloques_nuevos.append({
                    "uid": uuid.uuid4().hex[:8],
                    "titulo": ej["titulo"],
                    "tipo": tipo,
                    "inicio": inicio,
                    "cantidad": cantidad,
                    "puntos": 1.0,
                    "ubicacion": "Izquierda",
                    "respuestas": {it["numero"]: it["respuesta"] for it in items},
                    "columnas": {it["numero"]: "izq" for it in items},
                })
            return bloques_nuevos
        except Exception as e:
            if any(err in str(e) for err in ("503", "429", "404")):
                time.sleep(2)
                continue
            raise e
    raise RuntimeError("No se pudo leer el Answer Key en este momento.")

# ==========================================
# CONVERSIÓN ENTRE BLOQUES Y EXCEL (clave.xlsx)
# ==========================================
def df_a_bloques(df: pd.DataFrame) -> list[dict]:
    mapa_tipos = {"objetiva": "encerrar_opcion", "escrita": "escribir_palabra"}
    bloques_dict = {}

    for _, row in df.dropna(subset=["ID", "Respuesta_Correcta"]).iterrows():
        etiqueta = str(row["Etiqueta"]).strip()
        resp = str(row["Respuesta_Correcta"]).strip()
        tipo = mapa_tipos.get(str(row["Tipo"]).strip().lower(), str(row["Tipo"]).strip().lower())
        if tipo not in TIPOS_EJERCICIO:
            tipo = "escribir_palabra"
        col = str(row.get("Columna", "izq")).strip().lower()
        col = "der" if col == "der" else "izq"

        puntos_raw = row.get("Puntos", 1.0)
        try:
            puntos = float(puntos_raw) if pd.notna(puntos_raw) else 1.0
        except (ValueError, TypeError):
            puntos = 1.0

        partes = etiqueta.split(" - ")
        titulo = partes[0].strip() if len(partes) > 1 else "Ejercicio 1"
        m_num = re.search(r"(\d+)\s*$", etiqueta)
        num_item = int(m_num.group(1)) if m_num else 1

        if titulo not in bloques_dict:
            bloques_dict[titulo] = {
                "uid": uuid.uuid4().hex[:8],
                "titulo": titulo,
                "tipo": tipo,
                "puntos": puntos,
                "items": [],
            }
        bloques_dict[titulo]["items"].append((num_item, resp, col))

    bloques = []
    for b in bloques_dict.values():
        items_ord = sorted(b["items"], key=lambda x: x[0])
        nums = [x[0] for x in items_ord]
        cols = {n: c for n, _, c in items_ord}
        inicio = min(nums)
        cantidad = max(nums) - inicio + 1

        valores_col = set(cols.values())
        if valores_col == {"izq"}:
            ubicacion = "Izquierda"
        elif valores_col == {"der"}:
            ubicacion = "Derecha"
        else:
            ubicacion = "Mixta"

        bloques.append({
            "uid": b["uid"],
            "titulo": b["titulo"],
            "tipo": b["tipo"],
            "inicio": inicio,
            "cantidad": cantidad,
            "puntos": b["puntos"],
            "ubicacion": ubicacion,
            "respuestas": {n: r for n, r, _ in items_ord},
            "columnas": cols,
        })
    return bloques

def bloques_a_df(bloques: list[dict], usar_dos_columnas: bool) -> pd.DataFrame:
    filas = []
    for idx_b, b in enumerate(bloques, start=1):
        prefijo_id = re.sub(r"\W+", "_", b["titulo"].lower()).strip("_") or f"ej{idx_b}"
        inicio = int(b["inicio"])
        cantidad = int(b["cantidad"])
        puntos = float(b.get("puntos", 1.0))

        for num in range(inicio, inicio + cantidad):
            resp = str(b["respuestas"].get(num, "")).strip()
            if not resp:
                continue

            col = b["columnas"].get(num, "izq") if usar_dos_columnas else "izq"
            filas.append({
                "ID": f"{prefijo_id}_{num}",
                "Etiqueta": f"{b['titulo']} - ítem {num}",
                "Respuesta_Correcta": resp,
                "Tipo": b["tipo"],
                "Columna": col,
                "Puntos": puntos,
            })
    return pd.DataFrame(filas)

# ==========================================
# FUNCIONES DE IA Y EXCEL
# ==========================================
def construir_esquemas(df_clave, usar_dos_columnas):
    clave_dict = {}
    campos_izq = {
        "nombre_alumno": (str, Field(description="Nombre escrito a mano arriba, o 'Desconocido'")),
        "items_con_tachon": (list[str], Field(description="IDs de esta sección con tachón o borrón")),
    }
    campos_der = {
        "items_con_tachon": (list[str], Field(description="IDs de esta sección con tachón o borrón")),
    }

    descripciones_tipo = {
        "elegir_palabra": "Palabra exacta subrayada o encerrada a mano en",
        "escribir_palabra": "Texto exacto escrito a mano en",
        "encerrar_opcion": "Letra/opción exacta encerrada en un círculo o escrita a mano en",
    }

    for _, row in df_clave.iterrows():
        cid = re.sub(r"\W+", "_", str(row["ID"]).strip())
        etiqueta = str(row["Etiqueta"]).strip()
        resp = str(row["Respuesta_Correcta"]).strip()
        tipo = str(row["Tipo"]).strip().lower()
        col = str(row["Columna"]).strip().lower() if usar_dos_columnas else "izq"
        peso = float(row.get("Puntos", 1.0))

        clave_dict[cid] = (etiqueta, resp, tipo, peso)
        prefijo = descripciones_tipo.get(tipo, "Respuesta exacta marcada o escrita en")
        desc = f"{prefijo} {etiqueta}, o código de anomalía (EN_BLANCO, DOBLE_MARCA, ILEGIBLE)"

        if col == "der":
            campos_der[cid] = (str, Field(description=desc))
        else:
            campos_izq[cid] = (str, Field(description=desc))

    modelo_izq = create_model("ColumnaIzquierda", **campos_izq)
    modelo_der = create_model("ColumnaDerecha", **campos_der) if len(campos_der) > 1 else None
    return clave_dict, modelo_izq, modelo_der

def llamar_gemini(client, imagen_pil, esquema, instrucciones_extra=""):
    prompt_final = PROMPT_ANTI_SESGO
    if instrucciones_extra.strip():
        prompt_final += f"\n\nINSTRUCCIONES EXTRA PARA ESTA TAREA:\n{instrucciones_extra.strip()}"

    for modelo in MODELOS_RESPALDO:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=[prompt_final, imagen_pil],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=esquema,
                    temperature=0.0,
                ),
            )
            return json.loads(response.text)
        except Exception as e:
            if any(err in str(e) for err in ("503", "429", "404")):
                time.sleep(2)
                continue
            raise e
    raise RuntimeError("Todos los modelos están saturados en este momento.")

def normalizar(texto: str, importan_tildes: bool = False) -> str:
    texto = texto.strip().lower()
    if not importan_tildes:
        texto = "".join(
            c for c in unicodedata.normalize("NFD", texto)
            if unicodedata.category(c) != "Mn"
        )
    texto = re.sub(r"[^\w\s]", "", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto

def limpiar_nombre(texto: str) -> str:
    texto = re.sub(r'[\\/*?:"<>|]', "", texto.strip())
    texto = re.sub(r"\s+", "_", texto)
    return texto or "Desconocido"

def evaluar_item(
    resp_raw, respuesta_correcta, tipo, peso_item,
    tuvo_tachon, factor_orto, umbral_orto,
    doble_marca_revisar, importan_tildes
):
    codigo = resp_raw.strip().upper()
    if codigo == "EN_BLANCO":
        return 0.0, "Incorrecto", "Sin respuesta"
    if codigo == "DOBLE_MARCA":
        if doble_marca_revisar:
            return 0.0, "Revisar", "Doble respuesta marcada (revisar manualmente)"
        return 0.0, "Incorrecto", "Doble respuesta marcada (0 automático)"
    if codigo == "ILEGIBLE":
        return 0.0, "Revisar", "Letra ilegible o trazo confuso"

    r_norm = normalizar(resp_raw, importan_tildes)
    alternativas = [alt.strip() for alt in respuesta_correcta.split("/") if alt.strip()]
    alts_norm = [normalizar(alt, importan_tildes) for alt in alternativas]

    if r_norm in alts_norm:
        obs = "Correcto (con tachón previo)" if tuvo_tachon else "OK"
        return round(peso_item, 2), "Correcto", obs

    if tipo == "escribir_palabra":
        mejor_sim = 0.0
        for c_norm in alts_norm:
            if len(c_norm) >= 4:
                sim = difflib.SequenceMatcher(None, r_norm, c_norm).ratio()
                if sim > mejor_sim:
                    mejor_sim = sim
        if mejor_sim >= umbral_orto:
            pts_orto = round(peso_item * factor_orto, 2)
            return pts_orto, "Ortografía", f"Posible falta ortográfica ({int(mejor_sim * 100)}% similar)"

    obs = "Respuesta incorrecta (con tachón)" if tuvo_tachon else "Respuesta incorrecta"
    return 0.0, "Incorrecto", obs

def dar_formato_hoja(ws, tipo_hoja, enlaces=None):
    enlaces = enlaces or {}
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    rojo_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    rojo_font = Font(color="9C0006", bold=True)
    amarillo_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    amarillo_font = Font(color="9C6500", bold=True)
    verde_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    verde_font = Font(color="006100", bold=True)
    link_font = Font(color="0563C1", underline="single")
    borde = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"),
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    headers = [cell.value for cell in ws[1]]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = borde
        if "Estado" in headers:
            idx = headers.index("Estado")
            estado = row[idx].value
            if estado == "Incorrecto":
                if tipo_hoja == "transcripcion":
                    for c in row:
                        c.fill = rojo_fill
                row[idx].fill = rojo_fill
                row[idx].font = rojo_font
            elif estado in ("Revisar", "Ortografía"):
                for c in row:
                    c.fill = amarillo_fill
                row[idx].font = amarillo_font
            elif estado == "Correcto":
                row[idx].fill = verde_fill
                row[idx].font = verde_font

        if tipo_hoja == "errores" and "Respuesta Correcta" in headers:
            c_corr = row[headers.index("Respuesta Correcta")]
            c_corr.fill = verde_fill
            c_corr.font = verde_font

        if "Archivo" in headers:
            c_arch = row[headers.index("Archivo")]
            if c_arch.value in enlaces:
                c_arch.hyperlink = enlaces[c_arch.value]
                c_arch.font = link_font

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

def agregar_formulas(ws, num_alumnos, col_nota):
    fila_inicio = num_alumnos + 3
    rango = f"{col_nota}2:{col_nota}{num_alumnos + 1}"
    gris_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    negrita = Font(bold=True)
    for i, (etq, form) in enumerate([
        ("Promedio del curso", f"=AVERAGE({rango})"),
        ("Nota máxima", f"=MAX({rango})"),
        ("Nota mínima", f"=MIN({rango})"),
    ]):
        f = fila_inicio + i
        ws[f"A{f}"], ws[f"{col_nota}{f}"] = etq, form
        ws[f"{col_nota}{f}"].number_format = "0.00"
        ws[f"A{f}"].font = ws[f"{col_nota}{f}"].font = negrita
        ws[f"A{f}"].fill = ws[f"{col_nota}{f}"].fill = gris_fill

def generar_excel_bytes(resumen, errores, trans, enlaces=None, incluir_formulas=False):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(resumen).to_excel(writer, sheet_name="Resumen", index=False)
        if errores:
            pd.DataFrame(errores).to_excel(writer, sheet_name="Detalle_Errores", index=False)
        pd.DataFrame(trans).to_excel(writer, sheet_name="Transcripcion_Completa", index=False)

        ws_res = writer.sheets["Resumen"]
        if incluir_formulas and resumen:
            agregar_formulas(ws_res, len(resumen), "D")
        dar_formato_hoja(ws_res, "resumen", enlaces)
        if errores:
            dar_formato_hoja(writer.sheets["Detalle_Errores"], "errores", enlaces)
        dar_formato_hoja(writer.sheets["Transcripcion_Completa"], "transcripcion", enlaces)
    return buf.getvalue()

# ==========================================
# INICIALIZAR ESTADO DE BLOQUES
# ==========================================
if "bloques" not in st.session_state:
    if os.path.exists("clave.xlsx"):
        st.session_state.bloques = df_a_bloques(pd.read_excel("clave.xlsx"))
    else:
        st.session_state.bloques = [{
            "uid": uuid.uuid4().hex[:8],
            "titulo": "Ejercicio 1",
            "tipo": "elegir_palabra",
            "inicio": 1,
            "cantidad": 5,
            "puntos": 1.0,
            "ubicacion": "Izquierda",
            "respuestas": {},
            "columnas": {},
        }]

if "ultimo_archivo_clave" not in st.session_state:
    st.session_state.ultimo_archivo_clave = None

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("📝 Calificador Automático de Hojas de Trabajo")
st.caption("Arma tu clave por bloques de ejercicios, sube las fotos o PDFs de los estudiantes y descarga todos los reportes listos.")

st.info(
    "📌 **Indicaciones para que la lectura sea exacta:** "
    "En ejercicios de encerrar o subrayar, el estudiante debe marcar **únicamente** la palabra o letra elegida, "
    "sin abarcar palabras vecinas (ej. encerrar solo *the*, no *the park*). "
    "Para aceptar varias respuestas válidas en un mismo ítem, sepáralas con `/` (ej. `don't / do not`)."
)

with st.sidebar:
    st.header("⚙️ Configuración")
    usar_dos_columnas = st.toggle("Hoja dividida en 2 columnas", value=True)
    escala_elegida = st.selectbox("Escala de calificación", options=list(ESCALAS_NOTA.keys()), index=0)
    escala_valor, col_nota_nombre = ESCALAS_NOTA[escala_elegida]

    st.divider()
    st.subheader("📏 Reglas de calificación")
    factor_orto = st.selectbox(
        "Valor por falta ortográfica leve",
        options=[0.5, 0.0, 1.0],
        index=0,
        format_func=lambda x: f"{int(x * 100)}% del ítem",
    )
    umbral_orto = st.slider("Similitud mínima para ortografía (%)", 70, 95, 85) / 100.0
    doble_marca_revisar = st.toggle(
        "Doble marca pasa a Revisar (amarillo)",
        value=False,
        help="Desactivado = 0 automático. Activado = se marca en amarillo para revisión manual.",
    )
    importan_tildes = st.toggle(
        "Evaluar tildes / acentos",
        value=False,
        help="Actívalo si calificas hojas en español donde las tildes son obligatorias.",
    )

    st.divider()
    pausa_seg = st.number_input("Pausa entre hojas (segundos)", min_value=0.0, max_value=10.0, value=2.0, step=0.5)

# --- PASO 1: ACTIVIDAD Y CARGA OPCIONAL ---
st.subheader("1. Actividad y Clave por Bloques")
c1, c2, c3 = st.columns([1, 1, 1])

with c1:
    sugerencia = f"CL_{datetime.now().strftime('%d_%m_%Y')}"
    nombre_actividad = limpiar_nombre(st.text_input("Nombre de la actividad", value=sugerencia))

with c2:
    archivo_clave = st.file_uploader("Cargar clave.xlsx guardada (opcional)", type=["xlsx"])
    if archivo_clave is not None and archivo_clave.name != st.session_state.ultimo_archivo_clave:
        st.session_state.bloques = df_a_bloques(pd.read_excel(archivo_clave))
        st.session_state.ultimo_archivo_clave = archivo_clave.name
        st.rerun()

with c3:
    archivo_answer_key = st.file_uploader("Subir foto/PDF del Answer Key (IA)", type=["jpg", "jpeg", "png", "pdf"])
    if archivo_answer_key is not None:
        if st.button("✨ Leer Answer Key con IA", width="stretch"):
            with st.spinner("Extrayendo respuestas del Answer Key..."):
                client = obtener_cliente()
                datos_ak = archivo_answer_key.getvalue()
                if archivo_answer_key.name.lower().endswith(".pdf"):
                    with pymupdf.open(stream=datos_ak, filetype="pdf") as doc:
                        pix = doc[0].get_pixmap(dpi=150)
                        img_ak = Image.open(BytesIO(pix.tobytes("png")))
                else:
                    img_ak = Image.open(BytesIO(datos_ak))
                st.session_state.bloques = extraer_clave_desde_imagen(client, img_ak)
                st.rerun()

instrucciones_extra = st.text_area(
    "💡 Nota para la IA en esta tarea (opcional)",
    placeholder="Ej: El nombre del alumno está abajo a la derecha. Ignora el ítem 1 porque es ejemplo.",
    height=75,
)

# --- CONSTRUCTOR VISUAL DE BLOQUES (1 COLUMNA + COLORES PASTEL + BOTONES ◀ ▶) ---
indices_a_eliminar = []

for i, bloque in enumerate(st.session_state.bloques):
    uid = bloque["uid"]
    if "columnas" not in bloque:
        bloque["columnas"] = {}
    if "puntos" not in bloque:
        bloque["puntos"] = 1.0

    fondo, borde_color = COLORES_PASTEL[i % len(COLORES_PASTEL)]
    st.markdown(
        f"""
        <style>
        .st-key-bloque_{uid} {{
            background-color: {fondo};
            border: 1.5px solid {borde_color} !important;
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 12px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True, key=f"bloque_{uid}"):
        col_tit, col_tipo, col_ini, col_cant, col_pts, col_ubi, col_del = st.columns(
            [2.0, 1.9, 1.0, 1.0, 1.1, 1.6, 0.65], vertical_alignment="bottom"
        )

        with col_tit:
            bloque["titulo"] = st.text_input("Nombre del bloque", value=bloque["titulo"], key=f"tit_{uid}")
        with col_tipo:
            idx_t = TIPOS_EJERCICIO.index(bloque["tipo"]) if bloque["tipo"] in TIPOS_EJERCICIO else 0
            bloque["tipo"] = st.selectbox("Tipo de ejercicio", TIPOS_EJERCICIO, index=idx_t, key=f"tipo_{uid}")
        with col_ini:
            bloque["inicio"] = st.number_input("Inicia en #", min_value=1, max_value=50, value=int(bloque["inicio"]), key=f"ini_{uid}")
        with col_cant:
            bloque["cantidad"] = st.number_input("Cantidad", min_value=1, max_value=30, value=int(bloque["cantidad"]), key=f"cant_{uid}")
        with col_pts:
            bloque["puntos"] = st.number_input("Pts / ítem", min_value=0.25, max_value=20.0, value=float(bloque["puntos"]), step=0.5, key=f"pts_{uid}")

        inicio = int(bloque["inicio"])
        fin = inicio + int(bloque["cantidad"]) - 1

        cols_bloque = {bloque["columnas"].get(n, "izq") for n in range(inicio, fin + 1)}
        ubi_actual = "Izquierda" if cols_bloque == {"izq"} else ("Derecha" if cols_bloque == {"der"} else "Mixta")

        with col_ubi:
            if usar_dos_columnas:
                nueva_ubi = st.selectbox(
                    "Ubicación",
                    UBICACIONES,
                    index=UBICACIONES.index(ubi_actual),
                    key=f"ubi_{uid}",
                )
                if nueva_ubi != ubi_actual and nueva_ubi in ("Izquierda", "Derecha"):
                    val_col = "izq" if nueva_ubi == "Izquierda" else "der"
                    for n in range(inicio, fin + 1):
                        bloque["columnas"][n] = val_col
                    st.rerun()
            else:
                st.text_input("Ubicación", value="Hoja completa", disabled=True, key=f"ubi_dis_{uid}")

        with col_del:
            if st.button("🗑️", key=f"del_{uid}", help="Eliminar este ejercicio", width="stretch"):
                indices_a_eliminar.append(i)

        st.markdown("**RESPUESTAS:** *(separa alternativas válidas con `/`)*")

        for num in range(inicio, fin + 1):
            col_actual = bloque["columnas"].get(num, "izq")
            c_resp, c_izq, c_der, _ = st.columns([3.5, 0.65, 0.65, 2.2], vertical_alignment="bottom")

            with c_resp:
                val_actual = bloque["respuestas"].get(num, "")
                bloque["respuestas"][num] = st.text_input(
                    f"Ítem {num}",
                    value=val_actual,
                    key=f"resp_{uid}_{num}",
                    placeholder=f"Respuesta del ítem {num}...",
                )

            if usar_dos_columnas:
                with c_izq:
                    if st.button(
                        "◀",
                        key=f"btn_izq_{uid}_{num}",
                        type="primary" if col_actual == "izq" else "secondary",
                        help="Columna izquierda",
                        width="stretch",
                    ):
                        bloque["columnas"][num] = "izq"
                        st.rerun()

                with c_der:
                    if st.button(
                        "▶",
                        key=f"btn_der_{uid}_{num}",
                        type="primary" if col_actual == "der" else "secondary",
                        help="Columna derecha",
                        width="stretch",
                    ):
                        bloque["columnas"][num] = "der"
                        st.rerun()

if indices_a_eliminar:
    for idx_del in sorted(indices_a_eliminar, reverse=True):
        st.session_state.bloques.pop(idx_del)
    st.rerun()

col_add, col_dl, _ = st.columns([1.8, 2.4, 2.8])
with col_add:
    if st.button("➕ Agregar Ejercicio", width="stretch"):
        nuevo_num = len(st.session_state.bloques) + 1
        st.session_state.bloques.append({
            "uid": uuid.uuid4().hex[:8],
            "titulo": f"Ejercicio {nuevo_num}",
            "tipo": "encerrar_opcion",
            "inicio": 1,
            "cantidad": 5,
            "puntos": 1.0,
            "ubicacion": "Izquierda",
            "respuestas": {},
            "columnas": {},
        })
        st.rerun()

df_clave_final = bloques_a_df(st.session_state.bloques, usar_dos_columnas)
total_puntos_clave = df_clave_final["Puntos"].sum() if not df_clave_final.empty else 0.0

with col_dl:
    buf_clave = BytesIO()
    df_clave_final.to_excel(buf_clave, index=False)
    st.download_button(
        f"💾 Guardar clave ({len(df_clave_final)} ítems · {total_puntos_clave:g} pts)",
        data=buf_clave.getvalue(),
        file_name=f"clave_{nombre_actividad}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

st.divider()

# --- PASO 2: SUBIR HOJAS ---
st.subheader("2. Sube las hojas de los estudiantes")
archivos_subidos = st.file_uploader(
    "Arrastra aquí las fotos (JPG, PNG) o PDFs",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True,
)

# --- PASO 3: CALIFICAR ---
if st.button("🚀 Calificar Hojas", type="primary", disabled=not archivos_subidos):
    if df_clave_final.empty:
        st.error("No hay respuestas configuradas. Llena al menos un ítem en tus bloques.")
        st.stop()

    client = obtener_cliente()
    clave_oficial, modelo_izq, modelo_der = construir_esquemas(df_clave_final, usar_dos_columnas)
    total_puntos_posibles = sum(item[3] for item in clave_oficial.values())

    paginas_a_procesar = []
    for archivo in archivos_subidos:
        ext = os.path.splitext(archivo.name)[1].lower()
        datos_bytes = archivo.getvalue()
        if ext == ".pdf":
            with pymupdf.open(stream=datos_bytes, filetype="pdf") as doc:
                for i, pag in enumerate(doc):
                    pix = pag.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")
                    paginas_a_procesar.append((f"{archivo.name} (Pág {i+1})", Image.open(BytesIO(img_bytes)), img_bytes, ".png"))
        else:
            paginas_a_procesar.append((archivo.name, Image.open(BytesIO(datos_bytes)), datos_bytes, ext))

    barra = st.progress(0, text="Iniciando calificación...")
    lista_resumen, lista_errores, lista_trans = [], [], []
    enlaces_consolidado = {}
    nombres_usados = set()
    fallidos = []

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (nombre_orig, img_pil, img_bytes, ext) in enumerate(paginas_a_procesar):
            barra.progress(idx / len(paginas_a_procesar), text=f"Leyendo [{idx+1}/{len(paginas_a_procesar)}]: {nombre_orig}...")
            try:
                if usar_dos_columnas and modelo_der is not None:
                    w, h = img_pil.size
                    img_izq = img_pil.crop((0, 0, int(w * 0.52), h))
                    img_der = img_pil.crop((int(w * 0.48), 0, w, h))
                    with ThreadPoolExecutor(max_workers=2) as ex:
                        f_izq = ex.submit(llamar_gemini, client, img_izq, modelo_izq, instrucciones_extra)
                        f_der = ex.submit(llamar_gemini, client, img_der, modelo_der, instrucciones_extra)
                        datos_izq, datos_der = f_izq.result(), f_der.result()
                else:
                    datos_izq = llamar_gemini(client, img_pil, modelo_izq, instrucciones_extra)
                    datos_der = {}

                alumno_raw = datos_izq.get("nombre_alumno", "").strip()
                if not alumno_raw or alumno_raw.lower() in ("desconocido", "en_blanco", "ilegible", "sin nombre"):
                    nombre_sin_ext = os.path.splitext(nombre_orig)[0]
                    alumno = f"SIN NOMBRE ({nombre_sin_ext})"
                else:
                    alumno = alumno_raw

                base = f"{nombre_actividad}_{limpiar_nombre(alumno)}"
                candidato = base
                c = 2
                while candidato in nombres_usados:
                    candidato = f"{base}_{c}"
                    c += 1
                nombres_usados.add(candidato)

                nom_img = f"{candidato}{ext}"
                zf.writestr(f"{nombre_actividad}/Estudiantes/{nom_img}", img_bytes)
                enlaces_consolidado[nom_img] = f"Estudiantes/{nom_img}"

                tachones = set(datos_izq.get("items_con_tachon", [])) | set(datos_der.get("items_con_tachon", []))
                respuestas = {**datos_izq, **datos_der}

                puntaje, cant_err, cant_rev = 0.0, 0, 0
                errs_est, trans_est = [], []

                for cid, (etq, resp_corr, tipo, peso_item) in clave_oficial.items():
                    resp_raw = respuestas.get(cid, "EN_BLANCO")
                    tuvo_tachon = cid in tachones
                    pts, est, obs = evaluar_item(
                        resp_raw, resp_corr, tipo, peso_item,
                        tuvo_tachon, factor_orto, umbral_orto,
                        doble_marca_revisar, importan_tildes,
                    )
                    puntaje += pts
                    if est == "Incorrecto":
                        cant_err += 1
                    elif est in ("Revisar", "Ortografía"):
                        cant_rev += 1

                    if est != "Correcto" or tuvo_tachon:
                        errs_est.append({
                            "Alumno": alumno, "Ejercicio / Ítem": etq, "Tipo": tipo,
                            "Estado": est, "Respuesta del Alumno": resp_raw,
                            "Respuesta Correcta": resp_corr, "Observación": obs, "Archivo": nom_img,
                        })
                    trans_est.append({
                        "Alumno": alumno, "Ejercicio / Ítem": etq, "Tipo": tipo,
                        "Leído por IA": resp_raw, "Clave Oficial": resp_corr,
                        "Puntos": pts, "Valor Ítem": peso_item, "Estado": est, "Observación": obs,
                    })

                puntaje = round(puntaje, 2)
                if escala_valor is None:
                    nota_final = puntaje
                else:
                    nota_final = round((puntaje / total_puntos_posibles) * escala_valor, 2)

                res_est = {
                    "Alumno": alumno,
                    "Puntaje": puntaje,
                    "Total Pts": round(total_puntos_posibles, 2),
                    col_nota_nombre: nota_final,
                    "Errores": cant_err,
                    "Por Revisar (Amarillo)": cant_rev,
                    "Tachones": len(tachones),
                    "Archivo": nom_img,
                }

                excel_ind = generar_excel_bytes([res_est], errs_est, trans_est, enlaces={nom_img: nom_img})
                zf.writestr(f"{nombre_actividad}/Estudiantes/{candidato}.xlsx", excel_ind)

                lista_resumen.append(res_est)
                lista_errores.extend(errs_est)
                lista_trans.extend(trans_est)

            except Exception as e:
                fallidos.append((nombre_orig, str(e)))

            if idx < len(paginas_a_procesar) - 1 and pausa_seg > 0:
                time.sleep(pausa_seg)

        if lista_resumen:
            excel_cons = generar_excel_bytes(
                lista_resumen, lista_errores, lista_trans,
                enlaces=enlaces_consolidado, incluir_formulas=True
            )
            zf.writestr(f"{nombre_actividad}/{nombre_actividad}_Consolidado.xlsx", excel_cons)

            ordenados = sorted(lista_resumen, key=lambda x: x["Alumno"].lower())
            buf_notas = BytesIO()
            with pd.ExcelWriter(buf_notas, engine="openpyxl") as writer:
                pd.DataFrame([{"Alumno": r["Alumno"], col_nota_nombre: r[col_nota_nombre]} for r in ordenados]).to_excel(
                    writer, sheet_name="Notas", index=False
                )
                ws_n = writer.sheets["Notas"]
                agregar_formulas(ws_n, len(ordenados), "B")
                dar_formato_hoja(ws_n, "resumen")
            zf.writestr(f"{nombre_actividad}/{nombre_actividad}_Solo_Notas.xlsx", buf_notas.getvalue())

    barra.progress(1.0, text="¡Calificación terminada!")

    if fallidos:
        st.warning(f"No se pudieron procesar {len(fallidos)} hoja(s): " + ", ".join(f[0] for f in fallidos))

    if lista_resumen:
        st.success(f"Se calificaron {len(lista_resumen)} hojas correctamente.")
        st.dataframe(pd.DataFrame(lista_resumen), width="stretch")

        st.download_button(
            label=f"📦 Descargar carpeta completa ({nombre_actividad}.zip)",
            data=zip_buffer.getvalue(),
            file_name=f"{nombre_actividad}.zip",
            mime="application/zip",
            type="primary",
        )
