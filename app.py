import io
import os
import re
import time
import sqlite3
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
from difflib import SequenceMatcher
from sklearn.ensemble import IsolationForest

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stApp [data-testid="stToolbar"] {display: none;}
    .viewerBadge_container__1QSob {display: none !important;}
    .viewerBadge_link__1S137 {display: none !important;}
    div[data-testid="stDecoration"] {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# pdfplumber es opcional: si no está instalado, la app sigue funcionando con Excel.
# Solo se exigirá cuando el usuario suba un .pdf en Datos Base.
try:
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # type: ignore

# groq es opcional: si no está instalado, el chatbot HITL queda deshabilitado
# pero el resto de la app sigue funcionando con normalidad.
try:
    from groq import Groq  # type: ignore
except ImportError:
    Groq = None  # type: ignore


def _obtener_api_key_groq() -> str:
    """Resuelve la API Key de Groq desde múltiples fuentes, en este orden:

    1. ``st.secrets["GROQ_API_KEY"]``  → Producción (Streamlit Community Cloud).
       Se configura en *Settings → Secrets* del dashboard de Streamlit.
    2. ``os.environ["GROQ_API_KEY"]``  → Desarrollo local (variable de entorno
       o archivo ``.env`` cargado previamente).
    3. Cadena vacía                    → El usuario deberá ingresarla a mano
       en la pestaña *Configuración*.

    Acceder a ``st.secrets`` cuando no existe ``.streamlit/secrets.toml`` lanza
    ``StreamlitSecretNotFoundError``; por eso envolvemos todo en try/except y
    nunca propagamos errores: la app debe arrancar siempre, con o sin secrets.
    """
    try:
        valor = st.secrets["GROQ_API_KEY"]  # type: ignore[index]
        if valor:
            return str(valor).strip()
    except Exception:
        pass
    return (os.environ.get("GROQ_API_KEY") or "").strip()

# plotly es opcional: si no está instalado, las visualizaciones avanzadas
# (grafo NLP + scatter 3D ML) muestran un aviso amigable, pero la app
# sigue funcionando. Recomendado: pip install plotly
try:
    import plotly.express as px  # type: ignore
    import plotly.graph_objects as go  # type: ignore
    PLOTLY_OK = True
except ImportError:
    px = None  # type: ignore
    go = None  # type: ignore
    PLOTLY_OK = False

# streamlit-option-menu para una navegación lateral profesional con iconos
# de Bootstrap. Si no está instalada, caemos a un selectbox nativo (zero CSS).
# Recomendado: pip install streamlit-option-menu
try:
    from streamlit_option_menu import option_menu  # type: ignore
    OPTION_MENU_OK = True
except ImportError:
    option_menu = None  # type: ignore
    OPTION_MENU_OK = False

# ============================================================
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="LicitAI · Dashboard de Análisis de Costos",
    page_icon=":material/analytics:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CSS MÍNIMO  ·  Solo oculta el chrome nativo de Streamlit
# ------------------------------------------------------------
# Toda la paleta corporativa vive en .streamlit/config.toml.
# Aquí NO forzamos colores, fondos, ni bordes — el tema nativo
# se encarga de eso para evitar problemas de contraste (texto
# blanco sobre fondo blanco) y mantener la coherencia.
# ============================================================
st.markdown(
    """
    <style>
        /* Oculta el header / footer / menú nativo de Streamlit que
           introduce ruido visual en una app empresarial. */
        #MainMenu, footer { visibility: hidden; }
        header[data-testid="stHeader"] { background: transparent; }

        /* Limita el ancho útil del bloque principal a un layout
           empresarial (similar a un dashboard SaaS B2B). */
        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 3rem !important;
            max-width: 1500px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# (CSS HISTÓRICO ELIMINADO)
# El bloque <style> de ~480 líneas que vivía aquí causaba
# problemas de contraste y forzaba colores rotos sobre el tema
# claro definido en config.toml. Lo reemplazamos por componentes
# nativos: st.metric, st.container(border=True), st.divider(),
# st.subheader, st.success/info/warning/error y Material Icons
# (`:material/...:` en labels) + streamlit-option-menu en la sidebar.
# ============================================================
# (Bloque CSS legacy de ~480 líneas eliminado intencionalmente.)


# ============================================================
# ESTANDARIZACIÓN DE UNIDADES DE MEDIDA  (S10 + PDF contratistas)
# ------------------------------------------------------------
# Se aplica a la columna "Unidad" de TODO DataFrame justo después
# de la extracción (Excel o PDF) y ANTES del merge / NLP / ML.
# Garantiza que "KG", "kg", "Kg ", "KG." colapsen a "kg", y que
# "m³", "M3", "mt3" colapsen a "m3", etc. Sin esto, el cruce
# semántico puede fallar solo por diferencias tipográficas.
# ============================================================
UNIDAD_CANONICA_MAP: dict[str, str] = {
    # --- Mano de obra / tiempo ---
    "hh": "hh", "h.h": "hh", "h-h": "hh",
    "hm": "hm", "h.m": "hm", "h-m": "hm",
    "he": "he", "h.e": "he",
    "d/e": "d/e",
    "dia": "dia", "día": "dia", "d": "dia", "dias": "dia", "días": "dia",
    "mes": "mes", "meses": "mes",
    "hra": "hra", "hrs": "hra", "hora": "hra", "horas": "hra", "h": "hra",
    # --- Longitud ---
    "m": "m", "ml": "m", "m.l": "m", "mtl": "m", "metro": "m", "metros": "m",
    "cm": "cm", "mm": "mm", "km": "km",
    "pulg": "pulg", "in": "pulg", "\"": "pulg",
    # --- Área / volumen ---
    "m2": "m2", "m²": "m2", "mt2": "m2", "metro2": "m2",
    "m3": "m3", "m³": "m3", "mt3": "m3", "metro3": "m3",
    "p2": "p2", "pie2": "p2", "pt": "p2", "pt2": "p2",
    # --- Peso ---
    "kg": "kg", "kilo": "kg", "kilos": "kg", "kgs": "kg",
    "tn": "tn", "ton": "tn", "t": "tn", "tonelada": "tn", "toneladas": "tn",
    "gr": "gr", "g": "gr", "gramo": "gr", "gramos": "gr",
    # --- Volumen líquido ---
    "gal": "gal", "galon": "gal", "galón": "gal", "galones": "gal",
    "lt": "lt", "l": "lt", "litro": "lt", "litros": "lt",
    # --- Envases / piezas (S10 canónico → "u") ---
    "u": "u", "und": "u", "uni": "u", "unid": "u",
    "unidad": "u", "unidades": "u",
    "pza": "u", "pieza": "u", "piezas": "u",
    # --- Otros envases ---
    "bls": "bls", "bolsa": "bls", "bolsas": "bls",
    "jgo": "jgo", "juego": "jgo", "juegos": "jgo",
    "par": "par", "pares": "par",
    "rll": "rll", "rollo": "rll", "rollos": "rll",
    "cj": "cj", "caja": "cj", "cajas": "cj",
    # --- Globales / porcentajes ---
    "glb": "glb", "global": "glb",
    "%mo": "%mo", "%m.o": "%mo", "%mo.": "%mo",
    "%eq": "%eq", "%equipos": "%eq",
    "est": "est", "estudio": "est",
}


def estandarizar_unidad(unidad) -> str:
    """Canonicaliza una unidad de medida S10.

    - Trims, baja a minúsculas y elimina espacios/tabs internos.
    - Mapea superíndices (`m²` → `m2`, `m³` → `m3`).
    - Resuelve la forma canónica vía `UNIDAD_CANONICA_MAP`.
      Si la unidad no está en el catálogo, devuelve su forma limpia.
    - Celdas vacías / NaN / 'None' devuelven `""` (no rompe el merge).
    """
    if unidad is None:
        return ""
    try:
        if isinstance(unidad, float) and pd.isna(unidad):
            return ""
    except Exception:
        pass
    s = str(unidad).strip().lower()
    if not s or s in {"nan", "none", "null", "-"}:
        return ""
    s = re.sub(r"\s+", "", s)
    s = s.replace("²", "2").replace("³", "3")
    s = s.rstrip(".")
    return UNIDAD_CANONICA_MAP.get(s, s)


def estandarizar_unidades_df(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica `estandarizar_unidad` a la columna 'Unidad' si existe."""
    if df is None or "Unidad" not in df.columns:
        return df
    df["Unidad"] = df["Unidad"].apply(estandarizar_unidad)
    return df


# ============================================================
# LÓGICA DE LIMPIEZA  (sin tocar — basada en prueba_excel.py)
# ============================================================
@st.cache_data(show_spinner="Procesando archivo de S10...")
def cargar_datos_limpios(archivo) -> pd.DataFrame:
    """Lee un Excel de S10 (ruta, BytesIO o UploadedFile de Streamlit)
    y devuelve un DataFrame limpio listo para análisis."""
    df = pd.read_excel(archivo, header=None)

    header_row_idx = -1
    for idx, row in df.iterrows():
        texto_fila = " ".join([str(x).lower() for x in row])
        if "precio" in texto_fila and ("cod" in texto_fila or "cód" in texto_fila):
            header_row_idx = idx
            break

    if header_row_idx == -1:
        raise ValueError("No se encontró la cabecera en el Excel.")

    cabecera = df.iloc[header_row_idx]

    def obtener_indice(palabras_clave):
        for i, valor in cabecera.items():
            valor_str = str(valor).lower()
            for palabra in palabras_clave:
                if palabra in valor_str:
                    return i
        return -1

    idx_cod = obtener_indice(["cod", "cód"])
    idx_desc = obtener_indice(["desc", "recurso"])
    idx_und = obtener_indice(["und", "unidad"])
    idx_cuad = obtener_indice(["cuadrilla"])
    idx_cant = obtener_indice(["cant"])
    idx_prec = obtener_indice(["precio", "unit"])
    idx_parc = obtener_indice(["parcial", "total"])

    df_data = df.iloc[header_row_idx + 1:].copy()

    df_final = pd.DataFrame()
    df_final["Código"] = df_data[idx_cod] if idx_cod != -1 else None
    df_final["Descripción"] = df_data[idx_desc] if idx_desc != -1 else None
    df_final["Unidad"] = df_data[idx_und] if idx_und != -1 else None
    df_final["Cuadrilla"] = df_data[idx_cuad] if idx_cuad != -1 else None
    df_final["Cantidad"] = df_data[idx_cant] if idx_cant != -1 else None
    df_final["Precio"] = df_data[idx_prec] if idx_prec != -1 else None
    df_final["Parcial"] = df_data[idx_parc] if idx_parc != -1 else None

    df_final = df_final.replace(r"^\s*$", np.nan, regex=True)
    df_final = df_final.dropna(subset=["Descripción"])
    df_final = df_final.dropna(subset=["Precio"])

    for col in ["Cantidad", "Precio", "Parcial"]:
        df_final[col] = (
            df_final[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(r"[^\d\.]", "", regex=True)
        )
        df_final[col] = df_final[col].replace("", np.nan)
        df_final[col] = pd.to_numeric(df_final[col], errors="coerce")

    df_final = df_final[
        ~df_final["Código"]
        .astype(str)
        .str.lower()
        .str.contains("cód|cod|rendimiento|partida", na=False)
    ]
    df_final = df_final.dropna(subset=["Precio"])

    df_final = estandarizar_unidades_df(df_final)

    return df_final.reset_index(drop=True)


# ============================================================
# INGESTA DE PDF  ·  pdfplumber + cleaning robusto
# ============================================================
# Listas de palabras clave para mapear cabeceras heterogéneas de los PDFs
# de contratistas (cada empresa nombra las columnas distinto).
_PDF_HEADERS = {
    "descripcion": ["descripcion", "descripción", "partida", "concepto", "recurso", "detalle", "item"],
    "unidad":      ["und", "unidad", "u.m.", "u/m", "u.med"],
    "cantidad":    ["cant", "cantidad", "metrado", "qty"],
    "precio":      ["p.u.", "p/u", "precio unit", "precio uni", "precio", "valor unit", "costo unit", "p. unit"],
    "parcial":     ["parcial", "subtotal", "importe", "total"],
}


def _pdf_normalizar(txt) -> str:
    return str(txt or "").strip().lower()


def _pdf_es_header(fila) -> bool:
    """Detecta si una fila de tabla parece ser la cabecera (tiene 'descripción' + 'precio')."""
    texto = " | ".join(_pdf_normalizar(c) for c in fila)
    tiene_desc = any(p in texto for p in ["descripc", "partida", "concepto", "recurso", "detalle"])
    tiene_precio = any(p in texto for p in ["precio", "p.u", "p/u", "p. u", "valor unit", "costo unit"])
    return tiene_desc and tiene_precio


def _pdf_mapear_columnas(fila_header) -> dict:
    """Devuelve {clave_logica: indice_columna} a partir de una fila-cabecera."""
    idx_map = {k: -1 for k in _PDF_HEADERS}
    for i, cell in enumerate(fila_header):
        cell_n = _pdf_normalizar(cell)
        if not cell_n:
            continue
        for key, palabras in _PDF_HEADERS.items():
            if idx_map[key] != -1:
                continue
            for p in palabras:
                if p in cell_n:
                    idx_map[key] = i
                    break
    return idx_map


def _pdf_parsear_numero(valor):
    """Convierte celdas tipo 'S/. 1,234.56' o '1.234,56' o '12,50' a float."""
    if valor is None:
        return np.nan
    s = str(valor).strip()
    if not s or s.lower() in {"nan", "none", "-"}:
        return np.nan
    s = s.replace("S/.", "").replace("S/", "").replace("$", "").strip()
    # Si tiene tanto '.' como ',' decidimos cuál es decimal por la última posición
    if "," in s and "." in s:
        if s.rfind(".") > s.rfind(","):
            s = s.replace(",", "")              # formato 1,234.56 → 1234.56
        else:
            s = s.replace(".", "").replace(",", ".")  # formato 1.234,56 → 1234.56
    else:
        s = s.replace(",", ".")
    s = "".join(c for c in s if c.isdigit() or c in ".-")
    if s in ("", ".", "-", "-."):
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


# ---------- Estrategia B · Parseo de texto plano (S10 APU) ----------
# Catálogo de unidades típicas de construcción civil en Perú / S10.
# Ordenado por longitud DESC para que "m3" matchee antes que "m".
_UNIDADES_S10_LISTA = sorted(
    [
        "hh", "hm", "h-h", "h-m", "h.h", "h.m",
        "kg", "gr", "g", "tn", "ton", "sac", "bls", "bol",
        "m3", "m³", "m2", "m²", "ml", "p2", "m",
        "und", "unid", "u", "pza", "pieza", "jgo", "jgo.",
        "gal", "gln", "lt", "ltr", "l",
        "%mo", "%hh", "%pu", "%eq",
        "mes", "dia", "día", "d/e",
        "glb", "est", "par", "rll",
    ],
    key=len,
    reverse=True,
)
_UNIDADES_S10_RE = "|".join(re.escape(u) for u in _UNIDADES_S10_LISTA)
_NUM_RE = r"-?\d[\d,\.]*"

# Regex maestro para una línea-partida estilo S10 APU:
#   [código opcional]  Descripción …   Unidad   N1 [N2 [N3 [N4]]]
PATRON_LINEA_S10 = re.compile(
    r"^\s*"
    r"(?:(?P<codigo>\d{2,}(?:[\.\-]\d+)*)\s+)?"          # código S10 opcional
    r"(?P<descripcion>\S.*?)\s+"                          # descripción (no greedy)
    r"(?P<unidad>(?:" + _UNIDADES_S10_RE + r"))\b\s+"     # unidad de catálogo
    r"(?P<numeros>(?:" + _NUM_RE + r"\s+){0,3}"
    + _NUM_RE + r")\s*$",                                  # 1-4 números al final
    re.IGNORECASE,
)

# Líneas a descartar siempre (ruido típico del reporte S10)
_LINEAS_RUIDO = [
    re.compile(r"an[áa]lisis\s+de\s+precios\s+unitarios", re.IGNORECASE),
    re.compile(r"^\s*subpresupuesto\b", re.IGNORECASE),
    re.compile(r"^\s*partida\s+\d", re.IGNORECASE),
    re.compile(r"^\s*rendimiento\b", re.IGNORECASE),
    re.compile(r"costo\s+unitario\s+directo", re.IGNORECASE),
    re.compile(r"^\s*p[áa]gina\s*[:\s]*\d+", re.IGNORECASE),
    re.compile(r"^\s*fecha\s*[:\s]", re.IGNORECASE),
    re.compile(r"^\s*hora\s*[:\s]", re.IGNORECASE),
    re.compile(r"^\s*(presupuesto|cliente|lugar|propietario)\s*[:\s]", re.IGNORECASE),
    re.compile(r"^\s*\d{1,2}/\d{1,2}/\d{2,4}\s*$"),       # solo una fecha
    re.compile(r"^\s*-{3,}\s*$"),                         # separadores
    re.compile(r"^\s*={3,}\s*$"),
]

_SECCION_TITULO = re.compile(
    r"^\s*(mano\s+de\s+obra|materiales|equipos|equipo|herramientas|"
    r"subcontratos?|otros|varios)\s*:?\s*$",
    re.IGNORECASE,
)


def _es_linea_ruido(linea: str) -> bool:
    if not linea or not linea.strip():
        return True
    for patron in _LINEAS_RUIDO:
        if patron.search(linea):
            return True
    if _SECCION_TITULO.match(linea):
        return True
    return False


def _extraer_filas_tablas(pdf) -> list[dict]:
    """Estrategia A · Extracción a partir de tablas estructuradas (find_tables)."""
    filas: list[dict] = []
    idx_columnas: dict | None = None

    for pagina in pdf.pages:
        tablas = pagina.extract_tables() or []
        for tabla in tablas:
            for fila in tabla:
                if not fila or all(c is None or str(c).strip() == "" for c in fila):
                    continue

                # Cabecera puede repetirse en cada página → re-mapeamos
                if _pdf_es_header(fila):
                    idx_columnas = _pdf_mapear_columnas(fila)
                    continue
                if idx_columnas is None:
                    continue

                def _get(key):
                    i = idx_columnas[key]
                    if i < 0 or i >= len(fila):
                        return None
                    return fila[i]

                desc = _get("descripcion")
                if not desc or not str(desc).strip():
                    continue
                desc_str = " ".join(str(desc).split())

                precio = _pdf_parsear_numero(_get("precio"))
                if pd.isna(precio) or precio <= 0:
                    continue

                cant_raw = _pdf_parsear_numero(_get("cantidad"))
                cant = cant_raw if not pd.isna(cant_raw) else 1.0

                parcial_raw = _pdf_parsear_numero(_get("parcial"))
                parcial = parcial_raw if not pd.isna(parcial_raw) else cant * precio

                unidad = str(_get("unidad") or "").strip()

                filas.append(
                    {
                        "Código":      "",       # las tablas no suelen traer código S10
                        "Descripción": desc_str,
                        "Unidad":      unidad,
                        "Cantidad":    cant,
                        "Precio":      precio,
                        "Parcial":     parcial,
                    }
                )
    return filas


def _extraer_filas_texto(pdf) -> list[dict]:
    """Estrategia B · Parseo línea-a-línea con regex (PDFs S10 APU sin cuadrículas)."""
    filas: list[dict] = []

    for pagina in pdf.pages:
        texto = pagina.extract_text() or ""
        for raw_linea in texto.splitlines():
            # Normalizamos espacios múltiples (los PDFs de S10 usan tabuladores anchos)
            linea = re.sub(r"[ \t]+", " ", raw_linea).strip()
            if _es_linea_ruido(linea):
                continue

            match = PATRON_LINEA_S10.match(linea)
            if not match:
                continue

            codigo = (match.group("codigo") or "").strip()
            descripcion = " ".join(match.group("descripcion").split())
            unidad = match.group("unidad").strip().lower()
            numeros_raw = match.group("numeros")

            # Una descripción válida suele tener al menos 3 chars (filtra basura)
            if len(descripcion) < 3:
                continue

            nums = [_pdf_parsear_numero(n) for n in numeros_raw.split()]
            nums = [n for n in nums if not pd.isna(n)]
            if not nums:
                continue

            # Mapeo según cuántos números aparecen al final de la línea:
            #   1 → solo precio                    (cantidad=1, parcial=precio)
            #   2 → cantidad, precio               (parcial=cant*precio)
            #   3+ → ..., cantidad, precio, parcial (tomamos los últimos 3)
            if len(nums) == 1:
                cantidad, precio, parcial = 1.0, nums[0], nums[0]
            elif len(nums) == 2:
                cantidad, precio = nums
                parcial = cantidad * precio
            else:
                cantidad, precio, parcial = nums[-3], nums[-2], nums[-1]

            if precio <= 0:
                continue

            filas.append(
                {
                    "Código":      codigo,        # vacío si el PDF no traía código
                    "Descripción": descripcion,
                    "Unidad":      unidad,
                    "Cantidad":    cantidad,
                    "Precio":      precio,
                    "Parcial":     parcial,
                }
            )
    return filas


@st.cache_data(show_spinner="Extrayendo datos del PDF (estrategia híbrida)...")
def extraer_datos_pdf(archivo) -> pd.DataFrame:
    """Extrae partidas de un PDF de presupuesto con estrategia HÍBRIDA.

    Estrategia A · Tablas estructuradas
        Usa `pdfplumber.extract_tables()`. Funciona bien con Excels exportados
        a PDF, presupuestos con cuadrículas/bordes y tablas formateadas.

    Estrategia B · Texto plano + Regex
        Si la Estrategia A devuelve <3 filas, se extrae el texto bruto de cada
        página y se aplica un regex robusto para detectar líneas-partida estilo
        S10 APU: `[código] descripción ... unidad números`. Las unidades se
        validan contra un catálogo de construcción civil (hh, hm, kg, m2, m3,
        und, pza, gal, %mo, …) y se filtran encabezados, fechas y títulos.

    Política de Códigos
        - Si la línea trae un código S10 detectado por el regex, se conserva
          (potencialmente cruzará por código exacto).
        - Si NO trae código (típico en propuestas de contratistas),
          se rellena con `PDF-NNNN` único, lo que **fuerza** que el cruce
          falle por código y entre en acción el módulo NLP semántico.

    Devuelve un DataFrame con el schema canónico:
        Código · Descripción · Unidad · Cuadrilla · Cantidad · Precio · Parcial
    """
    if pdfplumber is None:
        raise ImportError(
            "pdfplumber no está instalado. Agrega `pdfplumber` a requirements.txt "
            "y ejecuta: pip install pdfplumber"
        )

    with pdfplumber.open(archivo) as pdf:
        # --- Estrategia A: tablas estructuradas ---
        filas_a = _extraer_filas_tablas(pdf)

        # --- Estrategia B: fallback a texto plano si A devolvió poco ---
        # Si A da resultados decentes, usamos A. Si no, probamos B y nos
        # quedamos con la que más filas devuelva.
        if len(filas_a) >= 3:
            filas = filas_a
        else:
            filas_b = _extraer_filas_texto(pdf)
            filas = filas_b if len(filas_b) > len(filas_a) else filas_a

    if not filas:
        raise ValueError(
            "No se pudieron extraer partidas del PDF. "
            "Verifica que el PDF contenga texto seleccionable (no sea una imagen escaneada) "
            "y que respete el formato de presupuesto S10 o tabla con columnas "
            "'Descripción' / 'Precio Unitario'."
        )

    df = pd.DataFrame(filas)

    # --- Política de Código: respeta los reales, sintetiza los vacíos ---
    # Empty-string + drop_duplicates(subset=["Código"]) en comparar_presupuestos
    # colapsaría todas las filas vacías a una sola → mal. Así que rellenamos
    # los huecos con PDF-NNNN únicos (siguen sin matchear códigos S10 → NLP).
    codigos = df["Código"].astype(str).str.strip()
    contador = 1
    nuevos_cod: list[str] = []
    for c in codigos:
        if not c or c.lower() == "nan":
            nuevos_cod.append(f"PDF-{contador:04d}")
            contador += 1
        else:
            nuevos_cod.append(c)
    df["Código"] = nuevos_cod

    # Mantener schema canónico (mismo que cargar_datos_limpios)
    if "Cuadrilla" not in df.columns:
        df.insert(3, "Cuadrilla", np.nan)
    df = df[["Código", "Descripción", "Unidad", "Cuadrilla", "Cantidad", "Precio", "Parcial"]]

    df = estandarizar_unidades_df(df)

    return df.reset_index(drop=True)


def cargar_presupuesto(archivo) -> pd.DataFrame:
    """Dispatcher de ingesta: enruta al parser según la extensión del archivo.

    - `.xls` / `.xlsx` → `cargar_datos_limpios()` (S10 estándar)
    - `.pdf`           → `extraer_datos_pdf()`   (extracción con pdfplumber)

    Ambas rutas devuelven un DataFrame con el mismo schema, por lo que el resto
    del pipeline (`comparar_presupuestos`, NLP, ML) funciona sin cambios.
    """
    nombre = (getattr(archivo, "name", None) or str(archivo)).lower()
    if nombre.endswith(".pdf"):
        return extraer_datos_pdf(archivo)
    return cargar_datos_limpios(archivo)


# ============================================================
# MOTOR DE INTELIGENCIA ARTIFICIAL (DÍA 2)
# ============================================================
@st.cache_data(show_spinner="Entrenando modelo de detección de anomalías...")
def clasificar_y_detectar_anomalias(df):
    df_ml = df.copy()
    
    # 1. Categorización basada en reglas simples (NLP básico)
    def categorizar(unidad):
        u = str(unidad).lower().strip()
        if u in ['hh', 'mes', 'dia']: return 'Mano de Obra'
        elif u in ['hm', 'he', 'd/e']: return 'Equipos'
        elif u in ['%mo', 'glb']: return 'Subcontratos/Varios'
        else: return 'Materiales'
        
    df_ml['Categoría'] = df_ml['Unidad'].apply(categorizar)
    
    # 2. Detección de Anomalías con Isolation Forest
    # Usaremos Precio Unitario, Cantidad y el Parcial para buscar patrones raros
    features = ['Cantidad', 'Precio', 'Parcial']
    
    # Llenamos vacíos con 0 por si acaso para que el modelo no crashee
    X = df_ml[features].fillna(0)
    
    # Configuramos el modelo asumiendo que el 5% del presupuesto podría tener errores/anomalías
    modelo_if = IsolationForest(contamination=0.05, random_state=42)
    
    # -1 significa Anomalía (Rojo), 1 significa Normal (Verde)
    df_ml['Es_Anomalia'] = modelo_if.fit_predict(X)
    
    # Calculamos un "Score" de rareza (mientras más negativo, más raro es)
    scores = modelo_if.decision_function(X)
    # Lo normalizamos a un % de "Nivel de Riesgo" (0% a 100%)
    df_ml['Nivel_Riesgo'] = np.interp(scores, (scores.min(), scores.max()), (100, 0))
    
    return df_ml


# ============================================================
# CATEGORIZADOR  ·  Clasifica cada partida por tipo de recurso
# ============================================================
def categorizar_unidad(unidad) -> str:
    """Clasifica una unidad de medida S10 en su Categoría de recurso."""
    u = str(unidad).lower().strip()
    if u in {"hh", "mes", "dia", "día", "d"}:
        return "Mano de Obra"
    if u in {"hm", "he", "d/e", "h.m", "h.e", "hra"}:
        return "Equipos"
    if u in {"%mo", "glb"}:
        return "Subcontratos / Varios"
    return "Materiales"


# ============================================================
# MOTOR NLP  ·  Similitud semántica de descripciones
# ============================================================
def comparar_semantica(base_desc: str, postor_desc: str) -> float:
    """Devuelve un score de similitud entre 0 y 100 entre dos descripciones.

    Usa difflib.SequenceMatcher (algoritmo de Ratcliff-Obershelp).
    Normaliza el texto (lower + strip + colapsar espacios) para que pequeñas
    diferencias de formato no penalicen el match.
    """
    a = " ".join(str(base_desc or "").lower().split())
    b = " ".join(str(postor_desc or "").lower().split())
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio() * 100.0


# ============================================================
# COMPARADOR DE PRESUPUESTOS  (DÍA 4 · NLP + ML)
# ============================================================
@st.cache_data(show_spinner="Cruzando presupuestos · NLP + Isolation Forest...")
def comparar_presupuestos(
    df_base: pd.DataFrame,
    df_contratista: pd.DataFrame,
    umbral_nlp: float = 80.0,
) -> pd.DataFrame:
    """Une dos presupuestos en dos pasadas y enriquece con un score de anomalía.

    Pasada 1 (cruce exacto):
        pd.merge por 'Código' → método 'Código S10'.

    Pasada 2 (cruce semántico):
        Para cada fila del Postor que NO encontró match exacto, busca la
        Descripción más parecida del Base (similitud > `umbral_nlp`) usando
        difflib.SequenceMatcher → método 'IA Semántica (NLP)'.

    Pasada 3 (ML):
        Aplica IsolationForest sobre [Precio_Base, Precio_Contratista,
        Sobrecosto, Desviación] y guarda 'Score de Anomalía' (0-100, 100 = más raro).

    Columnas de salida:
        Código · Descripción · Unidad · Cantidad
        Precio_Base · Precio_Contratista · Parcial_Base · Parcial_Contratista
        Sobrecosto (S/.) · Desviación (%)
        Método de Cruce · Score Similitud (%) · Score de Anomalía
    """
    # --- 0. Dedupe en ambos lados para evitar productos cartesianos ---
    base_unico = df_base.drop_duplicates(subset=["Código"], keep="first").copy()
    contra_unico = df_contratista.drop_duplicates(subset=["Código"], keep="first").copy()

    base_prep = base_unico[
        ["Código", "Descripción", "Unidad", "Cantidad", "Precio", "Parcial"]
    ].rename(columns={"Precio": "Precio_Base", "Parcial": "Parcial_Base"})

    contra_prep = contra_unico[["Código", "Descripción", "Precio", "Parcial"]].rename(
        columns={
            "Descripción": "Desc_Contratista",
            "Precio": "Precio_Contratista",
            "Parcial": "Parcial_Contratista",
        }
    )

    # --- 1. Cruce EXACTO por código S10 (1:1 garantizado) ---
    df_codigo = pd.merge(
        base_prep, contra_prep, on="Código", how="inner", validate="one_to_one"
    )
    df_codigo["Método de Cruce"] = "Código S10"
    df_codigo["Score Similitud (%)"] = 100.0
    df_codigo = df_codigo.drop(columns=["Desc_Contratista"])

    # --- 2. Cruce SEMÁNTICO (NLP) para los que no encontraron match exacto ---
    cods_matched = set(df_codigo["Código"].astype(str))
    base_resto = base_prep[~base_prep["Código"].astype(str).isin(cods_matched)].copy()
    contra_resto = contra_prep[~contra_prep["Código"].astype(str).isin(cods_matched)].copy()

    rescued = []
    base_idx_used: set = set()

    if not contra_resto.empty and not base_resto.empty:
        for _, c_row in contra_resto.iterrows():
            desc_c = str(c_row["Desc_Contratista"])
            mejor_score = 0.0
            mejor_idx = None
            for b_idx, b_row in base_resto.iterrows():
                if b_idx in base_idx_used:
                    continue
                score = comparar_semantica(b_row["Descripción"], desc_c)
                if score > mejor_score:
                    mejor_score = score
                    mejor_idx = b_idx
            if mejor_idx is not None and mejor_score >= umbral_nlp:
                b = base_resto.loc[mejor_idx]
                base_idx_used.add(mejor_idx)
                rescued.append(
                    {
                        "Código": b["Código"],
                        "Descripción": b["Descripción"],
                        "Unidad": b["Unidad"],
                        "Cantidad": b["Cantidad"],
                        "Precio_Base": b["Precio_Base"],
                        "Parcial_Base": b["Parcial_Base"],
                        "Precio_Contratista": c_row["Precio_Contratista"],
                        "Parcial_Contratista": c_row["Parcial_Contratista"],
                        "Método de Cruce": "IA Semántica (NLP)",
                        "Score Similitud (%)": round(mejor_score, 1),
                    }
                )

    df_nlp = pd.DataFrame(rescued)
    df = pd.concat([df_codigo, df_nlp], ignore_index=True) if not df_nlp.empty else df_codigo

    # --- 3. Cálculos de sobrecosto y desviación ---
    df["Sobrecosto (S/.)"] = df["Precio_Contratista"] - df["Precio_Base"]
    df["Desviación (%)"] = np.where(
        df["Precio_Base"] > 0,
        (df["Sobrecosto (S/.)"] / df["Precio_Base"]) * 100,
        0.0,
    )

    # --- 4. Score de Anomalía (Isolation Forest) ---
    if len(df) >= 2:
        features = df[
            ["Precio_Base", "Precio_Contratista", "Sobrecosto (S/.)", "Desviación (%)"]
        ].fillna(0)
        modelo = IsolationForest(contamination=0.1, random_state=42)
        modelo.fit(features)
        # decision_function: valores ALTOS = normales, valores BAJOS = anómalos
        scores = modelo.decision_function(features)
        # Normalizamos a 0-100 invertido → 100 = anomalía total
        df["Score de Anomalía"] = np.interp(
            scores, (scores.min(), scores.max()), (100, 0)
        ).round(1)
    else:
        df["Score de Anomalía"] = 0.0

    # --- 5. Categorización por tipo de recurso ---
    df["Categoría"] = df["Unidad"].apply(categorizar_unidad)

    return df.reset_index(drop=True)


# ============================================================
# REPORTE DE AUDITORÍA  ·  Exportación a Excel
# ============================================================
def generar_reporte_excel(
    df_comparativo: pd.DataFrame,
    comentarios: str,
    auditor: str = "Auditor LicitAI",
) -> bytes:
    """Genera un Excel multi-hoja con el comparativo + bitácora de auditoría."""
    output = io.BytesIO()
    timestamp = datetime.now()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_comparativo.to_excel(writer, sheet_name="Comparativo", index=False)

        resumen = pd.DataFrame(
            [
                {"Campo": "Fecha de auditoría",          "Valor": timestamp.strftime("%d/%m/%Y %H:%M")},
                {"Campo": "Auditor responsable",         "Valor": auditor},
                {"Campo": "Total de partidas analizadas","Valor": len(df_comparativo)},
                {"Campo": "Cruces por Código S10",       "Valor": int((df_comparativo["Método de Cruce"] == "Código S10").sum())},
                {"Campo": "Cruces por IA Semántica",     "Valor": int((df_comparativo["Método de Cruce"] == "IA Semántica (NLP)").sum())},
                {"Campo": "Sobrecosto total (S/.)",      "Valor": float(df_comparativo["Sobrecosto (S/.)"].sum())},
                {"Campo": "Anomalías críticas (>70)",    "Valor": int((df_comparativo["Score de Anomalía"] > 70).sum())},
                {"Campo": "Comentarios del auditor",     "Valor": comentarios or "(sin comentarios)"},
            ]
        )
        resumen.to_excel(writer, sheet_name="Auditoría", index=False)

    return output.getvalue()


# ============================================================
# CAPA DE PERSISTENCIA  ·  Historial de auditorías (SQLite)
# ============================================================
DB_PATH = Path(__file__).parent / "historial_auditorias.db"


def init_db() -> None:
    """Crea la tabla 'auditorias' si no existe. Idempotente."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS auditorias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                total_partidas INTEGER NOT NULL,
                sobrecosto_total REAL NOT NULL,
                anomalias_detectadas INTEGER NOT NULL,
                cruces_codigo INTEGER NOT NULL,
                cruces_nlp INTEGER NOT NULL,
                comentario TEXT,
                auditor TEXT
            )
            """
        )
        conn.commit()


def guardar_auditoria(
    total_partidas: int,
    sobrecosto_total: float,
    anomalias_detectadas: int,
    cruces_codigo: int,
    cruces_nlp: int,
    comentario: str,
    auditor: str = "Auditor LicitAI",
) -> int:
    """Inserta un registro en el historial. Devuelve el id creado."""
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            """
            INSERT INTO auditorias
                (fecha, total_partidas, sobrecosto_total, anomalias_detectadas,
                 cruces_codigo, cruces_nlp, comentario, auditor)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                fecha,
                int(total_partidas),
                float(sobrecosto_total),
                int(anomalias_detectadas),
                int(cruces_codigo),
                int(cruces_nlp),
                comentario or "",
                auditor,
            ),
        )
        conn.commit()
        return int(cur.lastrowid)


def obtener_historial() -> pd.DataFrame:
    """Devuelve todo el historial ordenado del más reciente al más antiguo."""
    with sqlite3.connect(DB_PATH) as conn:
        df = pd.read_sql_query(
            """
            SELECT id, fecha, total_partidas, sobrecosto_total,
                   anomalias_detectadas, cruces_codigo, cruces_nlp,
                   auditor, comentario
            FROM auditorias
            ORDER BY id DESC
            """,
            conn,
        )
    return df


def borrar_historial() -> None:
    """Elimina TODOS los registros del historial. Útil para demos."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM auditorias")
        conn.commit()


# ============================================================
# ESTADO COMPARTIDO  ·  Carga lazy a partir de session_state
# ============================================================
def cargar_estado_actual():
    """Devuelve los DataFrames PERSISTIDOS en session_state.

    Los DataFrames se procesan UNA sola vez en `render_datos_base()` y se
    guardan en `st.session_state['df_base']`, `st.session_state['df_contratista']`
    y `st.session_state['df_compare']`. El resto de páginas solo los LEEN, lo
    que garantiza la persistencia entre pestañas (no se pierden al navegar).

    Si el usuario cambió el `umbral_nlp` desde el último cómputo, recalculamos
    automáticamente `df_compare` aquí (es barato gracias al @st.cache_data).

    Returns
    -------
    tuple
        (df_base, df_contra, df_compare, modo_comparativo, error_msg)
    """
    df_base    = st.session_state.get("df_base")
    df_contra  = st.session_state.get("df_contratista")
    df_compare = st.session_state.get("df_compare")
    error_msg  = None

    umbral_actual = st.session_state.get("umbral_nlp", 80)
    umbral_usado  = st.session_state.get("umbral_nlp_usado")

    # Si tenemos ambos DataFrames pero el comparativo está vacío o el umbral
    # cambió, recalculamos. cargar_datos_limpios y comparar_presupuestos están
    # cacheados, así que esta llamada es barata.
    if df_base is not None and df_contra is not None:
        if df_compare is None or umbral_usado != umbral_actual:
            try:
                df_compare = comparar_presupuestos(
                    df_base, df_contra, umbral_nlp=float(umbral_actual)
                )
                st.session_state.df_compare       = df_compare
                st.session_state.umbral_nlp_usado = umbral_actual
            except Exception as e:
                error_msg = f"Error generando el comparativo: {e}"
                st.session_state.df_compare = None
                df_compare = None

    modo_comparativo = df_compare is not None and not df_compare.empty
    return df_base, df_contra, df_compare, modo_comparativo, error_msg


def render_page_header(titulo: str, subtitulo: str, icono: str | None = None) -> None:
    """Cabecera consistente para cada página, 100% nativa.

    Usa `st.title()` + `st.caption()` + `st.divider()` para una estética
    Enterprise B2B sobria, sin HTML inyectado. El `icono` opcional es un
    Material Icon (ej: ":material/dashboard:") que precede al título.
    """
    if icono:
        st.title(f"{icono} {titulo}")
    else:
        st.title(titulo)
    st.caption(subtitulo)
    st.divider()


# ============================================================
# HELPERS DE UI  (100% componentes nativos de Streamlit)
# ============================================================
def metric_card(
    icon: str,
    label: str,
    value: str,
    delta: str | None = None,
    up: bool = True,
):
    """Renderiza una tarjeta métrica corporativa con `st.metric` envuelto en
    `st.container(border=True)`.

    - `icon` se inyecta como prefijo del label (acepta sintaxis nativa
      `:material/...:` para Material Icons, o cualquier carácter Unicode).
    - `delta` y `up` se traducen a `delta_color` semántico de st.metric:
      verde cuando `up=True`, rojo cuando `up=False`.
    """
    with st.container(border=True):
        # Material Icon o emoji corporativo opcional como prefijo del label
        full_label = f"{icon} {label}" if icon else label
        st.metric(
            label=full_label,
            value=value,
            delta=delta,
            delta_color=("normal" if up else "inverse"),
        )


# ============================================================
# HELPER GENERAL DE FORMATO
# ============================================================
def fmt_soles_corto(valor: float) -> str:
    """Formatea valores grandes como S/. 438K  /  S/. 2.1M."""
    if valor is None or pd.isna(valor):
        return "S/. 0"
    if abs(valor) >= 1_000_000:
        return f"S/. {valor/1_000_000:.1f}M"
    if abs(valor) >= 1_000:
        return f"S/. {valor/1_000:.0f}K"
    return f"S/. {valor:,.2f}"


# ============================================================
# CONFIGURACIÓN DE PÁGINAS  (router)
# ============================================================
# NAV_ITEMS: cada entrada es (nombre_página, icono_bootstrap, icono_material).
# - icono_bootstrap → usado por streamlit-option-menu en la sidebar.
# - icono_material → usado en el título de cada página (st.title).
# Lista de iconos Bootstrap: https://icons.getbootstrap.com/
# Lista de iconos Material:  https://fonts.google.com/icons
NAV_ITEMS = [
    ("Resumen",            "speedometer2",        ":material/dashboard:"),
    ("Análisis de Costos", "bar-chart-line",      ":material/analytics:"),
    ("Datos Base",         "folder",              ":material/folder:"),
    ("Anomalías",          "exclamation-triangle",":material/warning:"),
    ("Reportes",           "file-earmark-text",   ":material/description:"),
    ("Configuración",      "gear",                ":material/settings:"),
]
NAV_NAMES   = [item[0] for item in NAV_ITEMS]
NAV_BS_ICON = {item[0]: item[1] for item in NAV_ITEMS}
NAV_MD_ICON = {item[0]: item[2] for item in NAV_ITEMS}

# Inicialización de session_state (se hace UNA sola vez por sesión).
# Mantenemos los DataFrames procesados en session_state para que persistan
# entre pestañas (st.file_uploader NO garantiza la persistencia del archivo
# cuando su widget no está en el DOM, así que NO confiamos en él para leer).
_DEFAULTS = {
    "pagina_actual":       "Resumen",
    "umbral_nlp":          80,
    "df_base":             None,   # DataFrame limpio del Presupuesto Base
    "df_contratista":      None,   # DataFrame limpio de la Propuesta Contratista
    "df_compare":          None,   # Resultado de comparar_presupuestos()
    "umbral_nlp_usado":    None,   # Umbral con el que se calculó df_compare (para invalidar cache)
    "nombre_archivo_base": None,
    "nombre_archivo_contra": None,
    # --- Chatbot HITL con Groq + Llama 3 ---
    # API Key del usuario en una clave que NO está atada al ciclo de vida del widget,
    # para que sobreviva al cambiar de pestaña. El widget usa una key "temp_api_key"
    # y un on_change callback (guardar_api_key) que copia temp → persistente.
    #
    # Valor inicial: se autocompleta desde st.secrets["GROQ_API_KEY"] (Streamlit
    # Cloud) o desde la variable de entorno GROQ_API_KEY (entorno local). Si no
    # hay ninguna configurada, queda vacío y el usuario la ingresa manualmente.
    "api_key_persistente": _obtener_api_key_groq(),
    "messages":            [],     # Historial de mensajes del chat [{role, content}, ...]
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# Aseguramos que la BD del historial existe desde el primer render
init_db()


# ============================================================
# RENDERER · LEYENDA + COMPONENTES REUTILIZABLES
# ============================================================
def render_legend_anomalia() -> None:
    """Leyenda del Score de Anomalía como tarjeta nativa con bordes (sin HTML)."""
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns([0.34, 0.22, 0.22, 0.22])
        c1.markdown(
            ":material/straighten: **Cómo leer el Score de Anomalía** "
            "*(Isolation Forest)*"
        )
        c2.markdown(":green[●] **0–50** Normal")
        c3.markdown(":orange[●] **50–75** Precaución")
        c4.markdown(":red[●] **75–100** Crítico")

ORDEN_CATEGORIAS = ["Mano de Obra", "Materiales", "Equipos", "Subcontratos / Varios"]


def _column_config_grupo():
    """Config de columnas reutilizable para la tabla agrupada."""
    return {
        "Descripción": st.column_config.TextColumn(
            "Descripción de la partida",
            width="large",
        ),
        "Precio_Base": st.column_config.NumberColumn(
            "Precio Base",
            help="Precio unitario referencial del Presupuesto Base",
            format="S/. %.2f",
        ),
        "Precio_Contratista": st.column_config.NumberColumn(
            "Precio Contratista",
            help="Precio unitario propuesto por el Contratista",
            format="S/. %.2f",
        ),
        "Score de Anomalía": st.column_config.ProgressColumn(
            "Score de Anomalía",
            help="Isolation Forest · 0=normal, 100=anomalía total",
            format="%.1f",
            min_value=0,
            max_value=100,
        ),
    }


# ============================================================
# VISUALIZACIONES AVANZADAS DE IA  (Plotly · Tema oscuro)
# ============================================================

# Paleta dark-pro consistente con el resto del dashboard
PLOTLY_BG       = "#0f172a"   # fondo del card
PLOTLY_PAPER    = "#0f172a"
PLOTLY_GRID     = "#1e293b"
PLOTLY_FONT     = "#e2e8f0"
PLOTLY_BLUE     = "#3b82f6"   # nodos Base
PLOTLY_RED      = "#ef4444"   # nodos Contratista / anomalías
PLOTLY_EDGE     = "rgba(148, 163, 184, 0.35)"  # líneas neutras


def _plotly_dark_layout(fig, height: int = 520, title: str | None = None):
    """Aplica un layout dark-pro consistente a cualquier figure de Plotly."""
    fig.update_layout(
        paper_bgcolor=PLOTLY_PAPER,
        plot_bgcolor=PLOTLY_BG,
        font=dict(color=PLOTLY_FONT, family="Inter, sans-serif", size=12),
        height=height,
        margin=dict(l=10, r=10, t=50 if title else 20, b=10),
        title=dict(text=title, x=0.01, font=dict(size=15, color=PLOTLY_FONT)) if title else None,
        legend=dict(bgcolor="rgba(15, 23, 42, 0.7)", bordercolor=PLOTLY_GRID, borderwidth=1),
    )
    return fig


def construir_grafo_nlp(
    df_base: pd.DataFrame,
    df_contratista: pd.DataFrame,
    umbral: int,
    max_nodos: int = 18,
):
    """Grafo bipartito interactivo: Base (azul) ←→ Contratista (rojo).

    Las aristas se dibujan SOLO cuando la similitud semántica entre la
    descripción Base y la descripción Contratista supera `umbral` (%).
    Limitamos a `max_nodos` por lado para mantener la legibilidad
    visual y el tiempo de cómputo del producto cartesiano.
    """
    if not PLOTLY_OK:
        return None

    # Tomamos las top-N partidas más caras de cada lado (las más relevantes
    # para el comparativo). Esto evita un grafo de 200 nodos ilegible.
    df_b = df_base.copy()
    df_b["_w"] = df_b.get("Parcial", df_b["Precio"] * df_b["Cantidad"]).abs()
    df_b = df_b.nlargest(max_nodos, "_w").reset_index(drop=True)

    df_c = df_contratista.copy()
    df_c["_w"] = df_c.get("Parcial", df_c["Precio"] * df_c["Cantidad"]).abs()
    df_c = df_c.nlargest(max_nodos, "_w").reset_index(drop=True)

    n_base = len(df_b)
    n_cont = len(df_c)
    if n_base == 0 or n_cont == 0:
        return None

    # Layout bipartito: Base a la izquierda (x=0), Contratista a la derecha (x=1)
    y_base = np.linspace(1, 0, n_base) if n_base > 1 else np.array([0.5])
    y_cont = np.linspace(1, 0, n_cont) if n_cont > 1 else np.array([0.5])

    # Calcular aristas (similitud) para cada par y filtrar por umbral
    edge_x, edge_y, edge_widths = [], [], []
    edges_meta = []
    for i, row_b in df_b.iterrows():
        desc_b = str(row_b["Descripción"])
        for j, row_c in df_c.iterrows():
            desc_c = str(row_c["Descripción"])
            sim = comparar_semantica(desc_b, desc_c)
            if sim >= umbral:
                edge_x += [0.0, 1.0, None]
                edge_y += [y_base[i], y_cont[j], None]
                edge_widths.append(sim)
                edges_meta.append((i, j, sim))

    # Trazo de aristas (líneas grises semitransparentes)
    edge_trace = go.Scatter(
        x=edge_x, y=edge_y,
        mode="lines",
        line=dict(color=PLOTLY_EDGE, width=1.2),
        hoverinfo="skip",
        showlegend=False,
    )

    # Trazo de nodos Base (azul)
    base_hover = [
        f"<b>BASE</b><br>{str(r['Descripción'])[:80]}<br>Cód: {r['Código']}<br>Precio: S/.{r['Precio']:.2f}"
        for _, r in df_b.iterrows()
    ]
    nodes_base = go.Scatter(
        x=[0.0] * n_base,
        y=y_base,
        mode="markers+text",
        marker=dict(size=18, color=PLOTLY_BLUE, line=dict(color="white", width=1.5),
                    symbol="circle"),
        text=[f"B{i+1}" for i in range(n_base)],
        textposition="middle left",
        textfont=dict(color=PLOTLY_FONT, size=10),
        hovertext=base_hover,
        hoverinfo="text",
        name="Insumos Base",
    )

    # Trazo de nodos Contratista (rojo)
    cont_hover = [
        f"<b>CONTRATISTA</b><br>{str(r['Descripción'])[:80]}<br>Cód: {r['Código']}<br>Precio: S/.{r['Precio']:.2f}"
        for _, r in df_c.iterrows()
    ]
    nodes_cont = go.Scatter(
        x=[1.0] * n_cont,
        y=y_cont,
        mode="markers+text",
        marker=dict(size=18, color=PLOTLY_RED, line=dict(color="white", width=1.5),
                    symbol="circle"),
        text=[f"C{j+1}" for j in range(n_cont)],
        textposition="middle right",
        textfont=dict(color=PLOTLY_FONT, size=10),
        hovertext=cont_hover,
        hoverinfo="text",
        name="Insumos Contratista",
    )

    fig = go.Figure(data=[edge_trace, nodes_base, nodes_cont])
    fig.update_xaxes(visible=False, range=[-0.25, 1.25])
    fig.update_yaxes(visible=False, range=[-0.1, 1.1])
    _plotly_dark_layout(
        fig,
        height=max(400, 32 * max(n_base, n_cont) + 100),
        title=f"Grafo de Conexiones Semánticas · umbral ≥ {umbral}%",
    )
    fig.update_layout(showlegend=True, legend=dict(orientation="h", y=1.06, x=0.5, xanchor="center"))

    return fig, len(edges_meta), n_base, n_cont


def construir_scatter_3d_anomalias(df_compare: pd.DataFrame):
    """Scatter 3D Plasma: Precio_Base × Precio_Contratista × Sobrecosto.

    El color codifica el Score de Anomalía (Isolation Forest):
    - Amarillo/blanco = anomalías críticas (puntos aislados de la masa).
    - Azul/morado oscuro = partidas normales (cluster central).
    El hover muestra Código, Descripción, Categoría y Desviación %.
    """
    if not PLOTLY_OK or df_compare is None or df_compare.empty:
        return None

    df = df_compare.copy()
    # Normalizamos columnas obligatorias
    for col in ["Precio_Base", "Precio_Contratista", "Sobrecosto (S/.)", "Score de Anomalía"]:
        if col not in df.columns:
            return None
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Precio_Base", "Precio_Contratista", "Sobrecosto (S/.)"])
    if df.empty:
        return None

    # El tamaño del marcador escala suavemente con el sobrecosto absoluto
    abs_sobre = df["Sobrecosto (S/.)"].abs()
    sizes = 6 + 20 * (abs_sobre - abs_sobre.min()) / max(abs_sobre.max() - abs_sobre.min(), 1)

    fig = px.scatter_3d(
        df,
        x="Precio_Base",
        y="Precio_Contratista",
        z="Sobrecosto (S/.)",
        color="Score de Anomalía",
        color_continuous_scale="Plasma",
        range_color=(0, 100),
        hover_data={
            "Código": True,
            "Descripción": True,
            "Categoría": True,
            "Desviación (%)": ":.2f",
            "Precio_Base": ":.2f",
            "Precio_Contratista": ":.2f",
            "Sobrecosto (S/.)": ":.2f",
            "Score de Anomalía": ":.1f",
        },
        labels={
            "Precio_Base": "Precio Base (S/.)",
            "Precio_Contratista": "Precio Contratista (S/.)",
            "Sobrecosto (S/.)": "Sobrecosto (S/.)",
            "Score de Anomalía": "Score IA",
        },
    )
    fig.update_traces(marker=dict(size=sizes, line=dict(color="rgba(255,255,255,0.3)", width=0.5),
                                  opacity=0.85))
    _plotly_dark_layout(
        fig,
        height=620,
        title="Mapa 3D de Anomalías · Isolation Forest sobre el espacio Base × Contratista × Sobrecosto",
    )
    fig.update_scenes(
        xaxis=dict(backgroundcolor=PLOTLY_BG, gridcolor=PLOTLY_GRID,
                   showbackground=True, color=PLOTLY_FONT),
        yaxis=dict(backgroundcolor=PLOTLY_BG, gridcolor=PLOTLY_GRID,
                   showbackground=True, color=PLOTLY_FONT),
        zaxis=dict(backgroundcolor=PLOTLY_BG, gridcolor=PLOTLY_GRID,
                   showbackground=True, color=PLOTLY_FONT),
    )
    return fig


def responder_hitl_local(pregunta: str, df_compare: pd.DataFrame) -> str:
    """Asistente IA *local* (sin API key) — fallback cuando Groq no está configurado.

    Hace una búsqueda fuzzy por la descripción/código en `df_compare` y devuelve
    los datos comparativos (Sobrecosto, Desviación, Score IA) de la mejor coincidencia.
    Soporta intents simples: "top anomalías", "resumen", o búsqueda por texto.
    """
    if df_compare is None or df_compare.empty:
        return "_No hay datos comparativos cargados aún. Sube ambos archivos en *Datos Base*._"

    p = pregunta.lower().strip()

    # Intent 1: resumen / kpis globales
    if any(k in p for k in ["resumen", "kpi", "global", "total", "general"]):
        n = len(df_compare)
        sob = df_compare["Sobrecosto (S/.)"].sum(skipna=True)
        crit = int((df_compare["Score de Anomalía"] > 75).sum())
        prec = int(((df_compare["Score de Anomalía"] > 50) &
                    (df_compare["Score de Anomalía"] <= 75)).sum())
        return (
            f"**Resumen del comparativo actual**\n\n"
            f"- Partidas analizadas: **{n}**\n"
            f"- Sobrecosto total: **S/. {sob:,.2f}**\n"
            f"- Anomalías críticas (>75): :red[**{crit}**]\n"
            f"- Anomalías de precaución (50-75): :orange[**{prec}**]"
        )

    # Intent 2: top anomalías
    if any(k in p for k in ["top", "peores", "más graves", "mas graves", "críticas", "criticas"]):
        top = df_compare.nlargest(5, "Score de Anomalía")
        lineas = [":red[**Top 5 anomalías por Score IA:**]\n"]
        for _, r in top.iterrows():
            lineas.append(
                f"- `{r['Código']}` · {str(r['Descripción'])[:60]}  \n"
                f"  Sobrecosto **S/. {r['Sobrecosto (S/.)']:+,.2f}** · "
                f"Desv **{r['Desviación (%)']:+.1f}%** · Score **{r['Score de Anomalía']:.1f}**"
            )
        return "\n".join(lineas)

    # Intent 3: búsqueda por texto/código (fuzzy match)
    mejor_score = 0.0
    mejor_row = None
    for _, row in df_compare.iterrows():
        texto = f"{row['Código']} {row['Descripción']}".lower()
        # Coincidencia simple por substring (más rápida que SequenceMatcher para muchos rows)
        if any(tok in texto for tok in p.split() if len(tok) >= 3):
            score = SequenceMatcher(None, p, texto).ratio()
            if score > mejor_score:
                mejor_score = score
                mejor_row = row

    if mejor_row is not None and mejor_score > 0.15:
        r = mejor_row
        score_val = r["Score de Anomalía"]
        if score_val > 75:
            nivel_md = ":red[**CRÍTICO**]"
        elif score_val > 50:
            nivel_md = ":orange[**PRECAUCIÓN**]"
        else:
            nivel_md = ":green[**NORMAL**]"
        return (
            f"{nivel_md} · **Encontré esta partida:**\n\n"
            f"- **Código:** `{r['Código']}`\n"
            f"- **Descripción:** {r['Descripción']}\n"
            f"- **Categoría:** {r.get('Categoría', '—')}\n"
            f"- **Precio Base:** S/. {r['Precio_Base']:,.2f}\n"
            f"- **Precio Contratista:** S/. {r['Precio_Contratista']:,.2f}\n"
            f"- **Sobrecosto:** S/. {r['Sobrecosto (S/.)']:+,.2f}\n"
            f"- **Desviación:** {r['Desviación (%)']:+.2f}%\n"
            f"- **Score de Anomalía:** {score_val:.1f}/100\n"
            f"- **Método de cruce:** {r['Método de Cruce']}\n\n"
            f"_Sugerencia:_ "
            + (
                "**RECHAZAR o REVISAR** — desviación crítica, requiere justificación técnica del contratista."
                if score_val > 75 else
                "**REVISAR MANUALMENTE** — desviación moderada, valida con análisis de mercado."
                if score_val > 50 else
                "**APROBAR** — partida dentro de rangos normales."
            )
        )

    return (
        "No encontré una partida que coincida claramente con tu consulta.  \n"
        "Prueba con:\n"
        "- *Resumen* → KPIs globales\n"
        "- *Top anomalías* → las 5 más críticas\n"
        "- *Acero corrugado*, *concreto*, o cualquier palabra de la descripción"
    )


# ============================================================
# CHATBOT HITL  ·  Groq + Llama 3.3 70B
# ------------------------------------------------------------
# NOTA: el modelo "llama3-70b-8192" fue DECOMMISSIONED por Groq
# (error 400 model_decommissioned). El reemplazo oficial es
# "llama-3.3-70b-versatile": mismo tamaño, ventana de contexto
# de 128K tokens, mejor calidad de razonamiento y soporte activo.
# Ver: https://console.groq.com/docs/deprecations
# ============================================================
GROQ_MODEL_ID = "llama-3.3-70b-versatile"


def construir_contexto_anomalias(df_compare: pd.DataFrame | None, top_n: int = 6) -> str:
    """Resumen MUY compacto del comparativo para inyectar al LLM.

    Mantenemos el payload pequeño a propósito: si el contexto es muy grande,
    Groq puede devolver errores de límite de tokens o la latencia se dispara.
    Solo enviamos:
    - KPIs globales en una línea por métrica.
    - Top-N (default 6) partidas más anómalas en formato Markdown ligero,
      con descripciones recortadas a 60 caracteres.
    """
    if df_compare is None or df_compare.empty:
        return (
            "Sin datos comparativos. El auditor aún no cargó ambos archivos "
            "en la pestaña 'Datos Base'."
        )

    df = df_compare.copy()
    n_total      = len(df)
    n_criticas   = int((df["Score de Anomalía"] > 75).sum())
    n_precaucion = int(((df["Score de Anomalía"] > 50) & (df["Score de Anomalía"] <= 75)).sum())
    sobrecosto   = float(df["Sobrecosto (S/.)"].sum(skipna=True))
    total_base   = float(df["Parcial_Base"].sum(skipna=True))
    desv_global  = (sobrecosto / total_base * 100) if total_base > 0 else 0.0

    encabezado = (
        f"### KPIs\n"
        f"- Partidas: {n_total} | Críticas (>75): {n_criticas} | Precaución (50-75): {n_precaucion}\n"
        f"- Sobrecosto total: S/. {sobrecosto:,.0f} ({desv_global:+.1f}%)\n\n"
        f"### Top {min(top_n, len(df))} anomalías\n"
    )

    df_top = df.sort_values("Score de Anomalía", ascending=False).head(top_n)
    lineas = []
    for _, row in df_top.iterrows():
        desc = str(row["Descripción"])[:60]
        cod  = str(row["Código"])[:14]
        pb   = float(row["Precio_Base"]) if pd.notna(row["Precio_Base"]) else 0.0
        pc   = float(row["Precio_Contratista"]) if pd.notna(row["Precio_Contratista"]) else 0.0
        des  = float(row["Desviación (%)"]) if pd.notna(row["Desviación (%)"]) else 0.0
        sc   = float(row["Score de Anomalía"]) if pd.notna(row["Score de Anomalía"]) else 0.0
        lineas.append(
            f"- `{cod}` {desc} | S/.{pb:.0f}→{pc:.0f} ({des:+.0f}%) | score {sc:.0f}"
        )

    return encabezado + "\n".join(lineas)


def _limpiar_api_key(raw: str) -> str:
    """Sanea la API Key eliminando espacios, saltos de línea, BOM, comillas y NBSP.

    Errores típicos de copy-paste que rompen la autenticación:
    - Espacio invisible al inicio/fin (`" gsk_..."`).
    - Salto de línea pegado del clipboard (`\\r`, `\\n`).
    - BOM UTF-8 (`\\ufeff`) si vino desde un archivo de texto.
    - NBSP (`\\xa0`) cuando se pega desde la web.
    - Comillas envolventes (`"gsk_..."`) si el usuario incluyó las comillas del .env.
    """
    if raw is None:
        return ""
    s = str(raw)
    # quitamos BOM y NBSP que .strip() NO elimina por defecto
    s = s.replace("\ufeff", "").replace("\xa0", " ")
    s = s.strip()
    # comillas envolventes accidentales
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"', "`"):
        s = s[1:-1].strip()
    return s


def consultar_groq(messages_historial: list[dict], contexto_anomalias: str, api_key: str) -> str:
    """Llama a la API de Groq con el contexto de anomalías inyectado.

    Lanza:
        ImportError: si la librería groq no está instalada.
        ValueError:  si no hay API Key.
        RuntimeError: si la API responde con error (la causa original va en .args).
    """
    if Groq is None:
        raise ImportError(
            "La librería 'groq' no está instalada. Ejecuta: pip install groq"
        )

    # ============================================================
    # 1. LIMPIEZA ESTRICTA + INYECCIÓN POR VARIABLE DE ENTORNO
    # ------------------------------------------------------------
    # En lugar de pasar la llave por parámetro al constructor de Groq()
    # (donde algunas versiones del SDK la validan o cachean de forma rara),
    # la inyectamos directamente en os.environ["GROQ_API_KEY"] y dejamos
    # que el SDK la lea por sí mismo. Es el flujo "oficial" recomendado
    # por Groq y elimina toda una clase de errores de auth silenciosos.
    # ============================================================
    api_key_limpia = _limpiar_api_key(api_key)

    if not api_key_limpia:
        raise ValueError(
            "API Key de Groq no configurada. Ingrésala en la pestaña 'Configuración'."
        )

    os.environ["GROQ_API_KEY"] = api_key_limpia

    # ============================================================
    # 3. DEBUGGING (TEMPORAL) — verifica en la terminal que la llave llegue
    #    completa, sin recortes ni espacios. Quítalo en producción.
    # ============================================================
    print(
        f"[LicitAI/Groq DEBUG] Llamando a Groq con llave: "
        f"{api_key_limpia[:8]}... | Longitud: {len(api_key_limpia)} | "
        f"Termina en: ...{api_key_limpia[-4:]} | "
        f"Empieza con 'gsk_': {api_key_limpia.startswith('gsk_')} | "
        f"os.environ ok: {bool(os.environ.get('GROQ_API_KEY'))}"
    )

    system_prompt = f"""Eres "LicitAI Asistente", un experto senior en ingeniería de costos y \
análisis de presupuestos S10 (formato peruano de construcción civil). Trabajas dentro de la \
plataforma LicitAI ayudando a un AUDITOR a validar las anomalías detectadas por nuestro motor \
de IA (Isolation Forest + cruce semántico NLP) al comparar un Presupuesto Base vs. la \
Propuesta económica de un Contratista.

CONTEXTO DE LAS ANOMALÍAS ACTUALMENTE DETECTADAS EN ESTA AUDITORÍA:
{contexto_anomalias}

TU MISIÓN:
1. Explicar al auditor POR QUÉ una partida fue marcada como anómala (sobrecosto, desviación %, \
precio fuera de rango histórico, score Isolation Forest alto, cruce NLP de baja similitud).
2. SUGERIR si la partida debe ser APROBADA, REVISADA MANUALMENTE o RECHAZADA, justificando \
técnicamente desde el rubro de construcción civil peruana.
3. Dar recomendaciones basadas en buenas prácticas: precios de mercado, rendimientos S10 \
típicos (h-h, h-m), márgenes razonables, posibles sobrecostos legítimos vs. injustificados.
4. Si el auditor pregunta por una partida específica, búscala en el CONTEXTO de arriba y dale \
análisis detallado (cita su código, descripción y números reales).
5. Sé conciso (máximo 220 palabras por respuesta), profesional y conversacional. Usa **negritas** \
para destacar conclusiones (APROBAR / REVISAR / RECHAZAR).
6. Si la información del contexto es insuficiente para responder, pídele al auditor más detalles.
7. Responde SIEMPRE en español.

REGLAS DURAS:
- NUNCA inventes precios, partidas o datos que no estén en el CONTEXTO.
- NUNCA des consejos legales, fiscales o contractuales fuera del análisis técnico-económico.
- Si el auditor te pide algo fuera de tu rol (chistes, política, etc.), redirígelo amablemente \
a su tarea de auditoría."""

    full_messages = [{"role": "system", "content": system_prompt}] + list(messages_historial)

    # ============================================================
    # 2. INSTANCIACIÓN SEGURA: el cliente se crea AQUÍ, justo antes
    #    de la llamada. NO le pasamos api_key como parámetro: el SDK
    #    leerá GROQ_API_KEY desde os.environ (más confiable).
    # ============================================================
    try:
        client = Groq()  # lee GROQ_API_KEY del entorno automáticamente
        completion = client.chat.completions.create(
            model=GROQ_MODEL_ID,
            messages=full_messages,
            temperature=0.3,
            max_tokens=700,
        )
    except Exception as exc:  # cubrimos auth/rate-limit/network/etc.
        # Log detallado a terminal para diagnóstico real
        print(f"[LicitAI/Groq DEBUG] ERROR de la API: {type(exc).__name__}: {exc!r}")
        raise RuntimeError(str(exc)) from exc

    return completion.choices[0].message.content


# ============================================================
# EMPTY STATE CORPORATIVO  ·  Pantalla de bienvenida profesional
# ============================================================
def render_empty_state_bienvenida() -> None:
    """Pantalla de bienvenida tipo onboarding B2B cuando no hay archivos cargados.

    Estructura:
      1. Encabezado centrado (título + subtítulo).
      2. Guía de 3 pasos en tarjetas (Carga · IA · Auditoría).
      3. Call-to-Action hacia el menú lateral.
    """
    # --- 1. Encabezado centrado ---
    st.markdown(
        "<h1 style='text-align:center; margin-bottom:0.25rem;'>"
        "Bienvenido a LicitAI"
        "</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align:center; color:#6B7280; font-size:1.05rem; "
        "margin-top:0; margin-bottom:1.5rem;'>"
        "Plataforma de Auditoría de Presupuestos impulsada por Inteligencia Artificial"
        "</p>",
        unsafe_allow_html=True,
    )

    st.divider()

    # --- 2. Guía de 3 pasos (tarjetas con borde) ---
    col1, col2, col3 = st.columns(3, gap="large")

    with col1:
        with st.container(border=True):
            st.markdown("### :material/upload_file: 1. Carga de Datos")
            st.caption(
                "Sube tu **Presupuesto Base S10** y la **Propuesta del Contratista** "
                "en formato Excel (.xls / .xlsx) o PDF. El sistema limpia, estandariza "
                "unidades y prepara los datos automáticamente."
            )

    with col2:
        with st.container(border=True):
            st.markdown("### :material/memory: 2. Procesamiento IA")
            st.caption(
                "Un motor de **NLP (Procesamiento de Lenguaje Natural)** cruza las "
                "partidas por Código S10 y, cuando faltan códigos, por similitud "
                "semántica de descripciones sobre un umbral configurable."
            )

    with col3:
        with st.container(border=True):
            st.markdown("### :material/troubleshoot: 3. Auditoría")
            st.caption(
                "El modelo **Isolation Forest** calcula un Score de Anomalía por "
                "partida para priorizar la revisión Human-in-the-Loop y generar "
                "el reporte ejecutivo final."
            )

    st.markdown("")  # espacio vertical

    # --- 3. Call to Action hacia el menú lateral ---
    st.info(
        "Para comenzar, dirígete a la sección **ARCHIVOS CARGADOS** en el menú "
        "lateral y sube tus reportes.",
        icon=":material/arrow_back:",
    )


# ============================================================
# PÁGINA · RESUMEN  (vista de bienvenida + KPIs globales)
# ============================================================
def render_resumen():
    df_base, df_contra, df_compare, modo, error_msg = cargar_estado_actual()

    render_page_header(
        "Dashboard de Análisis de Costos",
        "Vista general del estado del análisis y atajos a las herramientas principales.",
        icono=NAV_MD_ICON["Resumen"],
    )

    if error_msg:
        st.error(error_msg, icon=":material/error:")

    if df_base is None:
        render_empty_state_bienvenida()
        historial = obtener_historial()
        if not historial.empty:
            st.divider()
            st.subheader("Últimas auditorías guardadas", divider="gray")
            st.dataframe(historial.head(5), use_container_width=True, hide_index=True)
        return

    # --- Métricas principales (containers con borde nativo) ---
    col_m1, col_m2, col_m3, col_m4 = st.columns(4, gap="medium")

    if modo:
        total_base    = float(df_compare["Parcial_Base"].sum(skipna=True))
        total_contra  = float(df_compare["Parcial_Contratista"].sum(skipna=True))
        sobrecosto    = total_contra - total_base
        desv_pct      = (sobrecosto / total_base * 100) if total_base > 0 else 0.0
        anomalias     = int((df_compare["Score de Anomalía"] > 70).sum())

        with col_m1:
            metric_card(":material/account_balance:", "Presupuesto Base",
                        fmt_soles_corto(total_base),
                        delta=f"{len(df_compare)} partidas", up=True)
        with col_m2:
            metric_card(":material/business:", "Propuesta Contratista",
                        fmt_soles_corto(total_contra),
                        delta=f"{desv_pct:+.1f}%", up=desv_pct <= 0)
        with col_m3:
            metric_card(":material/trending_up:" if sobrecosto > 0 else ":material/trending_down:",
                        "Sobrecosto Total", fmt_soles_corto(sobrecosto),
                        delta=f"{desv_pct:+.1f}%", up=sobrecosto <= 0)
        with col_m4:
            metric_card(":material/warning:", "Anomalías Críticas (>70)", f"{anomalias}",
                        delta=f"{(anomalias/len(df_compare)*100):.1f}%" if len(df_compare) else "0%",
                        up=anomalias == 0)
    else:
        total_parcial  = float(df_base["Parcial"].sum(skipna=True))
        total_partidas = len(df_base)
        parcial_prom   = float(df_base["Parcial"].mean(skipna=True)) if total_partidas else 0.0
        parcial_max    = float(df_base["Parcial"].max(skipna=True))  if total_partidas else 0.0

        with col_m1: metric_card(":material/account_balance:", "Presupuesto Base",  fmt_soles_corto(total_parcial))
        with col_m2: metric_card(":material/list_alt:",        "Total de Partidas", f"{total_partidas:,}")
        with col_m3: metric_card(":material/show_chart:",      "Parcial Promedio",  fmt_soles_corto(parcial_prom))
        with col_m4: metric_card(":material/north_east:",      "Parcial Máximo",    fmt_soles_corto(parcial_max))

    if not modo:
        st.info(
            "Sube también la **Propuesta del Contratista** en *Datos Base* "
            "para activar el comparativo, las anomalías y el reporte de auditoría.",
            icon=":material/info:",
        )

    # --- Bloque ejecutivo: Distribución de Riesgo + Top 5 Alertas ---
    if modo and df_compare is not None and not df_compare.empty:
        st.divider()
        col1, col2 = st.columns(2, gap="large")

        # Conteo por nivel de riesgo (siempre, lo usan ambas columnas)
        scores = pd.to_numeric(df_compare["Score de Anomalía"], errors="coerce").fillna(0)
        n_normal = int((scores <= 50).sum())
        n_prec   = int(((scores > 50) & (scores <= 75)).sum())
        n_crit   = int((scores > 75).sum())

        # ---- Columna 1: Dona de distribución de riesgo ----
        with col1:
            with st.container(border=True):
                st.markdown("### :material/donut_large: Distribución de Riesgo")
                st.caption(
                    "Clasificación de las partidas según el Score de Anomalía "
                    "calculado por Isolation Forest."
                )
                if not PLOTLY_OK:
                    st.info(
                        "Instala `plotly` para ver la distribución gráfica: "
                        "`pip install plotly`",
                        icon=":material/extension:",
                    )
                else:
                    fig_dona = px.pie(
                        names=["Normal (0-50)", "Precaución (50-75)", "Crítico (75-100)"],
                        values=[n_normal, n_prec, n_crit],
                        hole=0.4,
                        color=["Normal (0-50)", "Precaución (50-75)", "Crítico (75-100)"],
                        color_discrete_map={
                            "Normal (0-50)":     "#2E7D32",
                            "Precaución (50-75)": "#ED8936",
                            "Crítico (75-100)":  "#C53030",
                        },
                    )
                    fig_dona.update_traces(
                        textposition="inside",
                        textinfo="percent+value",
                        hovertemplate="<b>%{label}</b><br>Partidas: %{value}<br>%{percent}<extra></extra>",
                    )
                    fig_dona.update_layout(
                        height=350,
                        margin=dict(l=0, r=0, t=10, b=0),
                        legend=dict(orientation="h", yanchor="bottom", y=-0.2,
                                    xanchor="center", x=0.5),
                        showlegend=True,
                    )
                    st.plotly_chart(
                        fig_dona,
                        use_container_width=True,
                        config={"displayModeBar": False},
                    )

        # ---- Columna 2: Top 5 desviaciones críticas ----
        with col2:
            with st.container(border=True):
                st.markdown("### :material/warning: Top 5 Desviaciones Críticas")
                st.caption(
                    "Insumos con el Score de Anomalía más alto detectado por la IA. "
                    "Prioriza estas partidas en la auditoría."
                )
                top5 = (
                    df_compare
                    .sort_values("Score de Anomalía", ascending=False)
                    .head(5)
                    [["Descripción", "Precio_Base", "Precio_Contratista", "Score de Anomalía"]]
                    .rename(columns={
                        "Precio_Base": "Precio Base",
                        "Precio_Contratista": "Precio Contratista",
                        "Score de Anomalía": "Score",
                    })
                )
                st.dataframe(
                    top5,
                    use_container_width=True,
                    hide_index=True,
                    height=350,
                    column_config={
                        "Descripción": st.column_config.TextColumn(
                            "Descripción", width="large",
                        ),
                        "Precio Base": st.column_config.NumberColumn(
                            "Precio Base", format="S/. %.2f",
                        ),
                        "Precio Contratista": st.column_config.NumberColumn(
                            "Precio Contratista", format="S/. %.2f",
                        ),
                        "Score": st.column_config.ProgressColumn(
                            "Score IA",
                            format="%.1f",
                            min_value=0,
                            max_value=100,
                        ),
                    },
                )

    # --- Mini-historial ---
    historial = obtener_historial()
    if not historial.empty:
        st.divider()
        st.subheader("Últimas auditorías guardadas", divider="gray")
        st.dataframe(historial.head(5), use_container_width=True, hide_index=True)


# ============================================================
# PÁGINA · DATOS BASE  (uploaders + tablas crudas/limpias)
# ============================================================
def render_datos_base():
    render_page_header(
        "Datos Base",
        "Carga aquí los archivos S10 (Drag & Drop) y revisa la limpieza automática. "
        "Los datos se conservan al cambiar de pestaña.",
        icono=NAV_MD_ICON["Datos Base"],
    )

    if pdfplumber is None:
        st.caption(
            ":material/info: Soporte PDF deshabilitado. Para activarlo: "
            "`pip install pdfplumber` y reinicia la app."
        )
    else:
        st.caption(
            ":material/check_circle: Soporte para **Excel (.xls / .xlsx)** y **PDF (.pdf)**. "
            "Los PDF de contratistas sin código S10 activan automáticamente el cruce NLP."
        )

    col_u1, col_u2 = st.columns(2, gap="large")

    tipos_aceptados = ["xls", "xlsx"] + (["pdf"] if pdfplumber is not None else [])

    with col_u1:
        st.subheader(":material/upload_file: 1 · Presupuesto Base (S10)", divider="gray")
        archivo_base_subido = st.file_uploader(
            "Arrastra el archivo o haz clic para seleccionar",
            type=tipos_aceptados,
            key="upload_base",
            help=(
                "Archivo S10 (.xls / .xlsx) con el presupuesto referencial. "
                "También acepta PDF si está instalado pdfplumber."
            ),
        )
    with col_u2:
        st.subheader(":material/upload_file: 2 · Propuesta Contratista", divider="gray")
        archivo_contra_subido = st.file_uploader(
            "Arrastra el archivo o haz clic para seleccionar",
            type=tipos_aceptados,
            key="upload_contratista",
            help=(
                "Excel (.xls / .xlsx) o PDF (.pdf) con la propuesta económica del contratista. "
                "Los PDF sin Código S10 dispararán el cruce semántico (NLP)."
            ),
        )

    # ===========================================================
    # PROCESAMIENTO + PERSISTENCIA EN SESSION_STATE
    # Solo procesamos si llega un archivo NUEVO (distinto al ya guardado),
    # para evitar reprocesar en cada rerun.
    # ===========================================================
    error_msg = None

    if archivo_base_subido is not None:
        if st.session_state.nombre_archivo_base != archivo_base_subido.name \
                or st.session_state.df_base is None:
            try:
                with st.spinner(f"Procesando {archivo_base_subido.name}..."):
                    st.session_state.df_base = cargar_presupuesto(archivo_base_subido)
                    st.session_state.nombre_archivo_base = archivo_base_subido.name
                    st.session_state.df_compare = None
                    st.toast(
                        f"Presupuesto Base cargado · {archivo_base_subido.name}",
                        icon=":material/check_circle:",
                    )
            except Exception as e:
                error_msg = f"Error procesando el Presupuesto Base: {e}"
                st.session_state.df_base = None
                st.session_state.nombre_archivo_base = None

    if archivo_contra_subido is not None:
        if st.session_state.nombre_archivo_contra != archivo_contra_subido.name \
                or st.session_state.df_contratista is None:
            try:
                with st.spinner(f"Procesando {archivo_contra_subido.name}..."):
                    st.session_state.df_contratista = cargar_presupuesto(archivo_contra_subido)
                    st.session_state.nombre_archivo_contra = archivo_contra_subido.name
                    st.session_state.df_compare = None

                    fue_pdf = archivo_contra_subido.name.lower().endswith(".pdf")
                    msg_extra = " · NLP activado (PDF sin código S10)" if fue_pdf else ""
                    st.toast(
                        f"Propuesta Contratista cargada · {archivo_contra_subido.name}{msg_extra}",
                        icon=":material/check_circle:",
                    )
            except Exception as e:
                error_msg = f"Error procesando la Propuesta Contratista: {e}"
                st.session_state.df_contratista = None
                st.session_state.nombre_archivo_contra = None

    df_base   = st.session_state.df_base
    df_contra = st.session_state.df_contratista

    if error_msg:
        st.error(error_msg, icon=":material/error:")

    if df_base is None:
        st.info(
            "Sube al menos el **Presupuesto Base** para ver los datos limpios.",
            icon=":material/info:",
        )
        return

    st.divider()

    # --- Tabs con tablas limpias ---
    tabs_labels = ["Presupuesto Base"]
    if df_contra is not None:
        tabs_labels.append("Propuesta Contratista")
    tabs = st.tabs(tabs_labels)

    config_cols = {
        "Cantidad": st.column_config.NumberColumn(format="%.2f"),
        "Precio":   st.column_config.NumberColumn("Precio",  format="S/. %.2f"),
        "Parcial":  st.column_config.NumberColumn("Parcial", format="S/. %.2f"),
    }

    with tabs[0]:
        st.caption(
            f"**{len(df_base):,}** partidas limpias · "
            f"Total: **{fmt_soles_corto(float(df_base['Parcial'].sum(skipna=True)))}**"
        )
        st.dataframe(
            df_base, use_container_width=True, height=520,
            hide_index=True, column_config=config_cols,
        )

    if df_contra is not None:
        with tabs[1]:
            st.caption(
                f"**{len(df_contra):,}** partidas limpias · "
                f"Total: **{fmt_soles_corto(float(df_contra['Parcial'].sum(skipna=True)))}**"
            )
            st.dataframe(
                df_contra, use_container_width=True, height=520,
                hide_index=True, column_config=config_cols,
            )


# ============================================================
# PÁGINA · ANÁLISIS DE COSTOS  (métricas + tabla agrupada Figma)
# ============================================================
def render_analisis_costos():
    df_base, _, df_compare, modo, error_msg = cargar_estado_actual()

    render_page_header(
        "Análisis de Costos",
        "Comparativo Base vs. Contratista con cruce inteligente y agrupación por categoría.",
        icono=NAV_MD_ICON["Análisis de Costos"],
    )

    if error_msg:
        st.error(error_msg, icon=":material/error:")

    if df_base is None:
        st.info(
            "Sube el **Presupuesto Base** en *Datos Base* para iniciar el análisis.",
            icon=":material/info:",
        )
        return

    if not modo:
        st.warning(
            "Sube también la **Propuesta del Contratista** en *Datos Base* "
            "para activar el análisis comparativo agrupado por categoría.",
            icon=":material/warning:",
        )
        return

    n_cod = int((df_compare["Método de Cruce"] == "Código S10").sum())
    n_nlp = int((df_compare["Método de Cruce"] == "IA Semántica (NLP)").sum())
    st.caption(
        f"**{len(df_compare):,}** partidas cruzadas · "
        f"**{n_cod}** por Código S10 · **{n_nlp}** por IA Semántica NLP"
    )

    # --- Métricas (containers con borde nativo) ---
    total_base    = float(df_compare["Parcial_Base"].sum(skipna=True))
    total_contra  = float(df_compare["Parcial_Contratista"].sum(skipna=True))
    sobrecosto    = total_contra - total_base
    desv_pct      = (sobrecosto / total_base * 100) if total_base > 0 else 0.0
    partidas_alto = int((df_compare["Desviación (%)"] > 10).sum())

    col_m1, col_m2, col_m3, col_m4 = st.columns(4, gap="medium")
    with col_m1:
        metric_card(":material/account_balance:", "Presupuesto Base",
                    fmt_soles_corto(total_base),
                    delta=f"{len(df_compare)} partidas", up=True)
    with col_m2:
        metric_card(":material/business:", "Propuesta Contratista",
                    fmt_soles_corto(total_contra),
                    delta=f"{desv_pct:+.1f}%", up=desv_pct <= 0)
    with col_m3:
        metric_card(":material/trending_up:" if sobrecosto > 0 else ":material/trending_down:",
                    "Sobrecosto Total", fmt_soles_corto(sobrecosto),
                    delta=f"{desv_pct:+.1f}%", up=sobrecosto <= 0)
    with col_m4:
        metric_card(":material/warning:", "Partidas con desv. >10%", f"{partidas_alto}",
                    delta=f"{(partidas_alto/len(df_compare)*100):.1f}%" if len(df_compare) else "0%",
                    up=partidas_alto == 0)

    render_legend_anomalia()

    # ============================================================
    # GRAFO DE CONEXIONES SEMÁNTICAS (NLP)
    # ============================================================
    df_contra_state = st.session_state.get("df_contratista")
    if df_contra_state is not None and not df_contra_state.empty:
        st.divider()
        with st.expander(
            ":material/visibility: Ver Visualización del Modelo IA · Grafo de Conexiones Semánticas (NLP)",
            expanded=False,
        ):
            if not PLOTLY_OK:
                st.info(
                    "Instala `plotly` para activar las visualizaciones avanzadas: "
                    "`pip install plotly`",
                    icon=":material/extension:",
                )
            else:
                col_g1, col_g2 = st.columns([0.7, 0.3])
                with col_g1:
                    st.caption(
                        "Cada **punto azul** es un insumo del *Presupuesto Base*; "
                        "cada **punto rojo** es un insumo de la *Propuesta del Contratista*. "
                        "Una línea conecta dos insumos cuya similitud semántica supera el umbral."
                    )
                with col_g2:
                    umbral_grafo = st.slider(
                        "Umbral del grafo (%)",
                        min_value=50, max_value=100,
                        value=int(st.session_state.get("umbral_nlp", 80)),
                        key="umbral_grafo_nlp",
                        help="Misma escala que el umbral global de NLP en *Configuración*.",
                    )

                with st.spinner("Calculando similitud semántica entre partidas..."):
                    resultado = construir_grafo_nlp(
                        df_base, df_contra_state, umbral_grafo, max_nodos=18,
                    )

                if resultado is None:
                    st.info("No hay suficientes partidas para construir el grafo.")
                else:
                    fig_grafo, n_edges, n_b, n_c = resultado
                    fig_grafo.update_layout(
                        height=400,
                        margin=dict(l=0, r=0, t=30, b=0),
                    )
                    st.plotly_chart(fig_grafo, use_container_width=True,
                                    config={"displayModeBar": False})
                    col_s1, col_s2, col_s3 = st.columns(3)
                    with col_s1:
                        with st.container(border=True):
                            st.metric("Insumos Base mostrados", n_b)
                    with col_s2:
                        with st.container(border=True):
                            st.metric("Insumos Contratista mostrados", n_c)
                    with col_s3:
                        with st.container(border=True):
                            st.metric(
                                "Conexiones semánticas",
                                n_edges,
                                delta=f"umbral ≥ {umbral_grafo}%",
                                delta_color="off",
                            )
                    if n_edges == 0:
                        st.warning(
                            f"Ningún par supera el umbral del **{umbral_grafo}%**. "
                            f"Baja el umbral para ver más conexiones.",
                            icon=":material/warning:",
                        )

    st.divider()

    # --- Header de la tabla agrupada (nativo, sin HTML) ---
    st.subheader(
        "Análisis comparativo agrupado por categoría",
        divider="gray",
    )
    st.caption(
        "Cruce por Código S10 + IA Semántica · Score de Anomalía por Isolation Forest"
    )

    # --- Tabla agrupada por categoría ---
    cats_presentes = [c for c in ORDEN_CATEGORIAS if c in df_compare["Categoría"].unique()]
    cats_presentes += [c for c in df_compare["Categoría"].unique() if c not in cats_presentes]
    column_config_grupo = _column_config_grupo()

    for categoria in cats_presentes:
        grupo = df_compare[df_compare["Categoría"] == categoria].copy()
        if grupo.empty:
            continue

        grupo = grupo.sort_values("Score de Anomalía", ascending=False)
        total_grupo = float(grupo["Parcial_Contratista"].sum(skipna=True))

        # Cabecera de categoría: subheader nativo + caption con totales
        st.subheader(categoria, divider="gray")
        st.caption(
            f"**{len(grupo)}** partidas · "
            f"Total contratista: **{fmt_soles_corto(total_grupo)}**"
        )
        grupo_view = grupo[["Descripción", "Precio_Base", "Precio_Contratista", "Score de Anomalía"]]
        st.dataframe(
            grupo_view,
            use_container_width=True,
            hide_index=True,
            column_config=column_config_grupo,
        )


# ============================================================
# PÁGINA · ANOMALÍAS  (Auditoría HITL + Confirmar Revisión)
# ============================================================
def render_anomalias():
    df_base, _, df_compare, modo, error_msg = cargar_estado_actual()

    render_page_header(
        "Anomalías · Auditoría Human-in-the-Loop",
        "Revisión manual de las desviaciones detectadas por la IA. "
        "Tu confirmación queda registrada en el historial.",
        icono=NAV_MD_ICON["Anomalías"],
    )

    if error_msg:
        st.error(error_msg, icon=":material/error:")

    if not modo:
        st.info(
            "Sube **ambos archivos** en *Datos Base* para detectar anomalías "
            "y habilitar el panel de auditoría.",
            icon=":material/info:",
        )
        return

    # ============================================================
    # MAPA 3D DE CLUSTERS · Isolation Forest
    # ============================================================
    with st.expander(
        ":material/visibility: Ver Visualización del Modelo IA · Mapa 3D de Clusters (Isolation Forest)",
        expanded=False,
    ):
        if not PLOTLY_OK:
            st.info(
                "Instala `plotly` para activar el mapa 3D: `pip install plotly`",
                icon=":material/extension:",
            )
        else:
            fig_3d = construir_scatter_3d_anomalias(df_compare)
            if fig_3d is None:
                st.info("No hay datos suficientes para generar el mapa 3D.")
            else:
                fig_3d.update_layout(
                    height=400,
                    margin=dict(l=0, r=0, t=30, b=0),
                )
                st.plotly_chart(fig_3d, use_container_width=True,
                                config={"displayModeBar": False})
                st.caption(
                    "*Color = Score de Anomalía (0 = normal, 100 = crítico). "
                    "Tamaño = magnitud absoluta del sobrecosto. "
                    "Pasa el cursor sobre cada punto para ver código, descripción y desviación.* "
                    "Rota el gráfico arrastrando con el mouse."
                )

    st.divider()

    col_alertas, col_form = st.columns([0.58, 0.42], gap="large")

    # ---------- Columna izquierda: alertas (con st.error / st.warning nativos) ----------
    with col_alertas:
        st.subheader(":material/notifications_active: Alertas de Sobrecosto", divider="gray")

        df_alertas = df_compare.assign(
            _abs_desv=df_compare["Desviación (%)"].abs()
        ).sort_values("_abs_desv", ascending=False)
        df_alertas = df_alertas[df_alertas["_abs_desv"] > 5]

        if df_alertas.empty:
            st.success(
                "Sin desviaciones significativas (>5%). El contratista se mantiene alineado al base.",
                icon=":material/check_circle:",
            )
        else:
            for _, row in df_alertas.head(10).iterrows():
                desv  = row["Desviación (%)"]
                sobre = row["Sobrecosto (S/.)"]
                desc  = str(row["Descripción"])[:75]
                cod   = row["Código"]
                score = row.get("Score de Anomalía", 0)
                signo = "+" if sobre >= 0 else ""

                # Una sola tarjeta nativa por alerta — evitamos HTML completo
                with st.container(border=True):
                    if abs(desv) > 10:
                        nivel_md = ":red[**CRÍTICO**]"
                        emoji_icon = ":material/error:"
                    elif abs(desv) > 5:
                        nivel_md = ":orange[**MODERADO**]"
                        emoji_icon = ":material/warning:"
                    else:
                        nivel_md = ":green[**OK**]"
                        emoji_icon = ":material/check:"

                    st.markdown(f"{emoji_icon}  {nivel_md} · **{desc}**")
                    st.caption(
                        f"Cód. `{cod}`  ·  "
                        f"Sobrecosto: **{signo}S/. {sobre:,.2f}**  ·  "
                        f"Desv: **{desv:+.1f}%**  ·  "
                        f"Score IA: **{score:.1f}**"
                    )

    # ============================================================
    # ---------- Columna derecha: CHATBOT HITL como justificación
    # ============================================================
    with col_form:
        st.subheader(":material/forum: Asistente IA · Justificación del Auditor", divider="gray")

        api_key = _limpiar_api_key(st.session_state.get("api_key_persistente", ""))
        usa_groq = bool(api_key) and (Groq is not None)

        if usa_groq:
            st.caption(
                f":material/sensors: Modo **IA Avanzada** activo · `{GROQ_MODEL_ID}` (Groq) · "
                f"Contexto: {len(df_compare)} partidas"
            )
        else:
            motivo = (
                "API Key no configurada"
                if Groq is not None else "librería `groq` no instalada"
            )
            st.caption(
                f":material/offline_bolt: Modo **IA Local** (fallback, {motivo}) · "
                f"Búsqueda fuzzy + análisis estadístico sobre {len(df_compare)} partidas"
            )

        col_btn1, col_btn2 = st.columns([0.5, 0.5])
        with col_btn1:
            if st.button(":material/delete_sweep: Limpiar chat",
                         use_container_width=True, key="btn_clear_chat"):
                st.session_state.messages = []
                st.rerun()
        with col_btn2:
            ejemplo = st.button(":material/lightbulb: Ejemplo",
                                use_container_width=True, key="btn_ejemplo_chat")

        if not st.session_state.messages:
            with st.chat_message("assistant"):
                st.markdown(
                    "Soy tu **Asistente IA de Auditoría**. Pregúntame por una partida "
                    "específica y te daré sus datos comparativos en tiempo real.\n\n"
                    "Ejemplos:\n"
                    "- *Top anomalías*\n"
                    "- *Resumen general*\n"
                    "- *Acero corrugado* (o cualquier descripción)"
                )

        if ejemplo:
            st.session_state.messages.append({"role": "user", "content": "Top anomalías"})
            respuesta_ej = responder_hitl_local("top anomalías", df_compare)
            st.session_state.messages.append({"role": "assistant", "content": respuesta_ej})
            st.rerun()

        chat_container = st.container(height=320, border=True)
        with chat_container:
            for msg in st.session_state.messages:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

        if prompt := st.chat_input(
            "Pregunta por una partida o escribe tu observación...",
            key="chat_input_hitl",
        ):
            st.session_state.messages.append({"role": "user", "content": prompt})

            respuesta = None
            if usa_groq:
                try:
                    contexto = construir_contexto_anomalias(df_compare)
                    respuesta = consultar_groq(
                        st.session_state.messages, contexto, api_key,
                    )
                except Exception as e:
                    st.error(
                        f"Error de Groq: {str(e)} — usando IA local de fallback.",
                        icon=":material/error:",
                    )
                    print(f"[LicitAI/Groq] Error detallado: {repr(e)}")
                    respuesta = responder_hitl_local(prompt, df_compare)
            else:
                respuesta = responder_hitl_local(prompt, df_compare)

            st.session_state.messages.append({"role": "assistant", "content": respuesta})
            st.rerun()

        st.divider()
        st.subheader(":material/task_alt: Estado de Revisión", divider="gray")
        st.caption(
            "Toda la conversación de arriba se guarda como **justificación técnica** "
            "al confirmar la revisión."
        )

        if st.button("Confirmar Revisión", type="primary",
                     use_container_width=True, icon=":material/check_circle:"):
            n_partidas       = len(df_compare)
            sobrecosto_total = float(df_compare["Sobrecosto (S/.)"].sum(skipna=True))
            anomalias        = int((df_compare["Score de Anomalía"] > 70).sum())
            n_cod_btn = int((df_compare["Método de Cruce"] == "Código S10").sum())
            n_nlp_btn = int((df_compare["Método de Cruce"] == "IA Semántica (NLP)").sum())

            if st.session_state.messages:
                comentarios = "\n\n".join(
                    f"[{m['role'].upper()}] {m['content']}"
                    for m in st.session_state.messages
                )
            else:
                comentarios = "(sin conversación con el asistente)"

            audit_id = guardar_auditoria(
                total_partidas=n_partidas,
                sobrecosto_total=sobrecosto_total,
                anomalias_detectadas=anomalias,
                cruces_codigo=n_cod_btn,
                cruces_nlp=n_nlp_btn,
                comentario=comentarios,
            )

            st.session_state.audit_confirmed = True
            st.session_state.audit_comments  = comentarios
            st.session_state.audit_timestamp = datetime.now()
            st.session_state.audit_n         = n_partidas
            st.session_state.audit_id        = audit_id
            st.toast(
                f"Auditoría #{audit_id} guardada en historial",
                icon=":material/check_circle:",
            )

        if st.session_state.get("audit_confirmed"):
            ts          = st.session_state.get("audit_timestamp", datetime.now())
            n_validados = st.session_state.get("audit_n", len(df_compare))
            audit_id    = st.session_state.get("audit_id", "—")
            coments_g   = st.session_state.get("audit_comments", "") or "(sin observaciones)"
            ts_str      = ts.strftime("%d/%m/%Y %H:%M")
            quote = f"{coments_g[:220]}{'…' if len(coments_g) > 220 else ''}"

            # Tarjeta de resumen 100% nativa: success() para el banner +
            # container con borde para los detalles. Sin gradients ni HTML.
            st.success(
                f"**Auditoría Finalizada · #{audit_id}**",
                icon=":material/verified:",
            )
            with st.container(border=True):
                col_a1, col_a2 = st.columns(2)
                col_a1.metric("Insumos validados", f"{n_validados:,}")
                col_a2.metric("Confirmada el", ts_str)
                st.caption(f"_Justificación (extracto):_ \"{quote}\"")
            st.caption("Consulta el historial completo en la pestaña **Reportes**.")


# ============================================================
# PÁGINA · REPORTES  (descarga, email simulado e historial SQLite)
# ============================================================
def render_reportes():
    _, _, df_compare, modo, _ = cargar_estado_actual()

    render_page_header(
        "Reportes y Auditorías",
        "Exporta el comparativo actual y consulta el historial completo de auditorías.",
        icono=NAV_MD_ICON["Reportes"],
    )

    st.subheader(":material/file_download: Reporte de la auditoría actual", divider="gray")

    if not modo:
        st.info(
            "Sube **ambos archivos** en *Datos Base* para generar el reporte de la auditoría actual.",
            icon=":material/info:",
        )
    else:
        try:
            excel_bytes = generar_reporte_excel(
                df_compare,
                comentarios=st.session_state.get("audit_comments", ""),
            )
            ts_file = datetime.now().strftime("%Y%m%d_%H%M")

            col_dl, col_email = st.columns(2, gap="medium")

            with col_dl:
                st.download_button(
                    label="Descargar Reporte de Auditoría (Excel)",
                    icon=":material/download:",
                    data=excel_bytes,
                    file_name=f"reporte_auditoria_licitai_{ts_file}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with col_email:
                email_factory = getattr(st, "popover", None)
                if email_factory is not None:
                    email_container = st.popover(
                        ":material/mail: Enviar Reporte por Email",
                        use_container_width=True,
                    )
                else:
                    email_container = st.expander(
                        ":material/mail: Enviar Reporte por Email", expanded=False
                    )

                with email_container:
                    st.caption(
                        "El reporte se enviará al destinatario con el Excel adjunto."
                    )
                    email_destino = st.text_input(
                        "Correo de destino",
                        placeholder="auditor@empresa.com",
                        key="email_destino",
                    )
                    copia_auditor = st.checkbox(
                        "Enviarme una copia (CC)", value=True, key="email_cc"
                    )
                    enviar = st.button(
                        "Enviar Reporte",
                        icon=":material/send:",
                        type="primary",
                        use_container_width=True,
                        key="btn_enviar_email",
                    )

                    if enviar:
                        email_ok = (
                            bool(email_destino)
                            and "@" in email_destino
                            and "." in email_destino.split("@")[-1]
                        )
                        if not email_ok:
                            st.error(
                                "Ingresa un correo válido antes de enviar.",
                                icon=":material/warning:",
                            )
                        else:
                            with st.spinner("Conectando con servidor SMTP..."):
                                time.sleep(1.6)
                            with st.spinner(f"Enviando reporte a {email_destino}..."):
                                time.sleep(1.0)
                            st.toast("Reporte enviado con éxito",
                                     icon=":material/mark_email_read:")
                            st.success(
                                f"Reporte enviado a **{email_destino}**"
                                + (" (con copia al auditor)" if copia_auditor else ""),
                                icon=":material/check_circle:",
                            )
        except Exception as e:
            st.warning(
                f"No se pudo generar el reporte: {e}",
                icon=":material/warning:",
            )

    st.divider()
    st.subheader(":material/history: Historial de Auditorías", divider="gray")

    historial = obtener_historial()

    if historial.empty:
        st.info(
            "Aún no hay auditorías guardadas. "
            "Confirma tu primera revisión en la pestaña *Anomalías* para iniciar el historial.",
            icon=":material/edit_note:",
        )
        return

    col_h1, col_h2, col_h3 = st.columns(3, gap="medium")
    with col_h1:
        metric_card(":material/inventory_2:", "Total de Auditorías", f"{len(historial):,}")
    with col_h2:
        metric_card(
            ":material/warning:", "Anomalías Acumuladas",
            f"{int(historial['anomalias_detectadas'].sum()):,}",
        )
    with col_h3:
        metric_card(
            ":material/account_balance:", "Sobrecosto Acumulado",
            fmt_soles_corto(float(historial["sobrecosto_total"].sum())),
        )

    # Tabla con formato amigable
    historial_view = historial.rename(columns={
        "id":                   "#",
        "fecha":                "Fecha",
        "total_partidas":       "Partidas",
        "sobrecosto_total":     "Sobrecosto (S/.)",
        "anomalias_detectadas": "Anomalías",
        "cruces_codigo":        "Cruces · Código",
        "cruces_nlp":           "Cruces · NLP",
        "comentario":           "Comentario del Auditor",
    })

    st.dataframe(
        historial_view,
        use_container_width=True,
        hide_index=True,
        height=420,
        column_config={
            "Sobrecosto (S/.)": st.column_config.NumberColumn(
                "Sobrecosto (S/.)", format="S/. %.2f"
            ),
            "Comentario del Auditor": st.column_config.TextColumn(
                "Comentario del Auditor", width="large"
            ),
        },
    )


# ============================================================
# PÁGINA · CONFIGURACIÓN  (slider NLP + utilidades del sistema)
# ============================================================
def render_configuracion():
    render_page_header(
        "Configuración",
        "Ajustes del motor de IA y utilidades del sistema.",
        icono=NAV_MD_ICON["Configuración"],
    )

    # ----- Asistente IA · Groq + Llama 3 -----
    st.subheader(":material/forum: Asistente IA · Groq API (Llama 3)", divider="gray")

    # Callback: copia el valor del widget (clave temporal) a la clave persistente.
    # Esto evita que la API Key se borre al cambiar de pestaña — el widget se
    # destruye al desmontarse, pero `api_key_persistente` permanece en la sesión.
    def guardar_api_key():
        st.session_state["api_key_persistente"] = st.session_state.get("temp_api_key", "")

    st.text_input(
        "API Key de Groq",
        type="password",
        key="temp_api_key",
        value=st.session_state["api_key_persistente"],
        on_change=guardar_api_key,
        placeholder="gsk_...",
        help=(
            "Obtén tu API Key gratuita en https://console.groq.com/keys . "
            "La key vive solo en esta sesión (no se guarda en disco). "
            "Se conserva al navegar entre pestañas gracias a un callback on_change."
        ),
    )
    api_key_actual = (st.session_state.get("api_key_persistente") or "").strip()
    if Groq is None:
        st.error(
            "La librería `groq` no está instalada. Ejecuta: `pip install groq` y reinicia.",
            icon=":material/extension_off:",
        )
    elif api_key_actual:
        st.success(
            f"API Key configurada (•••{api_key_actual[-4:]}). "
            f"El chatbot está habilitado en la pestaña *Anomalías*.",
            icon=":material/check_circle:",
        )
    else:
        st.info(
            "Ingresa tu API Key para habilitar el **chat de auditoría con Llama 3** "
            "en la pestaña *Anomalías*.",
            icon=":material/info:",
        )
    st.caption(
        f"Modelo usado: `{GROQ_MODEL_ID}` · "
        "El contexto de las anomalías detectadas se inyecta automáticamente en cada consulta."
    )

    # ----- Motor NLP -----
    st.divider()
    st.subheader(":material/psychology: Motor NLP · Cruce Semántico", divider="gray")
    st.slider(
        "Umbral de Similitud NLP (%)",
        min_value=50,
        max_value=100,
        step=1,
        key="umbral_nlp",
        help=(
            "Qué tan estricta es la IA al emparejar partidas por descripción. "
            "Mayor valor = solo coincidencias casi idénticas. "
            "Menor valor = más rescates pero con riesgo de falsos positivos."
        ),
    )
    st.caption(
        f"Valor actual: **{st.session_state.umbral_nlp}%**.  "
        "Este umbral se aplica al recálculo del comparativo en *Análisis de Costos* y *Anomalías*."
    )

    # ----- Estado de la BD -----
    st.divider()
    st.subheader(":material/database: Base de Datos · Historial", divider="gray")
    historial = obtener_historial()
    with st.container(border=True):
        st.markdown(f"**Registros guardados:** {len(historial)}")
        st.markdown(f"**Archivo de BD:** `{DB_PATH.name}`")
        st.markdown(f"**Ruta:** `{DB_PATH}`")

    if st.button("Borrar historial completo",
                 type="secondary", icon=":material/delete:"):
        if st.session_state.get("confirm_delete"):
            borrar_historial()
            st.session_state.confirm_delete = False
            st.toast("Historial borrado", icon=":material/delete:")
            st.rerun()
        else:
            st.session_state.confirm_delete = True
            st.warning(
                "Click otra vez en el botón para confirmar el borrado completo del historial.",
                icon=":material/warning:",
            )


# ============================================================
# SIDEBAR  ·  Navegación + estado global de archivos
# ============================================================
with st.sidebar:
    # ----- Brand corporativo (sin HTML) -----
    st.title("LicitAI")
    st.caption("Ingeniería de Costos · IA Aplicada")
    st.divider()

    # ----- Navegación principal con streamlit-option-menu -----
    # Si la librería está instalada usamos el option_menu (apariencia
    # SaaS/Enterprise con iconos de Bootstrap). Si no, caemos a un
    # st.radio nativo — sin botones, sin CSS roto.
    indice_actual = NAV_NAMES.index(st.session_state.pagina_actual) \
        if st.session_state.pagina_actual in NAV_NAMES else 0

    if OPTION_MENU_OK:
        seleccion_nav = option_menu(
            menu_title=None,
            options=NAV_NAMES,
            icons=[NAV_BS_ICON[n] for n in NAV_NAMES],
            default_index=indice_actual,
            orientation="vertical",
            key="nav_option_menu",
            styles={
                # Look & feel sobrio Enterprise B2B: contenedor transparente,
                # texto medio y solo el ítem activo enfatizado en navy.
                "container":      {"padding": "0!important", "background-color": "transparent"},
                "icon":           {"font-size": "16px"},
                "nav-link": {
                    "font-size":      "14px",
                    "text-align":     "left",
                    "margin":         "2px 0",
                    "padding":        "0.55rem 0.8rem",
                    "border-radius":  "8px",
                    "--hover-color":  "#EEF2F7",
                },
                "nav-link-selected": {
                    "background-color": "#0A2540",
                    "color":            "#FFFFFF",
                    "font-weight":      "600",
                },
            },
        )
    else:
        seleccion_nav = st.radio(
            "Navegación",
            options=NAV_NAMES,
            index=indice_actual,
            label_visibility="collapsed",
            key="nav_radio",
        )
        st.caption(
            "Tip: instala `streamlit-option-menu` para una navegación más rica."
        )

    if seleccion_nav != st.session_state.pagina_actual:
        st.session_state.pagina_actual = seleccion_nav
        st.rerun()

    # ----- Estado de archivos cargados (100% nativo) -----
    st.divider()
    st.caption("ARCHIVOS CARGADOS")

    base_ok    = st.session_state.get("df_base") is not None
    contra_ok  = st.session_state.get("df_contratista") is not None
    nombre_base   = st.session_state.get("nombre_archivo_base")   or "—"
    nombre_contra = st.session_state.get("nombre_archivo_contra") or "—"
    if len(nombre_base) > 24:   nombre_base   = nombre_base[:21] + "…"
    if len(nombre_contra) > 24: nombre_contra = nombre_contra[:21] + "…"

    icon_ok  = ":material/check_circle:"
    icon_off = ":material/radio_button_unchecked:"

    with st.container(border=True):
        st.markdown(
            f"{icon_ok if base_ok else icon_off}  **Presupuesto Base**"
        )
        st.caption(nombre_base)
        st.markdown(
            f"{icon_ok if contra_ok else icon_off}  **Propuesta Contratista**"
        )
        st.caption(nombre_contra)

    st.caption("Carga / cambia los archivos en la pestaña **Datos Base**.")

    st.divider()
    st.caption(
        f":material/schedule: Última actualización: "
        f"**{datetime.now().strftime('%d %b %Y, %H:%M')}**"
    )


# ============================================================
# ROUTER  ·  Despacha la página activa
# ============================================================
PAGE_RENDERERS = {
    "Resumen":            render_resumen,
    "Análisis de Costos": render_analisis_costos,
    "Datos Base":         render_datos_base,
    "Anomalías":          render_anomalias,
    "Reportes":           render_reportes,
    "Configuración":      render_configuracion,
}

PAGE_RENDERERS[st.session_state.pagina_actual]()

