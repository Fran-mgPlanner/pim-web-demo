# -*- coding: utf-8 -*-
"""
Sistema P.I.M. - Control de Inventarios | Planeamiento Mantenimiento Mina
Estética: iOS / macOS Warm Glassmorphism (Crema/Arena en Claro · Grafito en Oscuro · Acento Coral/Terracota)
Stack: Streamlit + SQLite3 + Pandas + Plotly + Pillow + Openpyxl + OS + Shutil
"""

import io
import os
import shutil
import sqlite3
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image

# ==============================================================================
# 0. CONFIGURACIÓN INICIAL Y CONSTANTES CORPORATIVAS (SIN CAMBIOS ESTRUCTURALES)
# ==============================================================================
st.set_page_config(
    page_title="P.I.M. | Control de Inventarios - Planeamiento Mantenimiento Mina",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = os.getcwd()
DB_PATH = os.path.join(BASE_DIR, "inventario_mina.db")
PHOTOS_DIR = os.path.join(BASE_DIR, "fotos_componentes")
DEFAULT_SAP_EXCEL = os.path.join(BASE_DIR, "Inventario SAP Nov.xlsx")

ESTADOS_COMPONENTE = [
    "Montado",
    "Desmontado",
    "Por Reparar",
    "Reparado",
    "Salida por Reclamo",
    "Nuevo",
    "De Baja",
]

ROLES_SISTEMA = ["Planner", "Administrador", "Campo-Tablet"]

MODOS_FALLA_COMUNES = [
    "N/A - Cambio Programado (PCR)",
    "Desgaste prematuro de rodamientos / bujes",
    "Fuga interna / externa de sellos hidráulicos",
    "Fisura estructural en carcasa / alojamiento",
    "Contaminación severa de lubricante (partículas metálicas)",
    "Sobrecalentamiento / Falla térmica o aislamiento",
    "Fractura de dientes de engranaje / piñón",
    "Pérdida de presión / eficiencia volumétrica",
    "Falla eléctrica / cortocircuito en bobinado",
    "Daño por impacto operativo externo",
    "Otro (especificar en bitácora)",
]

os.makedirs(PHOTOS_DIR, exist_ok=True)


# ==============================================================================
# 1. UTILIDADES DE VALIDACIÓN NUMÉRICA Y CORRECCIÓN DE VIDAS
# ==============================================================================
def format_vidas_label(vidas: int) -> str:
    """Devuelve '1 vida' en singular o '<n> vidas' en plural."""
    n = int(vidas) if vidas is not None else 0
    return "1 vida" if n == 1 else f"{n} vidas"


def get_vidas_navigation_state(total_vidas: int, vida_actual: int) -> Dict:
    """
    Controla la visibilidad y estado de navegación de vidas:
    - Con 0 o 1 vida: oculta controles de navegación innecesarios.
    - Con varias vidas (>1): muestra navegación y desactiva 'anterior' o 'siguiente' según corresponda.
    """
    total = max(0, int(total_vidas))
    if total <= 1:
        return {
            "label": format_vidas_label(total),
            "show_nav": False,
            "current": 1 if total == 1 else 0,
            "disable_prev": True,
            "disable_next": True,
        }
    actual = max(1, min(int(vida_actual), total))
    return {
        "label": format_vidas_label(total),
        "show_nav": True,
        "current": actual,
        "disable_prev": actual <= 1,
        "disable_next": actual >= total,
    }


def validate_numeric_input(
    raw_value: str,
    field_name: str = "Valor",
    allow_zero: bool = True,
    allow_negative: bool = False,
    allow_decimal: bool = True,
) -> Tuple[bool, Optional[float], str]:
    """
    Valida entrada numérica frente a vacío, texto, cero, negativos y decimales.
    Retorna: (es_valido, numero_parseado, mensaje_error)
    """
    if raw_value is None or str(raw_value).strip() == "":
        return False, None, f"El campo '{field_name}' no puede estar vacío."

    cleaned = str(raw_value).strip().replace(",", ".")
    try:
        num = float(cleaned)
    except ValueError:
        return False, None, f"El campo '{field_name}' debe ser un número válido (sin texto)."

    if not allow_negative and num < 0:
        return False, None, f"El campo '{field_name}' no permite valores negativos."

    if not allow_zero and num == 0:
        return False, None, f"El campo '{field_name}' debe ser mayor a cero."

    if not allow_decimal and not num.is_integer():
        return False, None, f"El campo '{field_name}' debe ser un número entero (sin decimales)."

    return True, num, ""


def show_popover_or_expander(label: str):
    """Compatibilidad segura entre Streamlit >=1.33 (st.popover) y versiones previas."""
    if hasattr(st, "popover"):
        return st.popover(label)
    return st.expander(label)


# ==============================================================================
# 2. ESTÉTICA iOS / macOS WARM GLASSMORPHISM Y TEMA ADAPTATIVO
# ==============================================================================
def inject_ios_macos_css(theme_mode: str, sidebar_mode: str) -> None:
    """
    Aplica diseño minimalista iOS/macOS con glassmorphism moderado, paleta cálida
    (crema/arena en modo claro, grafito en modo oscuro, acento coral/terracota),
    transiciones de 220ms, cuadrícula adaptable y 3 estados de menú lateral.
    """
    light_vars = """
        --bg-app: #FBF8F3;
        --bg-mesh-1: rgba(224, 109, 83, 0.08);
        --bg-mesh-2: rgba(217, 173, 130, 0.10);
        --bg-surface: rgba(255, 253, 249, 0.78);
        --bg-surface-solid: #FFFDF9;
        --bg-elevated: rgba(243, 236, 226, 0.85);
        --bg-input: rgba(255, 255, 255, 0.90);
        --text-primary: #1F1D1A;
        --text-secondary: #6B655D;
        --text-muted: #8E877E;
        --border-subtle: rgba(130, 110, 88, 0.16);
        --border-strong: rgba(130, 110, 88, 0.28);
        --accent: #E06D53;
        --accent-hover: #C85A40;
        --accent-soft: rgba(224, 109, 83, 0.14);
        --success: #2E9E6B;
        --success-soft: rgba(46, 158, 107, 0.14);
        --warning: #D9822B;
        --warning-soft: rgba(217, 130, 43, 0.15);
        --danger: #D14343;
        --danger-soft: rgba(209, 67, 67, 0.14);
        --shadow-card: 0 8px 24px rgba(45, 35, 25, 0.06);
        --shadow-hover: 0 12px 30px rgba(45, 35, 25, 0.10);
    """

    dark_vars = """
        --bg-app: #16171A;
        --bg-mesh-1: rgba(230, 126, 101, 0.09);
        --bg-mesh-2: rgba(180, 130, 95, 0.07);
        --bg-surface: rgba(31, 33, 38, 0.80);
        --bg-surface-solid: #1F2126;
        --bg-elevated: rgba(43, 46, 54, 0.85);
        --bg-input: rgba(38, 40, 47, 0.92);
        --text-primary: #F4F1EC;
        --text-secondary: #B5AFA6;
        --text-muted: #8B857D;
        --border-subtle: rgba(255, 255, 255, 0.10);
        --border-strong: rgba(255, 255, 255, 0.20);
        --accent: #E67E65;
        --accent-hover: #F08F78;
        --accent-soft: rgba(230, 126, 101, 0.18);
        --success: #45C28A;
        --success-soft: rgba(69, 194, 138, 0.16);
        --warning: #F0A352;
        --warning-soft: rgba(240, 163, 82, 0.18);
        --danger: #F06868;
        --danger-soft: rgba(240, 104, 104, 0.18);
        --shadow-card: 0 10px 28px rgba(0, 0, 0, 0.34);
        --shadow-hover: 0 14px 34px rgba(0, 0, 0, 0.45);
    """

    if theme_mode == "Claro":
        theme_css = f":root {{ {light_vars} }}"
    elif theme_mode == "Oscuro":
        theme_css = f":root {{ {dark_vars} }}"
    else:
        theme_css = f"""
        :root {{ {light_vars} }}
        @media (prefers-color-scheme: dark) {{
            :root {{ {dark_vars} }}
        }}
        """

    # Reglas CSS para los 3 estados del menú lateral
    sidebar_state_css = ""
    if sidebar_mode == "Contraído":
        sidebar_state_css = """
        @media (min-width: 769px) {
            section[data-testid="stSidebar"] {
                width: 84px !important;
                min-width: 84px !important;
                max-width: 84px !important;
            }
            section[data-testid="stSidebar"] .sidebar-text-hide {
                display: none !important;
            }
            section[data-testid="stSidebar"] div.stButton > button {
                padding: 0.55rem 0.35rem !important;
                justify-content: center !important;
                font-size: 1.15rem !important;
            }
        }
        """
    elif sidebar_mode == "Oculto":
        sidebar_state_css = """
        section[data-testid="stSidebar"] {
            display: none !important;
        }
        """

    st.markdown(
        f"""
        <style>
        {theme_css}

        /* Base iOS/macOS Typography & No Horizontal Scroll */
        html, body, .stApp {{
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "SF Pro Display", "Segoe UI", Roboto, sans-serif !important;
            background-color: var(--bg-app) !important;
            background-image:
                radial-gradient(at 10% 8%, var(--bg-mesh-1) 0px, transparent 48%),
                radial-gradient(at 92% 14%, var(--bg-mesh-2) 0px, transparent 45%) !important;
            color: var(--text-primary) !important;
            overflow-x: hidden !important;
        }}
        .block-container {{
            padding-top: 1.2rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1420px !important;
            overflow-x: hidden !important;
        }}
        p, span, label, div, h1, h2, h3, h4, h5, h6 {{
            color: var(--text-primary);
        }}

        /* Glassmorphism Moderado en Tarjetas */
        .mac-glass-card {{
            background: var(--bg-surface);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid var(--border-subtle);
            border-radius: 16px;
            box-shadow: var(--shadow-card);
            padding: 1.15rem 1.3rem;
            margin-bottom: 0.9rem;
            transition: transform 220ms cubic-bezier(0.25, 0.8, 0.25, 1),
                        box-shadow 220ms cubic-bezier(0.25, 0.8, 0.25, 1),
                        border-color 220ms ease;
        }}
        .mac-glass-card:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-hover);
            border-color: var(--border-strong);
        }}

        /* Barra Superior (Topbar) */
        .mac-topbar {{
            background: var(--bg-surface);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border: 1px solid var(--border-subtle);
            border-radius: 16px;
            box-shadow: var(--shadow-card);
            padding: 0.75rem 1.15rem;
            margin-bottom: 1rem;
        }}

        /* Encabezado de Sección */
        .section-header-box {{
            background: var(--bg-surface);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid var(--border-subtle);
            border-left: 4px solid var(--accent);
            border-radius: 14px;
            padding: 0.95rem 1.25rem;
            margin-bottom: 1rem;
        }}
        .section-header-title {{
            font-size: 1.28rem;
            font-weight: 700;
            letter-spacing: -0.015em;
            color: var(--text-primary);
            margin: 0;
        }}
        .section-header-sub {{
            font-size: 0.86rem;
            color: var(--text-secondary);
            margin-top: 0.2rem;
        }}

        /* Cuadrícula Adaptable: Escritorio (4-5 cols), Tableta (2 cols), Móvil (1 col) */
        .responsive-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 0.85rem;
            margin-bottom: 1rem;
        }}
        .responsive-card-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.9rem;
            margin-bottom: 1rem;
        }}
        @media (max-width: 1100px) {{
            .responsive-kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
            .responsive-card-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        @media (max-width: 640px) {{
            .responsive-kpi-grid,
            .responsive-card-grid {{
                grid-template-columns: 1fr;
            }}
        }}

        .kpi-label {{
            font-size: 0.73rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-secondary);
            margin-bottom: 0.25rem;
        }}
        .kpi-num {{
            font-size: 1.5rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            color: var(--text-primary);
            line-height: 1.15;
        }}
        .kpi-meta {{
            font-size: 0.78rem;
            color: var(--text-muted);
            margin-top: 0.25rem;
        }}

        /* Badges Semánticos */
        .badge {{
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            padding: 0.26rem 0.65rem;
            font-size: 0.73rem;
            font-weight: 700;
            border-radius: 999px;
            letter-spacing: 0.02em;
        }}
        .badge-success {{ background: var(--success-soft); color: var(--success); }}
        .badge-primary {{ background: var(--accent-soft); color: var(--accent); }}
        .badge-warning {{ background: var(--warning-soft); color: var(--warning); }}
        .badge-danger  {{ background: var(--danger-soft); color: var(--danger); }}
        .badge-neutral {{ background: var(--bg-elevated); color: var(--text-secondary); border: 1px solid var(--border-subtle); }}

        /* Menú Lateral Glass */
        section[data-testid="stSidebar"] {{
            background: var(--bg-surface) !important;
            backdrop-filter: blur(16px) !important;
            -webkit-backdrop-filter: blur(16px) !important;
            border-right: 1px solid var(--border-subtle) !important;
            transition: width 220ms cubic-bezier(0.25, 0.8, 0.25, 1) !important;
        }}

        /* Botones estilo iOS/macOS con transiciones 220ms */
        div.stButton > button, div.stDownloadButton > button {{
            border-radius: 12px !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
            transition: all 200ms cubic-bezier(0.25, 0.8, 0.25, 1) !important;
        }}
        div.stButton > button[kind="primary"] {{
            background-color: var(--accent) !important;
            border: 1px solid var(--accent) !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 12px var(--accent-soft) !important;
        }}
        div.stButton > button[kind="primary"]:hover {{
            background-color: var(--accent-hover) !important;
            border-color: var(--accent-hover) !important;
            transform: translateY(-1px);
        }}
        div.stButton > button[kind="secondary"], div.stDownloadButton > button {{
            background-color: var(--bg-surface) !important;
            color: var(--text-primary) !important;
            border: 1px solid var(--border-strong) !important;
        }}
        div.stButton > button[kind="secondary"]:hover, div.stDownloadButton > button:hover {{
            border-color: var(--accent) !important;
            color: var(--accent) !important;
        }}

        /* Inputs, Selectboxes y Contenedores adaptados al tema */
        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        textarea {{
            background-color: var(--bg-input) !important;
            border-color: var(--border-subtle) !important;
            border-radius: 12px !important;
            color: var(--text-primary) !important;
            transition: border-color 180ms ease, box-shadow 180ms ease !important;
        }}
        div[data-baseweb="input"] > div:focus-within,
        div[data-baseweb="select"] > div:focus-within,
        textarea:focus {{
            border-color: var(--accent) !important;
            box-shadow: 0 0 0 3px var(--accent-soft) !important;
        }}

        /* Tabs minimalistas */
        button[data-baseweb="tab"] {{
            border-radius: 10px !important;
            color: var(--text-secondary) !important;
            font-weight: 600 !important;
            transition: all 180ms ease !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            color: var(--accent) !important;
            background-color: var(--accent-soft) !important;
        }}

        /* Zona Destructiva Separada */
        .danger-zone {{
            background: var(--danger-soft);
            border: 1px solid var(--danger);
            border-radius: 14px;
            padding: 1rem 1.15rem;
            margin-top: 0.75rem;
        }}

        /* Panel de Formulario Derecho (Desktop) / Hoja (Móvil) */
        .side-form-panel {{
            background: var(--bg-surface);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--border-strong);
            border-radius: 18px;
            box-shadow: var(--shadow-hover);
            padding: 1.2rem 1.25rem;
        }}

        {sidebar_state_css}
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_status_badge_html(estado: str) -> str:
    mapping = {
        "Montado": ("badge-success", "● Montado"),
        "Reparado": ("badge-primary", "● Reparado"),
        "Nuevo": ("badge-primary", "✦ Nuevo"),
        "Desmontado": ("badge-neutral", "○ Desmontado"),
        "Por Reparar": ("badge-warning", "▲ Por Reparar"),
        "Salida por Reclamo": ("badge-warning", "◆ Reclamo"),
        "De Baja": ("badge-danger", "✕ De Baja"),
    }
    css_cls, label = mapping.get(estado, ("badge-neutral", estado))
    return f'<span class="badge {css_cls}">{label}</span>'


# ==============================================================================
# 3. CAPA DE PERSISTENCIA SQLITE3 Y ARQUITECTURA DE DATOS SAP PM (INTACTA)
# ==============================================================================
def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def parse_ubicacion_tecnica(ubic_tec: str) -> Tuple[str, str]:
    if not ubic_tec or pd.isna(ubic_tec):
        return ("SIN_FLOTA", "SIN_EQUIPO")
    bloques = [b.strip() for b in str(ubic_tec).strip().split("-") if b.strip()]
    flota = bloques[3] if len(bloques) >= 4 else "SIN_FLOTA"
    equipo_principal = bloques[4] if len(bloques) >= 5 else "SIN_EQUIPO"
    return flota, equipo_principal


def init_db() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS componentes (
            codigo_equipo TEXT PRIMARY KEY,
            descripcion TEXT,
            flota TEXT,
            equipo_principal TEXT,
            ubicacion_tecnica TEXT,
            criticidad_abc TEXT,
            fabricante TEXT,
            serie TEXT,
            estado_actual TEXT DEFAULT 'Montado',
            ciclo_actual INTEGER DEFAULT 1,
            horas_acumuladas_totales REAL DEFAULT 0.0,
            horas_ciclo_actual REAL DEFAULT 0.0,
            gasto_acumulado_total_usd REAL DEFAULT 0.0,
            gasto_ciclo_actual_usd REAL DEFAULT 0.0,
            foto_1_path TEXT,
            foto_1_fecha TEXT,
            foto_1_comentario TEXT,
            foto_2_path TEXT,
            foto_2_fecha TEXT,
            foto_2_comentario TEXT,
            ultima_actualizacion_horas TEXT
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS historial_eventos (
            id_evento INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_equipo TEXT,
            numero_ciclo INTEGER,
            fecha_hora TEXT,
            usuario_rol TEXT,
            tipo_operacion TEXT,
            estado_asignado TEXT,
            equipo_principal TEXT,
            horas_componente_evento REAL DEFAULT 0.0,
            horas_sumadas REAL DEFAULT 0.0,
            instalado_por TEXT DEFAULT 'N/A',
            motivo_cambio TEXT DEFAULT 'N/A',
            modo_falla TEXT DEFAULT 'N/A',
            numero_ot TEXT,
            numero_oc TEXT,
            proveedor_reparador TEXT,
            tipo_intervencion TEXT DEFAULT 'Ninguna',
            detalle_kit TEXT,
            costo_kit_usd REAL DEFAULT 0.0,
            costo_reparacion_usd REAL DEFAULT 0.0,
            costo_total_evento_usd REAL DEFAULT 0.0,
            comentario_bitacora TEXT,
            foto_path TEXT,
            FOREIGN KEY (codigo_equipo) REFERENCES componentes(codigo_equipo)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS log_carga_horas (
            id_carga INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_hora TEXT,
            usuario TEXT,
            metodo_carga TEXT,
            componentes_afectados INTEGER,
            detalle TEXT
        )
        """
    )

    conn.commit()
    cur.execute("SELECT COUNT(*) as total FROM componentes")
    count = cur.fetchone()["total"]
    conn.close()

    if count == 0:
        if os.path.exists(DEFAULT_SAP_EXCEL):
            try:
                df_sap = pd.read_excel(DEFAULT_SAP_EXCEL, sheet_name="Sheet1")
                import_sap_dataframe(df_sap, usuario_rol="Sistema (Auto-Init)")
            except Exception:
                seed_initial_demo_data()
        else:
            seed_initial_demo_data()


def seed_initial_demo_data() -> None:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    demo_components = [
        ("10045801", "MOTOR DE GIRO DELANTERO AC P&H", "495H", "SHO009", "SPCC-TOQ-MIN-495H-SHO009-GIRO", "A", "KOMATSU / P&H", "SN-495H-8821", "Montado", 2, 18450.0, 4250.0, 68500.0, 24500.0, None, None, None, None, None, None, now_str),
        ("10045802", "TRANSMISION DE GIRO DERECHA", "495H", "SHO009", "SPCC-TOQ-MIN-495H-SHO009-GIRO", "A", "KOMATSU / P&H", "SN-495H-9104", "Montado", 1, 6120.0, 6120.0, 12400.0, 12400.0, None, None, None, None, None, None, now_str),
        ("10045803", "MOTOR DE EMPUJE (CROWD) PRINCIPAL", "495H", "SHO008", "SPCC-TOQ-MIN-495H-SHO008-EMPU", "A", "SIEMENS", "SN-SIE-33019", "Montado", 3, 27800.0, 3100.0, 94200.0, 31200.0, None, None, None, None, None, None, now_str),
        ("10045804", "CILINDRO HIDRAULICO DE LEVANTE IZQ", "994F", "FLO005", "SPCC-TOQ-MIN-994F-FLO005-HIDR", "A", "CATERPILLAR", "CAT-994-7741", "Montado", 2, 15900.0, 5400.0, 48000.0, 18500.0, None, None, None, None, None, None, now_str),
        ("10045805", "BOMBA HIDRAULICA DE IMPLEMENTOS", "994F", "FLO005", "SPCC-TOQ-MIN-994F-FLO005-HIDR", "A", "REXROTH / CAT", "REX-55821", "Por Reparar", 2, 14200.0, 7100.0, 39500.0, 14500.0, None, None, None, None, None, None, now_str),
        ("10045806", "REDUCTOR DE TRASLACION (PROPEL) DER", "410A", "SHO004", "SPCC-TOQ-MIN-410A-SHO004-PROP", "A", "BUCYRUS / CAT", "BUC-410-1120", "Reparado", 2, 19500.0, 0.0, 82000.0, 29000.0, None, None, None, None, None, None, now_str),
        ("10045807", "COMPRESOR DE AIRE PRINCIPAL", "410A", "SHO004", "SPCC-TOQ-MIN-410A-SHO004-NEUM", "B", "SULLAIR", "SUL-88392", "Montado", 1, 4890.0, 4890.0, 6200.0, 6200.0, None, None, None, None, None, None, now_str),
        ("10045808", "FRENO DISCO DE IZAJE (HOIST) IZQ", "495H", "SHO009", "SPCC-TOQ-MIN-495H-SHO009-IZAJ", "B", "KOMATSU / P&H", "PH-BRK-601", "Desmontado", 1, 7850.0, 7850.0, 9800.0, 9800.0, None, None, None, None, None, None, now_str),
    ]

    conn = get_connection()
    cur = conn.cursor()
    cur.executemany(
        """
        INSERT OR IGNORE INTO componentes (
            codigo_equipo, descripcion, flota, equipo_principal, ubicacion_tecnica,
            criticidad_abc, fabricante, serie, estado_actual, ciclo_actual,
            horas_acumuladas_totales, horas_ciclo_actual, gasto_acumulado_total_usd,
            gasto_ciclo_actual_usd, foto_1_path, foto_1_fecha, foto_1_comentario,
            foto_2_path, foto_2_fecha, foto_2_comentario, ultima_actualizacion_horas
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        demo_components,
    )

    demo_events = [
        ("10045801", 1, "2025-11-15 08:30:00", "Planner", "Intervención / Reparación", "Reparado", "SHO009", 14200.0, 0.0, "N/A", "PCR por Horas", "N/A - Cambio Programado (PCR)", "OT-40089120", "4500192831", "KOMATSU MITSUI MAQUINARIAS PERU", "Kit + Reparación General", "Kit de rodamientos aislados SKF + Sellos Viton + Rebobinado rotor", 6500.0, 37500.0, 44000.0, "Overhaul mayor ciclo 1 completado según estándar OEM.", None),
        ("10045801", 2, "2026-02-10 14:15:00", "Planner", "Montaje en Equipo", "Montado", "SHO009", 14200.0, 0.0, "Southern", "PCR por Horas", "N/A - Cambio Programado (PCR)", "OT-40095410", "4500198004", "TALLER ELECTRICO SPCC", "Kit de Mantenimiento", "Kit de montaje, acoplamiento y sensores PT100", 24500.0, 0.0, 24500.0, "Montado en Pala SHO009 durante parada programada. Inicia Ciclo 2.", None),
        ("10045802", 1, "2026-04-02 10:00:00", "Planner", "Intervención / Reparación", "Montado", "SHO009", 3500.0, 0.0, "Southern", "N/A", "N/A - Cambio Programado (PCR)", "OT-40101220", "4500201145", "FERREYROS / TALLER CENTRAL", "Kit de Mantenimiento", "Kit de sellos de alta presión y respiradero", 12400.0, 0.0, 12400.0, "Mantenimiento preventivo intermedio en PM de pala.", None),
        ("10045806", 2, "2026-09-18 16:40:00", "Planner", "Intervención / Reparación", "Reparado", "SHO004", 19500.0, 0.0, "Contratista", "Falla por Condición", "Desgaste prematuro de rodamientos / bujes", "OT-40118900", "4500219087", "FLSMIDTH / MAESTRANZA SUR", "Kit + Reparación General", "Kit planetario etapa 1 y 2 + Rectificado de alojamiento", 9000.0, 20000.0, 29000.0, "Componente reparado y preservado en bahía de listos.", None),
    ]

    cur.executemany(
        """
        INSERT INTO historial_eventos (
            codigo_equipo, numero_ciclo, fecha_hora, usuario_rol, tipo_operacion,
            estado_asignado, equipo_principal, horas_componente_evento, horas_sumadas,
            instalado_por, motivo_cambio, modo_falla, numero_ot, numero_oc,
            proveedor_reparador, tipo_intervencion, detalle_kit, costo_kit_usd,
            costo_reparacion_usd, costo_total_evento_usd, comentario_bitacora, foto_path
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        demo_events,
    )

    conn.commit()
    conn.close()


def save_component_photo_fifo(
    codigo_equipo: str,
    uploaded_file,
    comentario_foto: str,
    usuario_rol: str,
    registrar_en_bitacora: bool = True,
) -> Tuple[bool, str, Optional[str]]:
    if uploaded_file is None:
        return False, "No se seleccionó ningún archivo de imagen.", None

    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT codigo_equipo, ciclo_actual, estado_actual, equipo_principal,
                   horas_ciclo_actual, foto_1_path, foto_1_fecha, foto_1_comentario,
                   foto_2_path, foto_2_fecha, foto_2_comentario
            FROM componentes
            WHERE codigo_equipo = ?
            """,
            (codigo_equipo,),
        )
        row = cur.fetchone()
        if not row:
            conn.close()
            return False, f"El componente {codigo_equipo} no existe.", None

        img = Image.open(uploaded_file)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.thumbnail((1280, 1280), Image.Resampling.LANCZOS)

        timestamp_file = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        now_display = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        safe_code = "".join(c for c in str(codigo_equipo) if c.isalnum() or c in ("-", "_"))
        filename = f"{safe_code}_{timestamp_file}.jpg"
        new_photo_path = os.path.join(PHOTOS_DIR, filename)

        img.save(new_photo_path, format="JPEG", quality=82, optimize=True)

        f1_path = row["foto_1_path"]
        f2_path = row["foto_2_path"]
        f2_fecha = row["foto_2_fecha"]
        f2_com = row["foto_2_comentario"]

        comentario_limpio = (comentario_foto or "Inspección visual de componente").strip()

        if not f1_path:
            cur.execute(
                "UPDATE componentes SET foto_1_path = ?, foto_1_fecha = ?, foto_1_comentario = ? WHERE codigo_equipo = ?",
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Guardada en Espacio Fotográfico #1."
        elif not f2_path:
            cur.execute(
                "UPDATE componentes SET foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ? WHERE codigo_equipo = ?",
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Guardada en Espacio Fotográfico #2."
        else:
            if f1_path and os.path.exists(f1_path):
                try:
                    os.remove(f1_path)
                except OSError:
                    pass
            cur.execute(
                """
                UPDATE componentes
                SET foto_1_path = ?, foto_1_fecha = ?, foto_1_comentario = ?,
                    foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ?
                WHERE codigo_equipo = ?
                """,
                (f2_path, f2_fecha, f2_com, new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = (
                "Rotación FIFO aplicada: se eliminó físicamente la foto #1, "
                "la #2 pasó a #1 y la nueva foto quedó en #2."
            )

        if registrar_en_bitacora:
            cur.execute(
                """
                INSERT INTO historial_eventos (
                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                    tipo_operacion, estado_asignado, equipo_principal,
                    horas_componente_evento, comentario_bitacora, foto_path
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    codigo_equipo,
                    row["ciclo_actual"],
                    now_display,
                    usuario_rol,
                    "Inspección Fotográfica / Bitácora",
                    row["estado_actual"],
                    row["equipo_principal"],
                    row["horas_ciclo_actual"],
                    comentario_limpio,
                    new_photo_path,
                ),
            )

        conn.commit()
        conn.close()
        return True, msg_rotacion, new_photo_path
    except Exception as exc:
        return False, f"Error al procesar imagen: {exc}", None


def import_sap_dataframe(df: pd.DataFrame, usuario_rol: str) -> Tuple[int, int, str]:
    required_cols = ["Equipo", "Denominación", "Ubicac.técnica"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        return 0, 0, f"Faltan columnas obligatorias de SAP: {', '.join(missing)}"

    conn = get_connection()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nuevos = 0
    actualizados = 0

    for _, row in df.iterrows():
        cod_raw = row.get("Equipo")
        if pd.isna(cod_raw) or str(cod_raw).strip() == "":
            continue
        codigo_equipo = str(cod_raw).strip()
        if codigo_equipo.endswith(".0"):
            codigo_equipo = codigo_equipo[:-2]

        descripcion = str(row.get("Denominación", "") or "").strip()
        denom3 = str(row.get("Denominación3", "") or "").strip()
        if denom3 and denom3 != "nan" and denom3 not in descripcion:
            descripcion = f"{descripcion} - {denom3}".strip(" -")

        ubic_tec = str(row.get("Ubicac.técnica", "") or "").strip()
        flota, equipo_principal = parse_ubicacion_tecnica(ubic_tec)

        criticidad_abc = str(row.get("Indicador ABC", "A") or "A").strip()
        if criticidad_abc == "nan" or not criticidad_abc:
            criticidad_abc = "A"

        fabricante = str(row.get("Fabricante", "S/I") or "S/I").strip()
        if fabricante == "nan":
            fabricante = "S/I"

        serie = str(row.get("Fabr. Nº-serie", "S/N") or "S/N").strip()
        if serie == "nan":
            serie = "S/N"

        cur.execute("SELECT codigo_equipo FROM componentes WHERE codigo_equipo = ?", (codigo_equipo,))
        existe = cur.fetchone()

        if not existe:
            cur.execute(
                """
                INSERT INTO componentes (
                    codigo_equipo, descripcion, flota, equipo_principal,
                    ubicacion_tecnica, criticidad_abc, fabricante, serie,
                    estado_actual, ciclo_actual, horas_acumuladas_totales,
                    horas_ciclo_actual, gasto_acumulado_total_usd,
                    gasto_ciclo_actual_usd, ultima_actualizacion_horas
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Montado', 1, 0.0, 0.0, 0.0, 0.0, ?)
                """,
                (codigo_equipo, descripcion, flota, equipo_principal, ubic_tec, criticidad_abc, fabricante, serie, now_str),
            )
            nuevos += 1
        else:
            cur.execute(
                """
                UPDATE componentes
                SET descripcion = ?, flota = ?, equipo_principal = ?,
                    ubicacion_tecnica = ?, criticidad_abc = ?,
                    fabricante = ?, serie = ?
                WHERE codigo_equipo = ?
                """,
                (descripcion, flota, equipo_principal, ubic_tec, criticidad_abc, fabricante, serie, codigo_equipo),
            )
            actualizados += 1

    cur.execute(
        """
        INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle)
        VALUES (?, ?, ?, ?, ?)
        """,
        (now_str, usuario_rol, "Sincronización Maestro SAP PM", nuevos + actualizados, f"Nuevos: {nuevos} | Actualizados: {actualizados}"),
    )
    conn.commit()
    conn.close()
    return nuevos, actualizados, "Sincronización SAP completada con éxito."


def apply_hours_to_mounted_components(
    updates: List[Dict],
    usuario_rol: str,
    metodo_carga: str,
) -> Tuple[int, List[Dict]]:
    conn = get_connection()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    afectados_detalle = []

    for item in updates:
        t_type = item.get("target_type", "codigo_equipo")
        t_val = str(item.get("target_value", "")).strip()
        if t_val.endswith(".0"):
            t_val = t_val[:-2]
        mode = item.get("mode", "add")
        val = float(item.get("value", 0.0) or 0.0)
        if not t_val or val <= 0:
            continue

        if t_type == "equipo_principal":
            cur.execute(
                """
                SELECT codigo_equipo, descripcion, equipo_principal, ciclo_actual,
                       horas_ciclo_actual, horas_acumuladas_totales
                FROM componentes
                WHERE UPPER(equipo_principal) = UPPER(?) AND estado_actual = 'Montado'
                """,
                (t_val,),
            )
        else:
            cur.execute(
                """
                SELECT codigo_equipo, descripcion, equipo_principal, ciclo_actual,
                       horas_ciclo_actual, horas_acumuladas_totales
                FROM componentes
                WHERE UPPER(codigo_equipo) = UPPER(?) AND estado_actual = 'Montado'
                """,
                (t_val,),
            )

        rows = cur.fetchall()
        for r in rows:
            cod = r["codigo_equipo"]
            h_ciclo_ant = float(r["horas_ciclo_actual"] or 0.0)
            h_total_ant = float(r["horas_acumuladas_totales"] or 0.0)

            if mode == "set_cycle":
                delta = val - h_ciclo_ant
                if delta <= 0:
                    continue
                nuevo_ciclo = val
                nuevo_total = h_total_ant + delta
            else:
                delta = val
                nuevo_ciclo = h_ciclo_ant + delta
                nuevo_total = h_total_ant + delta

            cur.execute(
                """
                UPDATE componentes
                SET horas_ciclo_actual = ?, horas_acumuladas_totales = ?, ultima_actualizacion_horas = ?
                WHERE codigo_equipo = ? AND estado_actual = 'Montado'
                """,
                (round(nuevo_ciclo, 2), round(nuevo_total, 2), now_str, cod),
            )

            cur.execute(
                """
                INSERT INTO historial_eventos (
                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                    tipo_operacion, estado_asignado, equipo_principal,
                    horas_componente_evento, horas_sumadas, comentario_bitacora
                ) VALUES (?, ?, ?, ?, ?, 'Montado', ?, ?, ?, ?)
                """,
                (
                    cod,
                    r["ciclo_actual"],
                    now_str,
                    usuario_rol,
                    "Carga de Horas (Montado)",
                    r["equipo_principal"],
                    round(nuevo_ciclo, 2),
                    round(delta, 2),
                    f"Actualización de horómetro ({metodo_carga}): +{delta:,.1f} hrs sumadas.",
                ),
            )

            afectados_detalle.append(
                {
                    "Código Componente": cod,
                    "Descripción": r["descripcion"],
                    "Equipo Principal": r["equipo_principal"],
                    "Horas Previas Ciclo": round(h_ciclo_ant, 1),
                    "Horas Sumadas (+)": round(delta, 1),
                    "Nuevas Horas Ciclo": round(nuevo_ciclo, 1),
                    "Nuevas Horas Totales": round(nuevo_total, 1),
                }
            )

    if afectados_detalle:
        resumen_str = f"Se actualizaron {len(afectados_detalle)} componentes montados. Ejemplos: {', '.join(x['Código Componente'] for x in afectados_detalle[:5])}"
        cur.execute(
            """
            INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle)
            VALUES (?, ?, ?, ?, ?)
            """,
            (now_str, usuario_rol, metodo_carga, len(afectados_detalle), resumen_str),
        )

    conn.commit()
    conn.close()
    return len(afectados_detalle), afectados_detalle


def export_full_database_to_excel() -> bytes:
    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
    df_hist = pd.read_sql_query("SELECT * FROM historial_eventos ORDER BY id_evento DESC", conn)
    df_logs = pd.read_sql_query("SELECT * FROM log_carga_horas ORDER BY id_carga DESC", conn)
    conn.close()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_comp.to_excel(writer, sheet_name="Inventario_Actual", index=False)
        df_hist.to_excel(writer, sheet_name="Historial_Eventos_Costos", index=False)
        df_logs.to_excel(writer, sheet_name="Bitacora_Carga_Horas", index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=12)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    output.seek(0)
    return output.getvalue()


def generate_sap_template_excel() -> bytes:
    sample_data = pd.DataFrame(
        [
            {
                "Equipo": "10045901",
                "Denominación": "MOTOR ELECTRICO IZAJE PRINCIPAL IZQ",
                "Indicador ABC": "A",
                "Ubicac.técnica": "SPCC-TOQ-MIN-495H-SHO009-IZAJ",
                "Denominación3": "PALA ELECTRICA P&H 495HR #9",
                "Fabricante": "KOMATSU / P&H",
                "Fabr. Nº-serie": "SN-495H-IZ01",
                "Válido de": "2024-01-01",
                "Fin de validez": "9999-12-31",
            },
            {
                "Equipo": "10045902",
                "Denominación": "CONVERTIDOR DE TORQUE PRINCIPAL",
                "Indicador ABC": "A",
                "Ubicac.técnica": "SPCC-TOQ-MIN-994F-FLO005-TRAN",
                "Denominación3": "CARGADOR FRONTAL CAT 994F #5",
                "Fabricante": "CATERPILLAR",
                "Fabr. Nº-serie": "CAT-CONV-551",
                "Válido de": "2024-01-01",
                "Fin de validez": "9999-12-31",
            },
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample_data.to_excel(writer, sheet_name="Sheet1", index=False)
    buf.seek(0)
    return buf.getvalue()


def fetch_executive_kpis() -> Dict:
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM componentes", conn)
    conn.close()

    if df.empty:
        return {
            "total": 0,
            "montados": 0,
            "por_reparar": 0,
            "reparados_listos": 0,
            "gasto_acum_usd": 0.0,
            "horas_totales": 0.0,
            "costo_hora_global": 0.0,
        }

    total = len(df)
    montados = int((df["estado_actual"] == "Montado").sum())
    por_reparar = int(df["estado_actual"].isin(["Por Reparar", "Desmontado", "Salida por Reclamo"]).sum())
    reparados_listos = int(df["estado_actual"].isin(["Reparado", "Nuevo"]).sum())
    gasto_acum = float(df["gasto_acumulado_total_usd"].sum())
    horas_tot = float(df["horas_acumuladas_totales"].sum())
    costo_hr = (gasto_acum / horas_tot) if horas_tot > 0 else 0.0

    return {
        "total": total,
        "montados": montados,
        "por_reparar": por_reparar,
        "reparados_listos": reparados_listos,
        "gasto_acum_usd": gasto_acum,
        "horas_totales": horas_tot,
        "costo_hora_global": costo_hr,
    }


# ==============================================================================
# 4. NAVEGACIÓN: MENÚ LATERAL (3 ESTADOS) Y BARRA SUPERIOR (TOPBAR)
# ==============================================================================
def init_ui_session_state() -> None:
    defaults = {
        "current_module": "LOBBY",
        "user_role": "Planner",
        "selected_component": None,
        "theme_mode": "Sistema",       # Claro | Oscuro | Sistema
        "sidebar_mode": "Expandido",   # Expandido | Contraído | Oculto
        "global_search_query": "",
        "show_help_modal": False,
        "show_add_panel": False,
        "mod1_view_mode": "Tabla",     # Tarjetas | Tabla
        "mod5_view_mode": "Tarjetas",  # Tarjetas | Tabla
        "vida_idx": 1,
        "vida_comp_ref": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_sidebar() -> None:
    if st.session_state.sidebar_mode == "Oculto":
        return

    is_collapsed = st.session_state.sidebar_mode == "Contraído"
    role = st.session_state.user_role
    current = st.session_state.current_module

    with st.sidebar:
        if is_collapsed:
            st.markdown(
                "<div style='text-align:center; padding:0.4rem 0 0.8rem 0; font-size:1.35rem; font-weight:900; color:var(--accent);'>⚙️</div>",
                unsafe_allow_html=True,
            )
            if st.button("◧", help="Expandir menú lateral", use_container_width=True):
                st.session_state.sidebar_mode = "Expandido"
                st.rerun()
        else:
            c_brand, c_ctrl = st.columns([3.2, 1.3])
            with c_brand:
                st.markdown(
                    """
                    <div style="padding: 0.2rem 0 0.6rem 0;">
                        <div style="color:var(--accent); font-size:1.35rem; font-weight:900; letter-spacing:0.03em;">⚙️ P.I.M.</div>
                        <div style="font-size:0.78rem; color:var(--text-primary); font-weight:700;">CONTROL DE INVENTARIOS</div>
                        <div style="font-size:0.68rem; color:var(--text-muted);">Planeamiento Mantenimiento Mina</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with c_ctrl:
                st.write("")
                if st.button("◧", help="Contraer a solo iconos", use_container_width=True):
                    st.session_state.sidebar_mode = "Contraído"
                    st.rerun()

        st.markdown("<hr style='margin:0.4rem 0 0.7rem 0; border-color:var(--border-subtle);'>", unsafe_allow_html=True)

        # Funciones más usadas arriba
        nav_items = [
            ("LOBBY", "🏠", "Inicio (Lobby)", ["Administrador", "Planner", "Campo-Tablet"]),
            ("MOD_1", "🔍", "Buscador 360°", ["Administrador", "Planner", "Campo-Tablet"]),
            ("MOD_2", "⏱️", "Carga de Horas", ["Administrador", "Planner"]),
            ("MOD_3", "⚡", "Estados, Kits y OC", ["Administrador", "Planner"]),
            ("MOD_4", "📱", "Campo Tablet", ["Administrador", "Planner", "Campo-Tablet"]),
            ("MOD_5", "➕", "Alta SAP y Backup", ["Administrador", "Planner"]),
            ("MOD_6", "📊", "Estrategia y Costos", ["Administrador", "Planner"]),
        ]

        for mod_id, icon, label, allowed_roles in nav_items:
            if role not in allowed_roles:
                continue
            is_active = current == mod_id
            btn_label = icon if is_collapsed else f"{icon}  {label}"
            if st.button(
                btn_label,
                key=f"side_nav_{mod_id}",
                type="primary" if is_active else "secondary",
                use_container_width=True,
                help=label,
            ):
                st.session_state.current_module = mod_id
                st.rerun()

        st.markdown("<hr style='margin:0.9rem 0 0.6rem 0; border-color:var(--border-subtle);'>", unsafe_allow_html=True)

        # Abajo: Configuración, Ayuda y Perfil
        if not is_collapsed:
            st.markdown("<div style='font-size:0.72rem; font-weight:700; color:var(--text-muted); text-transform:uppercase; margin-bottom:0.35rem;'>Configuración y Perfil</div>", unsafe_allow_html=True)
            new_role = st.selectbox(
                "Perfil Activo:",
                options=ROLES_SISTEMA,
                index=ROLES_SISTEMA.index(role),
                key="sidebar_role_select",
            )
            if new_role != role:
                st.session_state.user_role = new_role
                if new_role == "Campo-Tablet" and current not in ("LOBBY", "MOD_1", "MOD_4"):
                    st.session_state.current_module = "LOBBY"
                st.rerun()

            c_cfg1, c_cfg2 = st.columns(2)
            with c_cfg1:
                if st.button("❓ Ayuda", use_container_width=True):
                    st.session_state.show_help_modal = not st.session_state.show_help_modal
                    st.rerun()
            with c_cfg2:
                if st.button("✕ Ocultar", use_container_width=True, help="Ocultar menú lateral"):
                    st.session_state.sidebar_mode = "Oculto"
                    st.rerun()
        else:
            if st.button("❓", help="Ayuda rápida", use_container_width=True):
                st.session_state.show_help_modal = not st.session_state.show_help_modal
                st.rerun()
            if st.button("✕", help="Ocultar menú", use_container_width=True):
                st.session_state.sidebar_mode = "Oculto"
                st.rerun()


def render_topbar() -> None:
    kpis = fetch_executive_kpis()
    alert_count = kpis["por_reparar"]

    st.markdown('<div class="mac-topbar">', unsafe_allow_html=True)
    col_menu, col_title, col_search, col_notif, col_profile, col_theme = st.columns([0.65, 2.3, 2.1, 0.9, 1.15, 1.35])

    with col_menu:
        # Botón para mostrar / alternar el menú lateral
        next_mode = "Expandido" if st.session_state.sidebar_mode == "Oculto" else "Oculto"
        icon_menu = "☰" if st.session_state.sidebar_mode == "Oculto" else "◧"
        if st.button(icon_menu, key="topbar_menu_toggle", help="Mostrar / ocultar menú lateral", use_container_width=True):
            st.session_state.sidebar_mode = next_mode
            st.rerun()

    with col_title:
        st.markdown(
            """
            <div style="line-height:1.15; padding-top:0.12rem;">
                <div style="font-size:1.05rem; font-weight:800; color:var(--text-primary);">P.I.M. · Control de Inventarios</div>
                <div style="font-size:0.74rem; color:var(--text-secondary);">Planeamiento Mantenimiento Mina</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_search:
        q_top = st.text_input(
            "Buscar",
            value=st.session_state.global_search_query,
            placeholder="🔍 Buscar código SAP, equipo o serie...",
            label_visibility="collapsed",
            key="topbar_search_input",
        )
        if q_top != st.session_state.global_search_query:
            st.session_state.global_search_query = q_top
            if q_top.strip():
                st.session_state.current_module = "MOD_1"
            st.rerun()

    with col_notif:
        with show_popover_or_expander(f"🔔 {alert_count}"):
            st.markdown("**Notificaciones Operativas**")
            if alert_count > 0:
                st.warning(f"Hay **{alert_count}** componente(s) en taller / por reparar / reclamo.")
                if st.button("Ver en Buscador 360°", key="notif_go_mod1", use_container_width=True):
                    st.session_state.current_module = "MOD_1"
                    st.rerun()
            else:
                st.success("Sin pendientes críticos en taller.")

    with col_profile:
        role_short = {"Administrador": "Admin", "Planner": "Planner", "Campo-Tablet": "Campo"}[st.session_state.user_role]
        st.markdown(
            f"""
            <div style="background:var(--accent-soft); color:var(--accent); border-radius:10px; padding:0.42rem 0.6rem; text-align:center; font-size:0.8rem; font-weight:700;">
                👤 {role_short}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_theme:
        theme_sel = st.selectbox(
            "Apariencia",
            options=["Claro", "Oscuro", "Sistema"],
            index=["Claro", "Oscuro", "Sistema"].index(st.session_state.theme_mode),
            label_visibility="collapsed",
            key="topbar_theme_selector",
            help="Apariencia: Claro (Crema/Arena), Oscuro (Grafito) o Sistema",
        )
        if theme_sel != st.session_state.theme_mode:
            st.session_state.theme_mode = theme_sel
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.show_help_modal:
        with st.container():
            st.info(
                "**Guía rápida P.I.M.:** Utiliza el menú izquierdo para navegar entre módulos. "
                "En **Buscador 360°** ves la hoja de vida y sus vidas/ciclos; en **Carga de Horas** actualizas horómetros de montados; "
                "en **Alta SAP y Backup** agregas componentes o descargas tu respaldo `.xlsx` / `.db`."
            )
            if st.button("Cerrar ayuda", key="close_help_btn"):
                st.session_state.show_help_modal = False
                st.rerun()


def render_section_header(title: str, subtitle: str, primary_label: Optional[str] = None, primary_action=None) -> None:
    col_h, col_act = st.columns([3.5, 1.3])
    with col_h:
        st.markdown(
            f"""
            <div class="section-header-box">
                <div class="section-header-title">{title}</div>
                <div class="section-header-sub">{subtitle}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_act:
        st.write("")
        if primary_label and primary_action:
            if st.button(primary_label, type="primary", use_container_width=True):
                primary_action()


# ==============================================================================
# 5. LOBBY PRINCIPAL (ADAPTABLE CON TARJETAS GLASS)
# ==============================================================================
def render_top_kpis() -> None:
    kpis = fetch_executive_kpis()
    st.markdown(
        f"""
        <div class="responsive-kpi-grid">
            <div class="mac-glass-card">
                <div class="kpi-label">Total Componentes</div>
                <div class="kpi-num">{kpis['total']:,}</div>
                <div class="kpi-meta">Maestro SAP PM activo</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">En Operación (Montados)</div>
                <div class="kpi-num" style="color:var(--success);">{kpis['montados']:,}</div>
                <div class="kpi-meta">Acumulando horas</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Listos (Reparado / Nuevo)</div>
                <div class="kpi-num" style="color:var(--accent);">{kpis['reparados_listos']:,}</div>
                <div class="kpi-meta">Disponibles para montaje</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">En Taller / Por Reparar</div>
                <div class="kpi-num" style="color:var(--warning);">{kpis['por_reparar']:,}</div>
                <div class="kpi-meta">Pendientes OT / OC</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Gasto Acum. · Costo/Hr</div>
                <div class="kpi-num">${kpis['gasto_acum_usd']:,.0f}</div>
                <div class="kpi-meta" style="color:var(--danger); font-weight:700;">${kpis['costo_hora_global']:,.2f} USD/hr</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_lobby() -> None:
    def _go_add():
        st.session_state.current_module = "MOD_5"
        st.session_state.show_add_panel = True
        st.rerun()

    render_section_header(
        "Resumen Operativo — P.I.M.",
        f"Vista ejecutiva y accesos rápidos habilitados para el perfil {st.session_state.user_role}.",
        primary_label="➕ Agregar componente" if st.session_state.user_role in ("Planner", "Administrador") else None,
        primary_action=_go_add if st.session_state.user_role in ("Planner", "Administrador") else None,
    )
    render_top_kpis()

    role = st.session_state.user_role
    modules_config = [
        {"id": "MOD_1", "title": "🔍 Buscador 360° y Hoja de Vida", "desc": "Consulta por flota, equipo o código SAP. KPIs de horas, $/hr, 2 últimas fotos, desglose financiero y navegación por vidas.", "roles": ["Administrador", "Planner", "Campo-Tablet"], "badge": "Todos"},
        {"id": "MOD_2", "title": "⏱️ Carga Masiva de Horas", "desc": "Actualiza horómetros por Excel/CSV o tabla editable para componentes en estado Montado.", "roles": ["Administrador", "Planner"], "badge": "Planner / Admin"},
        {"id": "MOD_3", "title": "⚡ Estados, Kits, OC y Trazabilidad", "desc": "Montaje (Southern/Contratista), desmontaje (PCR o falla), reparación (OT, OC, proveedor y costos USD).", "roles": ["Administrador", "Planner"], "badge": "Planner / Admin"},
        {"id": "MOD_4", "title": "📱 Modo Tablet Campo", "desc": "Captura o carga fotográfica en terreno con compresión automática y rotación estricta FIFO de 2 fotos.", "roles": ["Administrador", "Planner", "Campo-Tablet"], "badge": "Campo / Todos"},
        {"id": "MOD_5", "title": "➕ Alta SAP, Ingreso y Backup", "desc": "Sincroniza 'Inventario SAP Nov.xlsx', registra nuevos componentes y descarga respaldos .xlsx y .db.", "roles": ["Administrador", "Planner"], "badge": "Planner / Admin"},
        {"id": "MOD_6", "title": "📊 Estrategia y Costo por Ciclo", "desc": "Gráficos de $/hr por flota, matriz horas vs. gasto, inversión en kits por vida y ranking crítico.", "roles": ["Administrador", "Planner"], "badge": "Planner / Admin"},
    ]
    allowed = [m for m in modules_config if role in m["roles"]]

    # Cuadrícula adaptable: 3 cols escritorio, 2 tableta, 1 móvil via columnas responsivas
    for idx in range(0, len(allowed), 3):
        cols = st.columns(3)
        for c_i in range(3):
            if idx + c_i < len(allowed):
                mod = allowed[idx + c_i]
                with cols[c_i]:
                    st.markdown(
                        f"""
                        <div class="mac-glass-card" style="min-height:142px;">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.45rem;">
                                <span style="font-weight:700; font-size:0.98rem;">{mod['title']}</span>
                                <span class="badge badge-primary">{mod['badge']}</span>
                            </div>
                            <div style="font-size:0.83rem; color:var(--text-secondary); line-height:1.42;">{mod['desc']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button("Abrir módulo", key=f"lobby_btn_{mod['id']}", use_container_width=True):
                        st.session_state.current_module = mod["id"]
                        st.rerun()


# ==============================================================================
# 6. MÓDULO 1: BUSCADOR 360° Y HOJA DE VIDA (CON CORRECCIÓN DE VIDAS)
# ==============================================================================
def render_module_1_buscador_360() -> None:
    def _open_new():
        st.session_state.current_module = "MOD_5"
        st.session_state.show_add_panel = True
        st.rerun()

    render_section_header(
        "🔍 Buscador 360° y Hoja de Vida",
        "Explora el inventario en tarjetas o tabla, revisa KPIs por vida y consulta las 2 últimas fotografías.",
        primary_label="➕ Agregar componente" if st.session_state.user_role in ("Planner", "Administrador") else None,
        primary_action=_open_new if st.session_state.user_role in ("Planner", "Administrador") else None,
    )

    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
    if df_comp.empty:
        conn.close()
        st.warning("No hay componentes registrados en la base de datos.")
        return

    # Filtros y selector de vista
    f1, f2, f3, f4, f5 = st.columns([1.1, 1.1, 1.1, 1.8, 0.9])
    with f1:
        sel_flota = st.selectbox("Flota:", ["Todas"] + sorted(df_comp["flota"].dropna().unique().tolist()))
    with f2:
        df_sub = df_comp if sel_flota == "Todas" else df_comp[df_comp["flota"] == sel_flota]
        sel_equipo = st.selectbox("Equipo Principal:", ["Todos"] + sorted(df_sub["equipo_principal"].dropna().unique().tolist()))
    with f3:
        sel_estado = st.selectbox("Estado:", ["Todos"] + ESTADOS_COMPONENTE)
    with f4:
        texto_busqueda = st.text_input(
            "Filtro de texto:",
            value=st.session_state.global_search_query,
            placeholder="Código SAP, serie, descripción o fabricante...",
        )
        if texto_busqueda != st.session_state.global_search_query:
            st.session_state.global_search_query = texto_busqueda
    with f5:
        vista_m1 = st.selectbox("Vista:", ["Tabla", "Tarjetas"], index=0 if st.session_state.mod1_view_mode == "Tabla" else 1)
        st.session_state.mod1_view_mode = vista_m1

    df_filtered = df_comp.copy()
    if sel_flota != "Todas":
        df_filtered = df_filtered[df_filtered["flota"] == sel_flota]
    if sel_equipo != "Todos":
        df_filtered = df_filtered[df_filtered["equipo_principal"] == sel_equipo]
    if sel_estado != "Todos":
        df_filtered = df_filtered[df_filtered["estado_actual"] == sel_estado]
    if texto_busqueda.strip():
        q = texto_busqueda.strip().lower()
        mask = (
            df_filtered["codigo_equipo"].astype(str).str.lower().str.contains(q, na=False)
            | df_filtered["descripcion"].astype(str).str.lower().str.contains(q, na=False)
            | df_filtered["serie"].astype(str).str.lower().str.contains(q, na=False)
            | df_filtered["fabricante"].astype(str).str.lower().str.contains(q, na=False)
            | df_filtered["ubicacion_tecnica"].astype(str).str.lower().str.contains(q, na=False)
            | df_filtered["equipo_principal"].astype(str).str.lower().str.contains(q, na=False)
        )
        df_filtered = df_filtered[mask]

    st.caption(f"**{len(df_filtered)}** componente(s) encontrados.")

    if df_filtered.empty:
        conn.close()
        return

    if vista_m1 == "Tabla":
        df_vista = df_filtered[
            [
                "codigo_equipo", "descripcion", "flota", "equipo_principal",
                "estado_actual", "ciclo_actual", "horas_ciclo_actual",
                "horas_acumuladas_totales", "gasto_ciclo_actual_usd",
                "gasto_acumulado_total_usd", "ultima_actualizacion_horas",
            ]
        ].copy()
        df_vista["vidas_texto"] = df_vista["ciclo_actual"].apply(format_vidas_label)
        df_vista["costo_por_hora_usd"] = df_vista.apply(
            lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2)
            if r["horas_acumuladas_totales"] > 0 else 0.0,
            axis=1,
        )
        st.dataframe(
            df_vista[
                [
                    "codigo_equipo", "descripcion", "flota", "equipo_principal",
                    "estado_actual", "vidas_texto", "horas_ciclo_actual",
                    "horas_acumuladas_totales", "gasto_ciclo_actual_usd",
                    "gasto_acumulado_total_usd", "costo_por_hora_usd",
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "codigo_equipo": "Código SAP",
                "descripcion": "Descripción",
                "flota": "Flota",
                "equipo_principal": "Equipo",
                "estado_actual": "Estado",
                "vidas_texto": "Vidas",
                "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Vida Actual", format="%.1f hrs"),
                "horas_acumuladas_totales": st.column_config.NumberColumn("Hrs Totales", format="%.1f hrs"),
                "gasto_ciclo_actual_usd": st.column_config.NumberColumn("Costo Vida ($)", format="$%.2f"),
                "gasto_acumulado_total_usd": st.column_config.NumberColumn("Costo Acum. ($)", format="$%.2f"),
                "costo_por_hora_usd": st.column_config.NumberColumn("Costo/Hr", format="$%.2f/hr"),
            },
        )
    else:
        cards_html = ['<div class="responsive-card-grid">']
        for _, r in df_filtered.head(24).iterrows():
            vidas_lbl = format_vidas_label(int(r["ciclo_actual"] or 1))
            c_hr = (r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"]) if r["horas_acumuladas_totales"] > 0 else 0.0
            cards_html.append(
                f"""
                <div class="mac-glass-card">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.35rem;">
                        <span style="font-size:0.75rem; font-weight:700; color:var(--text-muted);">{r['codigo_equipo']} · {r['flota']} / {r['equipo_principal']}</span>
                        {get_status_badge_html(r['estado_actual'])}
                    </div>
                    <div style="font-weight:700; font-size:0.95rem; margin-bottom:0.45rem;">{r['descripcion']}</div>
                    <div style="display:flex; justify-content:space-between; font-size:0.8rem; color:var(--text-secondary);">
                        <span>🔄 <b>{vidas_lbl}</b></span>
                        <span>⏱️ {float(r['horas_ciclo_actual']):,.0f} hrs</span>
                        <span>💵 ${c_hr:,.2f}/hr</span>
                    </div>
                </div>
                """
            )
        cards_html.append("</div>")
        st.markdown("".join(cards_html), unsafe_allow_html=True)

    st.markdown("---")

    opciones_comp = [
        f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']}) [{r['estado_actual']}]"
        for _, r in df_filtered.iterrows()
    ]
    codigos_lista = df_filtered["codigo_equipo"].tolist()
    default_idx = codigos_lista.index(st.session_state.selected_component) if st.session_state.selected_component in codigos_lista else 0

    seleccion_str = st.selectbox("Seleccionar Componente para Hoja de Vida 360°:", opciones_comp, index=default_idx)
    cod_sel = seleccion_str.split(" | ")[0].strip()

    if st.session_state.vida_comp_ref != cod_sel:
        st.session_state.vida_comp_ref = cod_sel
        comp_tmp = df_filtered[df_filtered["codigo_equipo"] == cod_sel].iloc[0]
        st.session_state.vida_idx = max(1, int(comp_tmp["ciclo_actual"] or 1))

    st.session_state.selected_component = cod_sel
    comp = df_filtered[df_filtered["codigo_equipo"] == cod_sel].iloc[0]

    total_vidas = max(0, int(comp["ciclo_actual"] or 1))
    nav_state = get_vidas_navigation_state(total_vidas, st.session_state.vida_idx)

    # Ficha Maestra + Control de Vidas (1 vida sin navegación vs varias vidas con Anterior/Siguiente)
    c_info, c_vidas = st.columns([3.2, 1.4])
    with c_info:
        st.markdown(
            f"""
            <div class="mac-glass-card">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:0.5rem;">
                    <div>
                        <span style="font-size:0.78rem; color:var(--text-muted); font-weight:700;">CÓDIGO SAP: {comp['codigo_equipo']} · SERIE: {comp['serie']}</span>
                        <h3 style="margin:0.15rem 0 0.3rem 0;">{comp['descripcion']}</h3>
                        <span style="font-size:0.85rem; color:var(--text-secondary);">
                            <b>Flota:</b> {comp['flota']} · <b>Equipo:</b> {comp['equipo_principal']} · <b>Fabricante:</b> {comp['fabricante']} · <b>ABC:</b> {comp['criticidad_abc']}
                        </span>
                    </div>
                    <div style="text-align:right;">
                        {get_status_badge_html(comp['estado_actual'])}
                        <div style="font-size:0.76rem; color:var(--text-muted); margin-top:0.35rem;">
                            Últ. Carga: {comp['ultima_actualizacion_horas'] or 'Sin registro'}
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c_vidas:
        st.markdown(
            f"""
            <div class="mac-glass-card" style="text-align:center;">
                <div class="kpi-label">Historial de Vidas</div>
                <div class="kpi-num" style="color:var(--accent);">{nav_state['label']}</div>
                <div class="kpi-meta">{'Vida única registrada' if not nav_state['show_nav'] else f"Viendo vida {nav_state['current']} de {total_vidas}"}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        # Regla estricta de Corrección de Vidas:
        # Con 1 vida -> ocultar controles de navegación innecesarios.
        # Con varias vidas -> permitir navegación, desactivando 'Anterior' o 'Siguiente' cuando corresponda.
        if nav_state["show_nav"]:
            b_prev, b_next = st.columns(2)
            with b_prev:
                if st.button("◀ Anterior", key="vida_prev_btn", disabled=nav_state["disable_prev"], use_container_width=True):
                    st.session_state.vida_idx = max(1, nav_state["current"] - 1)
                    st.rerun()
            with b_next:
                if st.button("Siguiente ▶", key="vida_next_btn", disabled=nav_state["disable_next"], use_container_width=True):
                    st.session_state.vida_idx = min(total_vidas, nav_state["current"] + 1)
                    st.rerun()

    # KPIs de la Hoja de Vida
    hrs_ciclo = float(comp["horas_ciclo_actual"] or 0.0)
    hrs_total = float(comp["horas_acumuladas_totales"] or 0.0)
    costo_ciclo = float(comp["gasto_ciclo_actual_usd"] or 0.0)
    costo_total = float(comp["gasto_acumulado_total_usd"] or 0.0)
    costo_por_hr = (costo_total / hrs_total) if hrs_total > 0 else 0.0
    costo_por_hr_ciclo = (costo_ciclo / hrs_ciclo) if hrs_ciclo > 0 else 0.0

    st.markdown(
        f"""
        <div class="responsive-kpi-grid">
            <div class="mac-glass-card">
                <div class="kpi-label">Horas Vida Actual</div>
                <div class="kpi-num" style="color:var(--accent);">{hrs_ciclo:,.1f} hrs</div>
                <div class="kpi-meta">{nav_state['label']}</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Horas Totales</div>
                <div class="kpi-num">{hrs_total:,.1f} hrs</div>
                <div class="kpi-meta">Acumulado histórico</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Costo Vida Actual</div>
                <div class="kpi-num" style="color:var(--warning);">${costo_ciclo:,.2f}</div>
                <div class="kpi-meta">Kits + Reparaciones</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Costo Acumulado</div>
                <div class="kpi-num" style="color:var(--success);">${costo_total:,.2f}</div>
                <div class="kpi-meta">Inversión total USD</div>
            </div>
            <div class="mac-glass-card">
                <div class="kpi-label">Costo por Hora</div>
                <div class="kpi-num" style="color:var(--danger);">${costo_por_hr:,.2f}/hr</div>
                <div class="kpi-meta">Actual: ${costo_por_hr_ciclo:,.2f}/hr</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 2 Últimas Fotos
    st.markdown("#### 📸 Registro Visual (2 Últimas Fotos - FIFO)")
    p1, p2 = st.columns(2)
    with p1:
        st.markdown('<div class="mac-glass-card">', unsafe_allow_html=True)
        st.markdown("**Fotografía #1 (Previa)**")
        if comp["foto_1_path"] and os.path.exists(comp["foto_1_path"]):
            st.image(comp["foto_1_path"], use_container_width=True)
            st.caption(f"📅 {comp['foto_1_fecha'] or 'S/F'} — {comp['foto_1_comentario'] or ''}")
        else:
            st.info("Espacio #1 disponible.")
        st.markdown("</div>", unsafe_allow_html=True)

    with p2:
        st.markdown('<div class="mac-glass-card">', unsafe_allow_html=True)
        st.markdown("**Fotografía #2 (Más Reciente)**")
        if comp["foto_2_path"] and os.path.exists(comp["foto_2_path"]):
            st.image(comp["foto_2_path"], use_container_width=True)
            st.caption(f"📅 {comp['foto_2_fecha'] or 'S/F'} — {comp['foto_2_comentario'] or ''}")
        else:
            st.info("Espacio #2 disponible.")
        st.markdown("</div>", unsafe_allow_html=True)

    df_eventos = pd.read_sql_query(
        "SELECT * FROM historial_eventos WHERE codigo_equipo = ? ORDER BY id_evento DESC",
        conn,
        params=(cod_sel,),
    )
    conn.close()

    if nav_state["show_nav"] and not df_eventos.empty:
        df_eventos_vida = df_eventos[df_eventos["numero_ciclo"] == nav_state["current"]]
        st.caption(f"Mostrando eventos filtrados para la **vida {nav_state['current']}** ({len(df_eventos_vida)} registros). También puedes consultar el historial completo abajo.")
    else:
        df_eventos_vida = df_eventos

    st.markdown("#### 💳 Desglose Financiero (Kits + Reparación + OC + OT)")
    if not df_eventos_vida.empty:
        df_fin = df_eventos_vida[
            (df_eventos_vida["costo_total_evento_usd"] > 0)
            | (df_eventos_vida["numero_ot"].fillna("") != "")
            | (df_eventos_vida["numero_oc"].fillna("") != "")
            | (df_eventos_vida["tipo_intervencion"].fillna("Ninguna") != "Ninguna")
        ]
        if not df_fin.empty:
            st.dataframe(
                df_fin[
                    [
                        "numero_ciclo", "fecha_hora", "tipo_intervencion", "detalle_kit",
                        "costo_kit_usd", "costo_reparacion_usd", "costo_total_evento_usd",
                        "numero_ot", "numero_oc", "proveedor_reparador", "instalado_por",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Sin costos registrados en esta vida.")

        st.markdown("#### 📓 Bitácora de Eventos y Comentarios")
        st.dataframe(
            df_eventos_vida[
                [
                    "id_evento", "fecha_hora", "numero_ciclo", "usuario_rol",
                    "tipo_operacion", "estado_asignado", "equipo_principal",
                    "horas_componente_evento", "horas_sumadas", "motivo_cambio",
                    "modo_falla", "comentario_bitacora",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No hay eventos registrados para este componente o vida seleccionada.")


# ==============================================================================
# 7. MÓDULO 2: CARGA MASIVA DE HORAS (COMPONENTES MONTADOS)
# ==============================================================================
def render_module_2_carga_horas() -> None:
    render_section_header(
        "⏱️ Carga Masiva de Horas (Componentes Montados)",
        "Suma horas simultáneamente a Horas de Vida/Ciclo Actual y Horas Acumuladas Totales.",
    )

    # Acciones ocasionales agrupadas en menú de tres puntos
    c_info, c_more = st.columns([4.2, 0.8])
    with c_info:
        st.markdown(
            '<div class="mac-glass-card">Solo se actualizan componentes con <code>estado_actual == "Montado"</code>.</div>',
            unsafe_allow_html=True,
        )
    with c_more:
        with show_popover_or_expander("⋮ Más"):
            df_template_hrs = pd.DataFrame(
                [
                    {"Equipo_Principal": "SHO009", "Horas_a_Sumar": 145.0},
                    {"Equipo_Principal": "SHO008", "Horas_a_Sumar": 132.5},
                    {"Equipo_Principal": "FLO005", "Horas_a_Sumar": 118.0},
                ]
            )
            buf_tpl = io.BytesIO()
            with pd.ExcelWriter(buf_tpl, engine="openpyxl") as writer:
                df_template_hrs.to_excel(writer, index=False, sheet_name="Carga_Horas")
            buf_tpl.seek(0)
            st.download_button(
                "📥 Plantilla Excel Horas",
                data=buf_tpl.getvalue(),
                file_name="Plantilla_Carga_Horas_Montados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    tab_a, tab_b, tab_log = st.tabs(
        [
            "📂 Por Archivo Excel / CSV",
            "⚡ Por Flota / Equipo en Pantalla",
            "📜 Historial de Cargas",
        ]
    )

    with tab_a:
        archivo_horas = st.file_uploader(
            "Subir Excel (.xlsx) o CSV (.csv) con [Equipo_Principal o Codigo_Componente, Horas_a_Sumar o Horometro_Actual]:",
            type=["xlsx", "xls", "csv"],
            key="uploader_horas_montados",
        )
        if archivo_horas is not None:
            try:
                df_up = pd.read_csv(archivo_horas) if archivo_horas.name.lower().endswith(".csv") else pd.read_excel(archivo_horas)
                st.dataframe(df_up.head(15), use_container_width=True, hide_index=True)

                cols_norm = {c.strip().lower(): c for c in df_up.columns}
                target_col = cols_norm.get("equipo_principal") or cols_norm.get("codigo_componente") or cols_norm.get("codigo_equipo")
                target_type = "equipo_principal" if "equipo_principal" in cols_norm else "codigo_equipo"
                value_col = cols_norm.get("horas_a_sumar") or cols_norm.get("horometro_actual")
                mode_op = "add" if "horas_a_sumar" in cols_norm else "set_cycle"

                if not target_col or not value_col:
                    st.error("El archivo debe contener 'Equipo_Principal' (o 'Codigo_Componente') y 'Horas_a_Sumar' (o 'Horometro_Actual').")
                else:
                    if st.button("Actualizar Horas de Componentes Montados", type="primary", use_container_width=True):
                        updates_list = []
                        errores_fila = []
                        for idx_r, r in df_up.iterrows():
                            val_id = r[target_col]
                            ok_n, num_h, err_n = validate_numeric_input(
                                "" if pd.isna(r[value_col]) else str(r[value_col]),
                                field_name=f"Fila {idx_r + 1}",
                                allow_zero=False,
                                allow_negative=False,
                                allow_decimal=True,
                            )
                            if pd.notna(val_id) and ok_n:
                                updates_list.append({"target_type": target_type, "target_value": str(val_id).strip(), "mode": mode_op, "value": num_h})
                            elif pd.notna(val_id) and not ok_n:
                                errores_fila.append(err_n)

                        if errores_fila:
                            st.warning(f"Se omitieron {len(errores_fila)} fila(s) con valores inválidos (cero, negativos, vacíos o texto).")
                        total_upd, detalle_upd = apply_hours_to_mounted_components(
                            updates_list, usuario_rol=st.session_state.user_role, metodo_carga=f"Archivo ({archivo_horas.name})"
                        )
                        if total_upd > 0:
                            st.success(f"Se actualizaron {total_upd} componente(s) montado(s).")
                            st.dataframe(pd.DataFrame(detalle_upd), use_container_width=True, hide_index=True)
                        else:
                            st.warning("No se actualizaron componentes (verifica que estén en estado 'Montado').")
            except Exception as exc:
                st.error(f"Error al leer el archivo: {exc}")

    with tab_b:
        conn = get_connection()
        df_montados = pd.read_sql_query(
            """
            SELECT codigo_equipo, descripcion, flota, equipo_principal,
                   ciclo_actual, horas_ciclo_actual, horas_acumuladas_totales,
                   ultima_actualizacion_horas
            FROM componentes
            WHERE estado_actual = 'Montado'
            ORDER BY flota, equipo_principal, codigo_equipo
            """,
            conn,
        )
        conn.close()

        if df_montados.empty:
            st.warning("No hay componentes en estado 'Montado'.")
        else:
            col_f1, col_f2, col_f3 = st.columns([1.3, 1.3, 1.6])
            with col_f1:
                sel_flota_b = st.selectbox("Flota:", ["Todas"] + sorted(df_montados["flota"].dropna().unique().tolist()), key="flota_op_b")
            with col_f2:
                df_sub_b = df_montados if sel_flota_b == "Todas" else df_montados[df_montados["flota"] == sel_flota_b]
                sel_eq_b = st.selectbox("Equipo Principal:", ["Todos"] + sorted(df_sub_b["equipo_principal"].dropna().unique().tolist()), key="eq_op_b")
            with col_f3:
                horas_bloque_txt = st.text_input("Horas a sumar al bloque (+ hrs):", value="0", key="horas_bloque_txt")

            ok_hb, horas_bloque_val, err_hb = validate_numeric_input(
                horas_bloque_txt, field_name="Horas a sumar", allow_zero=True, allow_negative=False, allow_decimal=True
            )
            if not ok_hb:
                st.error(err_hb)
                horas_bloque_val = 0.0

            df_edit_source = df_sub_b if sel_eq_b == "Todos" else df_sub_b[df_sub_b["equipo_principal"] == sel_eq_b]
            df_editor_input = df_edit_source.copy()
            df_editor_input["Horas_a_Sumar"] = float(horas_bloque_val)

            edited_df = st.data_editor(
                df_editor_input,
                use_container_width=True,
                hide_index=True,
                disabled=[
                    "codigo_equipo", "descripcion", "flota", "equipo_principal",
                    "ciclo_actual", "horas_ciclo_actual", "horas_acumuladas_totales", "ultima_actualizacion_horas",
                ],
                key="editor_horas_montados",
            )

            if st.button("Aplicar Horas a Montados", type="primary", use_container_width=True):
                updates_b = []
                for _, row_ed in edited_df.iterrows():
                    raw_h = row_ed.get("Horas_a_Sumar", 0.0)
                    ok_rh, val_rh, _ = validate_numeric_input(str(raw_h), allow_zero=False, allow_negative=False, allow_decimal=True)
                    if ok_rh and val_rh is not None:
                        updates_b.append({"target_type": "codigo_equipo", "target_value": str(row_ed["codigo_equipo"]), "mode": "add", "value": val_rh})
                if not updates_b:
                    st.warning("Ingresa un valor mayor a 0 en 'Horas_a_Sumar' en al menos un componente.")
                else:
                    n_upd, det_upd = apply_hours_to_mounted_components(
                        updates_b, usuario_rol=st.session_state.user_role, metodo_carga=f"Pantalla Rápida (Flota: {sel_flota_b} | Equipo: {sel_eq_b})"
                    )
                    st.success(f"Horas aplicadas correctamente a {n_upd} componente(s) montado(s).")
                    st.dataframe(pd.DataFrame(det_upd), use_container_width=True, hide_index=True)

    with tab_log:
        conn = get_connection()
        df_log = pd.read_sql_query("SELECT * FROM log_carga_horas ORDER BY id_carga DESC LIMIT 100", conn)
        conn.close()
        st.dataframe(df_log, use_container_width=True, hide_index=True) if not df_log.empty else st.info("Sin registros aún.")


# ==============================================================================
# 8. MÓDULO 3: GESTIÓN DE ESTADOS, KITS, OC Y TRAZABILIDAD
# ==============================================================================
def render_module_3_gestion_estados() -> None:
    render_section_header(
        "⚡ Gestión de Estados, Kits, OC y Trazabilidad",
        "Formulario adaptable al estado elegido: montaje, desmontaje, reparación, órdenes de compra y costos.",
    )

    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
    conn.close()
    if df_comp.empty:
        st.warning("No existen componentes cargados.")
        return

    opciones = [
        f"{r['codigo_equipo']} | {r['descripcion']} | Equipo: {r['equipo_principal']} | Estado: {r['estado_actual']} ({format_vidas_label(r['ciclo_actual'])})"
        for _, r in df_comp.iterrows()
    ]
    sel_comp_str = st.selectbox("Seleccionar Componente:", opciones)
    cod_sel = sel_comp_str.split(" | ")[0].strip()
    comp_row = df_comp[df_comp["codigo_equipo"] == cod_sel].iloc[0]

    st.markdown(
        f"""
        <div class="mac-glass-card">
            <b>{comp_row['codigo_equipo']}</b> — {comp_row['descripcion']} &nbsp;|&nbsp;
            {get_status_badge_html(comp_row['estado_actual'])} &nbsp;|&nbsp;
            <b>Equipo:</b> {comp_row['equipo_principal']} &nbsp;|&nbsp;
            <b>{format_vidas_label(comp_row['ciclo_actual'])}</b> &nbsp;|&nbsp;
            <b>Hrs Ciclo:</b> {float(comp_row['horas_ciclo_actual']):,.1f} hrs
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_est, col_op = st.columns(2)
    with col_est:
        nuevo_estado = st.selectbox(
            "Nuevo Estado a Asignar:",
            options=ESTADOS_COMPONENTE,
            index=ESTADOS_COMPONENTE.index(comp_row["estado_actual"]) if comp_row["estado_actual"] in ESTADOS_COMPONENTE else 0,
        )
    with col_op:
        equipo_destino = st.text_input("Equipo Principal Asociado:", value=str(comp_row["equipo_principal"]))

    instalado_por, reiniciar_ciclo = "N/A", False
    motivo_cambio, modo_falla = "N/A", "N/A"
    numero_ot, numero_oc, proveedor_reparador, tipo_intervencion, detalle_kit = "", "", "", "Ninguna", ""
    costo_kit_txt, costo_rep_txt = "0", "0"

    if nuevo_estado == "Montado":
        m1, m2 = st.columns(2)
        with m1:
            instalado_por = st.radio("¿Quién instaló el componente?", options=["Southern", "Contratista"], horizontal=True)
        with m2:
            reiniciar_ciclo = st.checkbox(
                f"Iniciar nueva vida (incrementar a {int(comp_row['ciclo_actual']) + 1} vidas y reiniciar horas de ciclo a 0)",
                value=(comp_row["estado_actual"] != "Montado"),
            )

    if nuevo_estado in ("Desmontado", "Por Reparar", "Salida por Reclamo", "De Baja"):
        d1, d2 = st.columns(2)
        with d1:
            motivo_cambio = st.selectbox("Motivo del Cambio:", options=["PCR por Horas", "Falla por Condición", "Reclamo Garantía"])
        with d2:
            modo_falla = st.selectbox("Modo de Falla:", options=MODOS_FALLA_COMUNES, index=0 if motivo_cambio == "PCR por Horas" else 1)

    mostrar_reparacion = nuevo_estado in ("Reparado", "Por Reparar", "Salida por Reclamo") or st.checkbox("Agregar costos de Kit / OT / OC", value=False)
    if mostrar_reparacion:
        r1, r2, r3 = st.columns(3)
        with r1:
            numero_ot = st.text_input("N° OT SAP:", placeholder="OT-40125901")
        with r2:
            numero_oc = st.text_input("N° Orden de Compra (OC):", placeholder="4500231099")
        with r3:
            proveedor_reparador = st.text_input("Proveedor / Taller:", placeholder="KOMATSU / FERREYROS / SPCC")

        k1, k2, k3 = st.columns([1.4, 1.3, 1.3])
        with k1:
            tipo_intervencion = st.selectbox("Tipo de Intervención:", options=["Kit + Reparación General", "Kit de Mantenimiento", "Reparación General", "Ninguna"])
        with k2:
            costo_kit_txt = st.text_input("Costo Kit (USD):", value="0")
        with k3:
            costo_rep_txt = st.text_input("Costo Reparación General (USD):", value="0")
        detalle_kit = st.text_input("Detalle de Kit / Alcance Técnico:", placeholder="Sellos, rodamientos, rectificado...")

    ok_kit, costo_kit_usd, err_kit = validate_numeric_input(costo_kit_txt, "Costo Kit (USD)", allow_zero=True, allow_negative=False, allow_decimal=True)
    ok_rep, costo_rep_usd, err_rep = validate_numeric_input(costo_rep_txt, "Costo Reparación General (USD)", allow_zero=True, allow_negative=False, allow_decimal=True)
    if not ok_kit:
        st.error(err_kit)
    if not ok_rep:
        st.error(err_rep)

    col_com, col_fot = st.columns([1.5, 1.3])
    with col_com:
        comentario_bitacora = st.text_area("Comentario en Bitácora:", height=95)
    with col_fot:
        foto_evento = st.file_uploader("Fotografía del Evento (FIFO 2 fotos):", type=["jpg", "jpeg", "png", "webp"], key="foto_mod3")

    # Acción destructiva separada con advertencia y confirmación explícita si elige "De Baja"
    confirm_baja = True
    if nuevo_estado == "De Baja":
        st.markdown(
            """
            <div class="danger-zone">
                <b>⚠️ Acción Destructiva:</b> Dar de baja técnica retirará el componente de la operación activa.
            </div>
            """,
            unsafe_allow_html=True,
        )
        confirm_baja = st.checkbox("Confirmo que deseo asignar el estado 'De Baja' a este componente.", key="confirm_baja_chk")

    if st.button("Guardar", type="primary", use_container_width=True, disabled=(not ok_kit or not ok_rep or not confirm_baja)):
        costo_total_evento = round(float(costo_kit_usd or 0.0) + float(costo_rep_usd or 0.0), 2)
        conn = get_connection()
        cur = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        ciclo_actual_db = int(comp_row["ciclo_actual"] or 1)
        horas_ciclo_db = float(comp_row["horas_ciclo_actual"] or 0.0)
        gasto_ciclo_db = float(comp_row["gasto_ciclo_actual_usd"] or 0.0)
        gasto_total_db = float(comp_row["gasto_acumulado_total_usd"] or 0.0)

        if nuevo_estado == "Montado" and reiniciar_ciclo:
            nuevo_ciclo_num = ciclo_actual_db + 1
            nuevas_horas_ciclo = 0.0
            nuevo_gasto_ciclo = costo_total_evento
        else:
            nuevo_ciclo_num = ciclo_actual_db
            nuevas_horas_ciclo = horas_ciclo_db
            nuevo_gasto_ciclo = gasto_ciclo_db + costo_total_evento

        nuevo_gasto_total = gasto_total_db + costo_total_evento

        cur.execute(
            """
            UPDATE componentes
            SET estado_actual = ?, equipo_principal = ?, ciclo_actual = ?,
                horas_ciclo_actual = ?, gasto_ciclo_actual_usd = ?, gasto_acumulado_total_usd = ?
            WHERE codigo_equipo = ?
            """,
            (nuevo_estado, equipo_destino.strip(), nuevo_ciclo_num, round(nuevas_horas_ciclo, 2), round(nuevo_gasto_ciclo, 2), round(nuevo_gasto_total, 2), cod_sel),
        )

        saved_photo_path = None
        msg_fifo = ""
        if foto_evento is not None:
            conn.commit()
            conn.close()
            _, msg_fifo, saved_photo_path = save_component_photo_fifo(
                codigo_equipo=cod_sel,
                uploaded_file=foto_evento,
                comentario_foto=comentario_bitacora or f"Cambio de estado a {nuevo_estado}",
                usuario_rol=st.session_state.user_role,
                registrar_en_bitacora=False,
            )
            conn = get_connection()
            cur = conn.cursor()

        cur.execute(
            """
            INSERT INTO historial_eventos (
                codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                tipo_operacion, estado_asignado, equipo_principal,
                horas_componente_evento, horas_sumadas, instalado_por,
                motivo_cambio, modo_falla, numero_ot, numero_oc,
                proveedor_reparador, tipo_intervencion, detalle_kit,
                costo_kit_usd, costo_reparacion_usd, costo_total_evento_usd,
                comentario_bitacora, foto_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                cod_sel, nuevo_ciclo_num, now_str, st.session_state.user_role,
                f"Cambio de Estado / Gestión ({nuevo_estado})", nuevo_estado, equipo_destino.strip(),
                round(nuevas_horas_ciclo, 2), instalado_por, motivo_cambio, modo_falla,
                numero_ot.strip(), numero_oc.strip(), proveedor_reparador.strip(),
                tipo_intervencion, detalle_kit.strip(), round(float(costo_kit_usd or 0.0), 2),
                round(float(costo_rep_usd or 0.0), 2), costo_total_evento,
                (comentario_bitacora or f"Actualización a estado {nuevo_estado}").strip(), saved_photo_path,
            ),
        )
        conn.commit()
        conn.close()
        st.success(f"Componente `{cod_sel}` guardado como **{nuevo_estado}** ({format_vidas_label(nuevo_ciclo_num)}). {msg_fifo}")


# ==============================================================================
# 9. MÓDULO 4: MODO TABLET CAMPO (FOTOS FIFO Y COMENTARIOS)
# ==============================================================================
def render_module_4_tablet_campo() -> None:
    render_section_header(
        "📱 Modo Tablet Campo: Fotos y Comentarios",
        "Inspección visual rápida en terreno con rotación estricta FIFO de las 2 últimas fotografías.",
    )

    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY equipo_principal, codigo_equipo", conn)
    conn.close()
    if df_comp.empty:
        st.warning("No hay componentes disponibles.")
        return

    c1, c2 = st.columns([1.2, 2.0])
    with c1:
        eq_tablet = st.selectbox("Equipo Principal:", ["Todos"] + sorted(df_comp["equipo_principal"].dropna().unique().tolist()))
    with c2:
        df_t = df_comp if eq_tablet == "Todos" else df_comp[df_comp["equipo_principal"] == eq_tablet]
        sel_t = st.selectbox(
            "Componente:",
            [f"{r['codigo_equipo']} — {r['descripcion']} ({r['equipo_principal']}) [{r['estado_actual']}]" for _, r in df_t.iterrows()],
        )

    cod_tablet = sel_t.split(" — ")[0].strip()
    comp_t = df_comp[df_comp["codigo_equipo"] == cod_tablet].iloc[0]

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        st.markdown('<div class="mac-glass-card"><b>🖼️ Foto #1 Actual</b>', unsafe_allow_html=True)
        if comp_t["foto_1_path"] and os.path.exists(comp_t["foto_1_path"]):
            st.image(comp_t["foto_1_path"], use_container_width=True)
            st.caption(f"📅 {comp_t['foto_1_fecha']} | {comp_t['foto_1_comentario']}")
        else:
            st.info("Espacio #1 libre")
        st.markdown("</div>", unsafe_allow_html=True)
    with col_a2:
        st.markdown('<div class="mac-glass-card"><b>🖼️ Foto #2 Actual (Más Reciente)</b>', unsafe_allow_html=True)
        if comp_t["foto_2_path"] and os.path.exists(comp_t["foto_2_path"]):
            st.image(comp_t["foto_2_path"], use_container_width=True)
            st.caption(f"📅 {comp_t['foto_2_fecha']} | {comp_t['foto_2_comentario']}")
        else:
            st.info("Espacio #2 libre")
        st.markdown("</div>", unsafe_allow_html=True)

    metodo = st.radio("Origen:", ["📁 Galería / Archivo", "📸 Cámara"], horizontal=True)
    img_input = st.file_uploader("Seleccionar foto:", type=["jpg", "jpeg", "png", "webp"], key="tablet_uploader") if metodo.startswith("📁") else st.camera_input("Capturar foto:")
    comentario_campo = st.text_area("Comentario de inspección:", height=90)

    b1, b2 = st.columns(2)
    with b1:
        if st.button("Guardar Foto (FIFO) y Comentario", type="primary", use_container_width=True):
            if img_input is None:
                st.error("Selecciona o captura una fotografía primero.")
            else:
                ok, msg, _ = save_component_photo_fifo(cod_tablet, img_input, comentario_campo, st.session_state.user_role, True)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    with b2:
        if st.button("Solo Guardar Comentario", type="secondary", use_container_width=True):
            if not comentario_campo.strip():
                st.warning("Escribe un comentario antes de guardar.")
            else:
                conn = get_connection()
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT INTO historial_eventos (
                        codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                        tipo_operacion, estado_asignado, equipo_principal,
                        horas_componente_evento, comentario_bitacora
                    ) VALUES (?, ?, ?, ?, 'Nota de Campo (Tablet)', ?, ?, ?, ?)
                    """,
                    (cod_tablet, comp_t["ciclo_actual"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), st.session_state.user_role, comp_t["estado_actual"], comp_t["equipo_principal"], comp_t["horas_ciclo_actual"], comentario_campo.strip()),
                )
                conn.commit()
                conn.close()
                st.success("Comentario guardado en bitácora.")


# ==============================================================================
# 10. MÓDULO 5: INGRESO DE COMPONENTES (PANEL DERECHO), SAP Y BACKUP
# ==============================================================================
def render_module_5_alta_sap_backup() -> None:
    def _toggle_panel():
        st.session_state.show_add_panel = not st.session_state.show_add_panel
        st.rerun()

    render_section_header(
        "➕ Ingreso de Componentes, Sincronización SAP y Backup",
        "Gestiona el maestro de componentes, sincroniza con Excel SAP PM y descarga respaldos.",
        primary_label="✕ Cerrar panel" if st.session_state.show_add_panel else "➕ Agregar componente",
        primary_action=_toggle_panel,
    )

    # Barra de herramientas secundarias y menú de tres puntos
    c_view, c_dl1, c_dl2, c_more = st.columns([1.2, 1.8, 1.8, 0.8])
    timestamp_dl = datetime.now().strftime("%Y%m%d_%H%M")
    with c_view:
        v5 = st.selectbox("Vista:", ["Tarjetas", "Tabla"], index=0 if st.session_state.mod5_view_mode == "Tarjetas" else 1, key="v5_sel", label_visibility="collapsed")
        st.session_state.mod5_view_mode = v5
    with c_dl1:
        st.download_button(
            "📥 Descargar Base de Datos (.xlsx)",
            data=export_full_database_to_excel(),
            file_name=f"PIM_Backup_Inventario_SAP_{timestamp_dl}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c_dl2:
        if os.path.exists(DB_PATH):
            with open(DB_PATH, "rb") as f_db:
                st.download_button(
                    "🗄️ Descargar SQLite (.db)",
                    data=f_db.read(),
                    file_name=f"inventario_mina_{timestamp_dl}.db",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
    with c_more:
        with show_popover_or_expander("⋮"):
            st.download_button(
                "📄 Plantilla SAP Nov.xlsx",
                data=generate_sap_template_excel(),
                file_name="Inventario SAP Nov.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # Distribución adaptable: Panel lateral derecho en escritorio para ingreso de componentes
    if st.session_state.show_add_panel:
        col_list, col_panel = st.columns([1.35, 1.0])
    else:
        col_list = st.container()
        col_panel = None

    with col_list:
        conn = get_connection()
        df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
        conn.close()

        if st.session_state.mod5_view_mode == "Tarjetas":
            html_cards = ['<div class="responsive-card-grid">']
            for _, r in df_comp.iterrows():
                html_cards.append(
                    f"""
                    <div class="mac-glass-card">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.35rem;">
                            <span style="font-size:0.75rem; font-weight:700; color:var(--text-muted);">{r['codigo_equipo']}</span>
                            {get_status_badge_html(r['estado_actual'])}
                        </div>
                        <div style="font-weight:700; font-size:0.92rem; margin-bottom:0.35rem;">{r['descripcion']}</div>
                        <div style="font-size:0.8rem; color:var(--text-secondary);">
                            {r['flota']} · {r['equipo_principal']} · <b>{format_vidas_label(r['ciclo_actual'])}</b> · {float(r['horas_ciclo_actual']):,.0f} hrs
                        </div>
                    </div>
                    """
                )
            html_cards.append("</div>")
            st.markdown("".join(html_cards), unsafe_allow_html=True)
        else:
            st.dataframe(
                df_comp[["codigo_equipo", "descripcion", "flota", "equipo_principal", "estado_actual", "ciclo_actual", "horas_ciclo_actual"]],
                use_container_width=True,
                hide_index=True,
            )

    if col_panel is not None:
        with col_panel:
            st.markdown('<div class="side-form-panel">', unsafe_allow_html=True)
            st.markdown("#### ➕ Ingreso de Componente")

            # Agrupado por: información principal, categoría, cantidad, vidas, características e información adicional
            st.markdown("**1. Información principal**")
            cod_m = st.text_input("Código Equipo SAP *:", placeholder="Ej. 10045990", key="alta_cod")
            desc_m = st.text_input("Descripción del Componente *:", placeholder="Ej. MOTOR DE GIRO TRASERO", key="alta_desc")

            st.markdown("**2. Categoría**")
            ubic_m = st.text_input("Ubicación Técnica SAP *:", value="SPCC-TOQ-MIN-495H-SHO009-GIRO", key="alta_ubic")
            flota_prev, eq_prev = parse_ubicacion_tecnica(ubic_m)
            st.caption(f"Flota detectada: **{flota_prev}** · Equipo Principal: **{eq_prev}**")
            crit_m = st.selectbox("Criticidad ABC:", options=["A", "B", "C"], index=0, key="alta_crit")

            st.markdown("**3. Cantidad (Horómetro e Inversión)**")
            c_q1, c_q2 = st.columns(2)
            with c_q1:
                hrs_ini_txt = st.text_input("Horas Iniciales:", value="0", key="alta_hrs")
            with c_q2:
                costo_ini_txt = st.text_input("Costo Inicial (USD):", value="0", key="alta_costo")

            st.markdown("**4. Vidas**")
            vidas_ini_txt = st.text_input("Vidas iniciales (entero >= 1):", value="1", key="alta_vidas")
            ok_v, v_num, err_v = validate_numeric_input(vidas_ini_txt, "Vidas", allow_zero=False, allow_negative=False, allow_decimal=False)
            if ok_v and v_num is not None:
                v_state = get_vidas_navigation_state(int(v_num), 1)
                st.caption(f"Registro inicial: **{v_state['label']}**")
            else:
                st.error(err_v)

            st.markdown("**5. Características**")
            c_c1, c_c2 = st.columns(2)
            with c_c1:
                fab_m = st.text_input("Fabricante:", placeholder="KOMATSU / CAT", key="alta_fab")
            with c_c2:
                serie_m = st.text_input("N° Serie:", placeholder="SN-884920", key="alta_serie")

            st.markdown("**6. Información adicional**")
            est_m = st.selectbox("Estado Inicial:", options=ESTADOS_COMPONENTE, index=0, key="alta_est")

            ok_h, hrs_val, err_h = validate_numeric_input(hrs_ini_txt, "Horas Iniciales", allow_zero=True, allow_negative=False, allow_decimal=True)
            ok_c, costo_val, err_c = validate_numeric_input(costo_ini_txt, "Costo Inicial (USD)", allow_zero=True, allow_negative=False, allow_decimal=True)
            if not ok_h:
                st.error(err_h)
            if not ok_c:
                st.error(err_c)

            btn_cancel, btn_save = st.columns(2)
            with btn_cancel:
                if st.button("Cancelar", type="secondary", use_container_width=True, key="alta_cancelar"):
                    st.session_state.show_add_panel = False
                    st.rerun()
            with btn_save:
                if st.button("Guardar", type="primary", use_container_width=True, key="alta_guardar"):
                    if not cod_m.strip() or not desc_m.strip():
                        st.error("Código Equipo SAP y Descripción son obligatorios.")
                    elif not (ok_h and ok_c and ok_v):
                        st.error("Corrige los campos numéricos antes de guardar.")
                    else:
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        conn = get_connection()
                        cur = conn.cursor()
                        try:
                            cur.execute(
                                """
                                INSERT INTO componentes (
                                    codigo_equipo, descripcion, flota, equipo_principal,
                                    ubicacion_tecnica, criticidad_abc, fabricante, serie,
                                    estado_actual, ciclo_actual, horas_acumuladas_totales,
                                    horas_ciclo_actual, gasto_acumulado_total_usd,
                                    gasto_ciclo_actual_usd, ultima_actualizacion_horas
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    cod_m.strip(), desc_m.strip().upper(), flota_prev, eq_prev,
                                    ubic_m.strip().upper(), crit_m, (fab_m.strip() or "S/I").upper(),
                                    (serie_m.strip() or "S/N").upper(), est_m, int(v_num),
                                    round(float(hrs_val), 2), round(float(hrs_val), 2),
                                    round(float(costo_val), 2), round(float(costo_val), 2), now_str,
                                ),
                            )
                            cur.execute(
                                """
                                INSERT INTO historial_eventos (
                                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                                    tipo_operacion, estado_asignado, equipo_principal,
                                    horas_componente_evento, costo_total_evento_usd, comentario_bitacora
                                ) VALUES (?, ?, ?, ?, 'Alta Inicial de Componente', ?, ?, ?, ?, ?)
                                """,
                                (
                                    cod_m.strip(), int(v_num), now_str, st.session_state.user_role,
                                    est_m, eq_prev, round(float(hrs_val), 2), round(float(costo_val), 2),
                                    f"Alta manual. Flota: {flota_prev} | Equipo: {eq_prev} | {format_vidas_label(int(v_num))}",
                                ),
                            )
                            conn.commit()
                            st.success(f"Componente `{cod_m.strip()}` guardado.")
                            st.session_state.show_add_panel = False
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error(f"El código `{cod_m.strip()}` ya existe.")
                        finally:
                            conn.close()
            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    tab_sap, tab_restore = st.tabs(["🔄 Sincronización SAP ('Inventario SAP Nov.xlsx')", "⚠️ Restaurar Base de Datos (.db)"])
    with tab_sap:
        if os.path.exists(DEFAULT_SAP_EXCEL):
            if st.button("Sincronizar desde 'Inventario SAP Nov.xlsx' local", use_container_width=True):
                try:
                    df_local = pd.read_excel(DEFAULT_SAP_EXCEL, sheet_name="Sheet1")
                    n_new, n_upd, msg = import_sap_dataframe(df_local, st.session_state.user_role)
                    st.success(f"{msg} (Nuevos: {n_new} | Actualizados: {n_upd})")
                except Exception as exc:
                    st.error(f"Error: {exc}")

        sap_file = st.file_uploader("Subir Excel de SAP ('Inventario SAP Nov.xlsx'):", type=["xlsx", "xls"], key="sap_excel_uploader")
        if sap_file is not None:
            xl = openpyxl.load_workbook(sap_file, read_only=True)
            sheet_to_use = "Sheet1" if "Sheet1" in xl.sheetnames else xl.sheetnames[0]
            sap_file.seek(0)
            df_sap_up = pd.read_excel(sap_file, sheet_name=sheet_to_use)
            st.dataframe(df_sap_up.head(10), use_container_width=True, hide_index=True)
            if st.button("Ejecutar Sincronización SAP", type="primary", use_container_width=True):
                n_new, n_upd, msg = import_sap_dataframe(df_sap_up, st.session_state.user_role)
                st.success(f"{msg} — Nuevos: {n_new} | Actualizados: {n_upd}")

    with tab_restore:
        st.markdown(
            """
            <div class="danger-zone">
                <b>⚠️ Zona Destructiva:</b> Reemplazar el archivo <code>inventario_mina.db</code> sobrescribirá los datos actuales en disco. Descarga un respaldo antes de continuar.
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded_db = st.file_uploader("Subir archivo `.db` para restaurar:", type=["db", "sqlite", "sqlite3"], key="restore_db_uploader")
        if uploaded_db is not None:
            confirm_restore = st.checkbox("Confirmo que deseo reemplazar la base de datos actual.", key="confirm_restore_chk")
            if st.button("Reemplazar Base de Datos Actual", type="primary", disabled=not confirm_restore):
                with open(DB_PATH, "wb") as out_db:
                    out_db.write(uploaded_db.getbuffer())
                st.success("Base de datos restaurada.")
                st.rerun()


# ==============================================================================
# 11. MÓDULO 6: ESTRATEGIA Y COSTO POR CICLO (PLOTLY)
# ==============================================================================
def render_module_6_estrategia_costos() -> None:
    render_section_header(
        "📊 Estrategia de Componentes y Costo por Ciclo",
        "Indicadores financieros $/hr por flota, matriz de desgaste y ranking de componentes críticos.",
    )

    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes", conn)
    df_hist = pd.read_sql_query("SELECT * FROM historial_eventos", conn)
    conn.close()
    if df_comp.empty:
        st.warning("No hay datos suficientes.")
        return

    df_comp["costo_por_hora_usd"] = df_comp.apply(
        lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2) if r["horas_acumuladas_totales"] > 0 else 0.0,
        axis=1,
    )

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown('<div class="mac-glass-card"><b>💵 Costo Promedio por Hora ($/hr) por Flota</b>', unsafe_allow_html=True)
        df_flota = df_comp.groupby("flota", as_index=False).agg(gasto_total=("gasto_acumulado_total_usd", "sum"), horas_totales=("horas_acumuladas_totales", "sum"))
        df_flota["costo_hr"] = df_flota.apply(lambda r: round(r["gasto_total"] / r["horas_totales"], 2) if r["horas_totales"] > 0 else 0.0, axis=1)
        fig_flota = px.bar(df_flota, x="flota", y="costo_hr", color="flota", text="costo_hr", color_discrete_sequence=["#E06D53", "#2E9E6B", "#D9822B", "#6B655D", "#4A7C89"])
        fig_flota.update_traces(texttemplate="$%{text:.2f}/hr", textposition="outside")
        fig_flota.update_layout(height=330, margin=dict(l=15, r=15, t=20, b=15), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
        st.plotly_chart(fig_flota, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_g2:
        st.markdown('<div class="mac-glass-card"><b>🎯 Horas Acumuladas vs. Gasto Total</b>', unsafe_allow_html=True)
        fig_scatter = px.scatter(
            df_comp,
            x="horas_acumuladas_totales",
            y="gasto_acumulado_total_usd",
            color="estado_actual",
            size="ciclo_actual",
            hover_name="descripcion",
            color_discrete_sequence=["#2E9E6B", "#E06D53", "#D9822B", "#8E877E", "#D14343"],
        )
        fig_scatter.update_layout(height=330, margin=dict(l=15, r=15, t=20, b=15), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_scatter, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    col_g3, col_g4 = st.columns(2)
    with col_g3:
        st.markdown('<div class="mac-glass-card"><b>🛠️ Inversión en Kits vs. Reparación por Vida</b>', unsafe_allow_html=True)
        if not df_hist.empty and df_hist["costo_total_evento_usd"].sum() > 0:
            df_ciclo = df_hist.groupby("numero_ciclo", as_index=False).agg(Kits=("costo_kit_usd", "sum"), Reparacion=("costo_reparacion_usd", "sum"))
            fig_ciclo = go.Figure()
            fig_ciclo.add_trace(go.Bar(x=df_ciclo["numero_ciclo"].astype(str), y=df_ciclo["Kits"], name="Kits ($)", marker_color="#E06D53"))
            fig_ciclo.add_trace(go.Bar(x=df_ciclo["numero_ciclo"].astype(str), y=df_ciclo["Reparacion"], name="Reparación ($)", marker_color="#D9822B"))
            fig_ciclo.update_layout(barmode="stack", height=320, margin=dict(l=15, r=15, t=20, b=15), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_ciclo, use_container_width=True)
        else:
            st.info("Sin costos históricos suficientes.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_g4:
        st.markdown('<div class="mac-glass-card"><b>🚨 Top Componentes por Costo/Hora ($/hr)</b>', unsafe_allow_html=True)
        df_rank = df_comp.sort_values(by="costo_por_hora_usd", ascending=False).head(10)
        st.dataframe(
            df_rank[["codigo_equipo", "descripcion", "flota", "equipo_principal", "ciclo_actual", "costo_por_hora_usd", "estado_actual"]],
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# 12. PUNTO DE ENTRADA PRINCIPAL
# ==============================================================================
def main() -> None:
    init_ui_session_state()
    inject_ios_macos_css(st.session_state.theme_mode, st.session_state.sidebar_mode)
    init_db()
    render_sidebar()
    render_topbar()

    mod = st.session_state.current_module
    role = st.session_state.user_role

    if role == "Campo-Tablet" and mod in ("MOD_2", "MOD_3", "MOD_5", "MOD_6"):
        st.warning("El rol 'Campo-Tablet' solo tiene acceso a Inicio, Buscador 360° y Modo Tablet Campo.")
        st.session_state.current_module = "LOBBY"
        render_lobby()
        return

    if mod == "LOBBY":
        render_lobby()
    elif mod == "MOD_1":
        render_module_1_buscador_360()
    elif mod == "MOD_2":
        render_module_2_carga_horas()
    elif mod == "MOD_3":
        render_module_3_gestion_estados()
    elif mod == "MOD_4":
        render_module_4_tablet_campo()
    elif mod == "MOD_5":
        render_module_5_alta_sap_backup()
    elif mod == "MOD_6":
        render_module_6_estrategia_costos()
    else:
        render_lobby()


if __name__ == "__main__":
    main()
