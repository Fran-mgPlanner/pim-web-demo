# -*- coding: utf-8 -*-
"""
Sistema P.I.M. - Control de Inventarios | Planeamiento Mantenimiento Mina
Estética: iOS / macOS Warm Glassmorphism (Crema/Arena en Claro · Grafito en Oscuro · Acento Coral/Terracota)
Arquitectura Simplificada: 4 Vistas Directas (Visualizar por Flota, Actualizar + Fotos, Agregar + SAP, Estrategia + Backup)
SQLite3 con modo WAL para concurrencia segura en PC Servidor / Red Local.
"""

import io
import os
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
# 0. CONFIGURACIÓN INICIAL Y CONSTANTES CORPORATIVAS
# ==============================================================================
st.set_page_config(
    page_title="P.I.M. | Control de Inventarios",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
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
# 1. UTILIDADES DE VALIDACIÓN NUMÉRICA, VIDAS Y BADGES DE FLOTA
# ==============================================================================
def format_vidas_label(vidas: int) -> str:
    n = int(vidas) if vidas is not None else 0
    return "1 vida" if n == 1 else f"{n} vidas"


def get_vidas_navigation_state(total_vidas: int, vida_actual: int) -> Dict:
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


def get_flota_badge_html(flota: str) -> str:
    """Indicador visual consistente por flota para distinguirla de inmediato."""
    flota_clean = str(flota or "SIN_FLOTA").strip()
    paleta = {
        "495H": ("rgba(224, 109, 83, 0.16)", "#D95D43"),
        "410A": ("rgba(59, 130, 246, 0.16)", "#2563EB"),
        "994F": ("rgba(46, 158, 107, 0.16)", "#2E9E6B"),
        "793F": ("rgba(217, 130, 43, 0.18)", "#C2701E"),
        "797F": ("rgba(139, 92, 246, 0.16)", "#7C3AED"),
    }
    if flota_clean in paleta:
        bg, fg = paleta[flota_clean]
    else:
        colores_extra = [
            ("rgba(14, 165, 233, 0.16)", "#0284C7"),
            ("rgba(236, 72, 153, 0.16)", "#DB2777"),
            ("rgba(20, 184, 166, 0.16)", "#0D9488"),
        ]
        bg, fg = colores_extra[sum(ord(c) for c in flota_clean) % len(colores_extra)]
    return f'&lt;span class=&quot;flota-pill&quot; style=&quot;background:{bg}; color:{fg}; border:1px solid {fg}33;&quot;&gt;🚜 Flota {flota_clean}&lt;/span&gt;'


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
    return f'&lt;span class=&quot;badge {css_cls}&quot;&gt;{label}&lt;/span&gt;'


# ==============================================================================
# 2. ESTÉTICA iOS / macOS ADAPTABLE (ESCRITORIO, TABLET Y MÓVIL)
# ==============================================================================
def inject_ios_macos_css(theme_mode: str) -> None:
    light_vars = """
        --bg-app: #FBF8F3;
        --bg-mesh-1: rgba(224, 109, 83, 0.08);
        --bg-mesh-2: rgba(217, 173, 130, 0.10);
        --bg-surface: rgba(255, 253, 249, 0.84);
        --bg-elevated: rgba(243, 236, 226, 0.88);
        --bg-input: rgba(255, 255, 255, 0.94);
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
        --shadow-card: 0 6px 20px rgba(45, 35, 25, 0.06);
    """

    dark_vars = """
        --bg-app: #16171A;
        --bg-mesh-1: rgba(230, 126, 101, 0.09);
        --bg-mesh-2: rgba(180, 130, 95, 0.07);
        --bg-surface: rgba(31, 33, 38, 0.84);
        --bg-elevated: rgba(43, 46, 54, 0.88);
        --bg-input: rgba(38, 40, 47, 0.94);
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
        --shadow-card: 0 8px 24px rgba(0, 0, 0, 0.34);
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

    st.markdown(
        f"""
        <style>
        {theme_css}
        html, body, .stApp {{
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif !important;
            background-color: var(--bg-app) !important;
            background-image:
                radial-gradient(at 10% 8%, var(--bg-mesh-1) 0px, transparent 48%),
                radial-gradient(at 92% 14%, var(--bg-mesh-2) 0px, transparent 45%) !important;
            color: var(--text-primary) !important;
            overflow-x: hidden !important;
        }}
        .block-container {{
            padding-top: 1rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1320px !important;
            overflow-x: hidden !important;
        }}
        p, span, label, div, h1, h2, h3, h4, h5, h6 {{
            color: var(--text-primary);
        }}

        /* Tarjetas Glassmorphism Moderado */
        .mac-glass-card {{
            background: var(--bg-surface);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid var(--border-subtle);
            border-radius: 16px;
            box-shadow: var(--shadow-card);
            padding: 1rem 1.15rem;
            margin-bottom: 0.75rem;
            transition: transform 200ms ease, border-color 200ms ease;
        }}
        .mac-glass-card:hover {{
            border-color: var(--border-strong);
        }}

        /* Cuadrícula Adaptable: 4/3 cols escritorio, 2 tablet, 1 móvil */
        .responsive-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.75rem;
            margin-bottom: 0.9rem;
        }}
        .responsive-card-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.85rem;
            margin-bottom: 0.9rem;
        }}
        @media (max-width: 1024px) {{
            .responsive-kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .responsive-card-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
        }}
        @media (max-width: 640px) {{
            .responsive-kpi-grid, .responsive-card-grid {{ grid-template-columns: 1fr; }}
            .block-container {{ padding-left: 0.75rem !important; padding-right: 0.75rem !important; }}
            div.stButton > button, div.stDownloadButton > button {{
                min-height: 46px !important;
                font-size: 0.95rem !important;
            }}
        }}

        .kpi-label {{
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-secondary);
        }}
        .kpi-num {{
            font-size: 1.45rem;
            font-weight: 800;
            color: var(--text-primary);
            margin-top: 0.15rem;
        }}
        .kpi-meta {{
            font-size: 0.76rem;
            color: var(--text-muted);
            margin-top: 0.15rem;
        }}

        /* Badges de Flota y Estado */
        .flota-pill {{
            display: inline-flex;
            align-items: center;
            padding: 0.22rem 0.6rem;
            font-size: 0.72rem;
            font-weight: 800;
            border-radius: 999px;
            letter-spacing: 0.02em;
        }}
        .badge {{
            display: inline-flex;
            align-items: center;
            padding: 0.22rem 0.6rem;
            font-size: 0.72rem;
            font-weight: 700;
            border-radius: 999px;
        }}
        .badge-success {{ background: var(--success-soft); color: var(--success); }}
        .badge-primary {{ background: var(--accent-soft); color: var(--accent); }}
        .badge-warning {{ background: var(--warning-soft); color: var(--warning); }}
        .badge-danger  {{ background: var(--danger-soft); color: var(--danger); }}
        .badge-neutral {{ background: var(--bg-elevated); color: var(--text-secondary); border: 1px solid var(--border-subtle); }}

        /* Botones limpios estilo iOS */
        div.stButton > button, div.stDownloadButton > button {{
            border-radius: 12px !important;
            font-weight: 600 !important;
            min-height: 40px !important;
            transition: all 180ms ease !important;
        }}
        div.stButton > button[kind="primary"] {{
            background-color: var(--accent) !important;
            border: 1px solid var(--accent) !important;
            color: #FFFFFF !important;
        }}
        div.stButton > button[kind="primary"]:hover {{
            background-color: var(--accent-hover) !important;
            border-color: var(--accent-hover) !important;
        }}
        div.stButton > button[kind="secondary"], div.stDownloadButton > button {{
            background-color: var(--bg-surface) !important;
            color: var(--text-primary) !important;
            border: 1px solid var(--border-strong) !important;
        }}

        /* Inputs */
        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        textarea {{
            background-color: var(--bg-input) !important;
            border-color: var(--border-subtle) !important;
            border-radius: 12px !important;
            color: var(--text-primary) !important;
        }}

        /* Tabs */
        button[data-baseweb="tab"] {{
            border-radius: 10px !important;
            color: var(--text-secondary) !important;
            font-weight: 600 !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            color: var(--accent) !important;
            background-color: var(--accent-soft) !important;
        }}

        .danger-zone {{
            background: var(--danger-soft);
            border: 1px solid var(--danger);
            border-radius: 14px;
            padding: 0.9rem 1.1rem;
            margin: 0.6rem 0;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# 3. PERSISTENCIA SQLITE3 CON MODO WAL (CONCURRENCIA SEGURA EN PC SERVIDOR)
# ==============================================================================
def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
    except sqlite3.Error:
        pass
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
            FROM componentes WHERE codigo_equipo = ?
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

        f1_path, f2_path = row["foto_1_path"], row["foto_2_path"]
        f2_fecha, f2_com = row["foto_2_fecha"], row["foto_2_comentario"]
        comentario_limpio = (comentario_foto or "Inspección visual de componente").strip()

        if not f1_path:
            cur.execute(
                "UPDATE componentes SET foto_1_path = ?, foto_1_fecha = ?, foto_1_comentario = ? WHERE codigo_equipo = ?",
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Foto guardada en Espacio #1."
        elif not f2_path:
            cur.execute(
                "UPDATE componentes SET foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ? WHERE codigo_equipo = ?",
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Foto guardada en Espacio #2."
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
            msg_rotacion = "Rotación FIFO aplicada: se borró la foto #1, la #2 pasó a #1 y la nueva quedó en #2."

        if registrar_en_bitacora:
            cur.execute(
                """
                INSERT INTO historial_eventos (
                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                    tipo_operacion, estado_asignado, equipo_principal,
                    horas_componente_evento, comentario_bitacora, foto_path
                ) VALUES (?, ?, ?, ?, 'Inspección Fotográfica / Bitácora', ?, ?, ?, ?, ?)
                """,
                (codigo_equipo, row["ciclo_actual"], now_display, usuario_rol, row["estado_actual"], row["equipo_principal"], row["horas_ciclo_actual"], comentario_limpio, new_photo_path),
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
    nuevos, actualizados = 0, 0

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
        "INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle) VALUES (?, ?, ?, ?, ?)",
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
                "SELECT codigo_equipo, descripcion, flota, equipo_principal, ciclo_actual, horas_ciclo_actual, horas_acumuladas_totales FROM componentes WHERE UPPER(equipo_principal) = UPPER(?) AND estado_actual = 'Montado'",
                (t_val,),
            )
        else:
            cur.execute(
                "SELECT codigo_equipo, descripcion, flota, equipo_principal, ciclo_actual, horas_ciclo_actual, horas_acumuladas_totales FROM componentes WHERE UPPER(codigo_equipo) = UPPER(?) AND estado_actual = 'Montado'",
                (t_val,),
            )

        for r in cur.fetchall():
            cod = r["codigo_equipo"]
            h_ciclo_ant = float(r["horas_ciclo_actual"] or 0.0)
            h_total_ant = float(r["horas_acumuladas_totales"] or 0.0)
            if mode == "set_cycle":
                delta = val - h_ciclo_ant
                if delta <= 0:
                    continue
                nuevo_ciclo, nuevo_total = val, h_total_ant + delta
            else:
                delta = val
                nuevo_ciclo, nuevo_total = h_ciclo_ant + delta, h_total_ant + delta

            cur.execute(
                "UPDATE componentes SET horas_ciclo_actual = ?, horas_acumuladas_totales = ?, ultima_actualizacion_horas = ? WHERE codigo_equipo = ? AND estado_actual = 'Montado'",
                (round(nuevo_ciclo, 2), round(nuevo_total, 2), now_str, cod),
            )
            cur.execute(
                """
                INSERT INTO historial_eventos (
                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                    tipo_operacion, estado_asignado, equipo_principal,
                    horas_componente_evento, horas_sumadas, comentario_bitacora
                ) VALUES (?, ?, ?, ?, 'Carga de Horas (Montado)', 'Montado', ?, ?, ?, ?)
                """,
                (cod, r["ciclo_actual"], now_str, usuario_rol, r["equipo_principal"], round(nuevo_ciclo, 2), round(delta, 2), f"Carga ({metodo_carga}): +{delta:,.1f} hrs"),
            )
            afectados_detalle.append(
                {
                    "Flota": r["flota"],
                    "Código": cod,
                    "Descripción": r["descripcion"],
                    "Equipo": r["equipo_principal"],
                    "Horas Sumadas (+)": round(delta, 1),
                    "Nuevas Hrs Ciclo": round(nuevo_ciclo, 1),
                    "Nuevas Hrs Totales": round(nuevo_total, 1),
                }
            )

    if afectados_detalle:
        cur.execute(
            "INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle) VALUES (?, ?, ?, ?, ?)",
            (now_str, usuario_rol, metodo_carga, len(afectados_detalle), f"Actualizados {len(afectados_detalle)} componentes montados en todas las flotas."),
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
                ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 45)
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


def generate_hours_template_excel() -> bytes:
    df_template_hrs = pd.DataFrame(
        [
            {"Equipo_Principal": "SHO009", "Horas_a_Sumar": 145.0},
            {"Equipo_Principal": "SHO008", "Horas_a_Sumar": 132.5},
            {"Equipo_Principal": "FLO005", "Horas_a_Sumar": 118.0},
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_template_hrs.to_excel(writer, index=False, sheet_name="Carga_Horas")
    buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# 4. NAVEGACIÓN SIMPLIFICADA Y SELECTOR DE FLOTA (ESPACIO DEL PLANNER)
# ==============================================================================
def init_ui_session_state() -> None:
    defaults = {
        "current_view": "VISUALIZAR",   # VISUALIZAR | ACTUALIZAR | AGREGAR | ESTRATEGIA
        "user_role": "Planner",
        "active_flota": "Todas",
        "selected_component": None,
        "theme_mode": "Sistema",
        "vida_idx": 1,
        "vida_comp_ref": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_compact_header_and_nav() -> None:
    """Cabecera limpia apta para escritorio y móvil, con navegación directa de 4 botones."""
    with st.sidebar:
        st.markdown("### ⚙️ P.I.M.")
        st.caption("Control de Inventarios · Planeamiento Mantenimiento Mina")
        new_role = st.selectbox("Perfil activo:", ROLES_SISTEMA, index=ROLES_SISTEMA.index(st.session_state.user_role))
        if new_role != st.session_state.user_role:
            st.session_state.user_role = new_role
            if new_role == "Campo-Tablet" and st.session_state.current_view not in ("VISUALIZAR", "ACTUALIZAR"):
                st.session_state.current_view = "VISUALIZAR"
            st.rerun()

        new_theme = st.selectbox("Apariencia:", ["Claro", "Oscuro", "Sistema"], index=["Claro", "Oscuro", "Sistema"].index(st.session_state.theme_mode))
        if new_theme != st.session_state.theme_mode:
            st.session_state.theme_mode = new_theme
            st.rerun()

        st.markdown("---")
        st.caption(f"🗄️ Base local: `{os.path.basename(DB_PATH)}`")

    # Cabecera superior compacta
    c_title, c_role, c_theme = st.columns([3.2, 1.0, 1.0])
    with c_title:
        st.markdown(
            """
            <div style="padding-bottom:0.35rem;">
                <span style="font-size:1.35rem; font-weight:800; color:var(--accent);">⚙️ P.I.M.</span>
                <span style="font-size:1.05rem; font-weight:700; color:var(--text-primary); margin-left:0.35rem;">Control de Inventarios</span>
                <div style="font-size:0.78rem; color:var(--text-secondary);">Planeamiento Mantenimiento Mina · Gestión por Flota</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c_role:
        st.markdown(
            f"<div class='mac-glass-card' style='padding:0.45rem 0.7rem; text-align:center; font-size:0.8rem; font-weight:700;'>👤 {st.session_state.user_role}</div>",
            unsafe_allow_html=True,
        )
    with c_theme:
        t_sel = st.selectbox(
            "Tema",
            ["Claro", "Oscuro", "Sistema"],
            index=["Claro", "Oscuro", "Sistema"].index(st.session_state.theme_mode),
            label_visibility="collapsed",
            key="top_theme_sel",
        )
        if t_sel != st.session_state.theme_mode:
            st.session_state.theme_mode = t_sel
            st.rerun()

    # Barra de navegación directa de máximo 4 ventanas (sin exceso de botones)
    role = st.session_state.user_role
    if role == "Campo-Tablet":
        views = [
            ("VISUALIZAR", "👁️ Visualizar"),
            ("ACTUALIZAR", "✏️ Actualizar / Foto"),
        ]
    else:
        views = [
            ("VISUALIZAR", "👁️ Visualizar"),
            ("ACTUALIZAR", "✏️ Actualizar / Foto"),
            ("AGREGAR", "➕ Agregar"),
            ("ESTRATEGIA", "📊 Estrategia y Backup"),
        ]

    nav_cols = st.columns(len(views))
    for i, (v_id, v_label) in enumerate(views):
        with nav_cols[i]:
            if st.button(
                v_label,
                key=f"nav_btn_{v_id}",
                type="primary" if st.session_state.current_view == v_id else "secondary",
                use_container_width=True,
            ):
                st.session_state.current_view = v_id
                st.rerun()


def render_flota_selector(df_all: pd.DataFrame) -> pd.DataFrame:
    """Permite a cada Planner elegir su Flota para sentir control propio sin duplicar tablas."""
    flotas = ["Todas"] + sorted(df_all["flota"].dropna().unique().tolist())
    if st.session_state.active_flota not in flotas:
        st.session_state.active_flota = "Todas"

    c_fl, c_badge = st.columns([2.2, 2.8])
    with c_fl:
        sel = st.selectbox(
            "🚜 Espacio de Flota (Control del Planner):",
            options=flotas,
            index=flotas.index(st.session_state.active_flota),
            key="global_flota_picker",
        )
        if sel != st.session_state.active_flota:
            st.session_state.active_flota = sel
            st.rerun()
    with c_badge:
        st.write("")
        if sel == "Todas":
            st.markdown("<span class='badge badge-primary'>Vista Global: Todas las Flotas</span>", unsafe_allow_html=True)
        else:
            st.markdown(get_flota_badge_html(sel), unsafe_allow_html=True)

    return df_all if sel == "Todas" else df_all[df_all["flota"] == sel]


# ==============================================================================
# 5. VENTANA 1: VISUALIZAR (POR FLOTA + TARJETAS + HOJA DE VIDA)
# ==============================================================================
def render_view_visualizar() -> None:
    conn = get_connection()
    df_all = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
    if df_all.empty:
        conn.close()
        st.info("Aún no hay componentes. Ve a **➕ Agregar** para importar tu Excel SAP o registrar uno.")
        return

    df_flota = render_flota_selector(df_all)

    # KPIs compactos de la flota activa
    total = len(df_flota)
    montados = int((df_flota["estado_actual"] == "Montado").sum())
    taller = int(df_flota["estado_actual"].isin(["Por Reparar", "Desmontado", "Salida por Reclamo"]).sum())
    gasto = float(df_flota["gasto_acumulado_total_usd"].sum())
    hrs = float(df_flota["horas_acumuladas_totales"].sum())
    c_hr = (gasto / hrs) if hrs > 0 else 0.0

    st.markdown(
        f"""
        <div class="responsive-kpi-grid">
            <div class="mac-glass-card"><div class="kpi-label">Componentes</div><div class="kpi-num">{total}</div><div class="kpi-meta">Flota: {st.session_state.active_flota}</div></div>
            <div class="mac-glass-card"><div class="kpi-label">Montados</div><div class="kpi-num" style="color:var(--success);">{montados}</div><div class="kpi-meta">En operación</div></div>
            <div class="mac-glass-card"><div class="kpi-label">Por Reparar / Taller</div><div class="kpi-num" style="color:var(--warning);">{taller}</div><div class="kpi-meta">Atención pendiente</div></div>
            <div class="mac-glass-card"><div class="kpi-label">Gasto · Costo/Hr</div><div class="kpi-num">${gasto:,.0f}</div><div class="kpi-meta" style="color:var(--danger); font-weight:700;">${c_hr:,.2f} USD/hr</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Filtros simples (sin sobrecargar)
    f1, f2, f3 = st.columns([1.2, 1.2, 2.0])
    with f1:
        sel_eq = st.selectbox("Equipo:", ["Todos"] + sorted(df_flota["equipo_principal"].dropna().unique().tolist()))
    with f2:
        sel_est = st.selectbox("Estado:", ["Todos"] + ESTADOS_COMPONENTE)
    with f3:
        q = st.text_input("Buscar:", placeholder="Código SAP, descripción o serie...")

    df_filt = df_flota.copy()
    if sel_eq != "Todos":
        df_filt = df_filt[df_filt["equipo_principal"] == sel_eq]
    if sel_est != "Todos":
        df_filt = df_filt[df_filt["estado_actual"] == sel_est]
    if q.strip():
        ql = q.strip().lower()
        df_filt = df_filt[
            df_filt["codigo_equipo"].astype(str).str.lower().str.contains(ql, na=False)
            | df_filt["descripcion"].astype(str).str.lower().str.contains(ql, na=False)
            | df_filt["serie"].astype(str).str.lower().str.contains(ql, na=False)
        ]

    if df_filt.empty:
        conn.close()
        st.warning("No se encontraron componentes con esos filtros.")
        return

    # Tarjetas con indicador claro de Flota + Estado + Vidas
    cards_html = ['<div class="responsive-card-grid">']
    for _, r in df_filt.head(18).iterrows():
        c_hr_item = (r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"]) if r["horas_acumuladas_totales"] > 0 else 0.0
        cards_html.append(
            f"""
            <div class="mac-glass-card">
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.3rem; margin-bottom:0.35rem;">
                    {get_flota_badge_html(r['flota'])}
                    {get_status_badge_html(r['estado_actual'])}
                </div>
                <div style="font-size:0.75rem; color:var(--text-muted); font-weight:700;">{r['codigo_equipo']} · Equipo {r['equipo_principal']}</div>
                <div style="font-weight:700; font-size:0.94rem; margin:0.15rem 0 0.4rem 0;">{r['descripcion']}</div>
                <div style="display:flex; justify-content:space-between; font-size:0.79rem; color:var(--text-secondary);">
                    <span>🔄 <b>{format_vidas_label(r['ciclo_actual'])}</b></span>
                    <span>⏱️ {float(r['horas_ciclo_actual']):,.0f} hrs</span>
                    <span>💵 ${c_hr_item:,.2f}/hr</span>
                </div>
            </div>
            """
        )
    cards_html.append("</div>")
    st.markdown("".join(cards_html), unsafe_allow_html=True)

    st.markdown("---")
    opciones = [f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']})" for _, r in df_filt.iterrows()]
    cod_list = df_filt["codigo_equipo"].tolist()
    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0
    sel_comp = st.selectbox("🔍 Ver Hoja de Vida del Componente:", opciones, index=idx_def)
    cod_sel = sel_comp.split(" | ")[0].strip()
    st.session_state.selected_component = cod_sel

    if st.session_state.vida_comp_ref != cod_sel:
        st.session_state.vida_comp_ref = cod_sel
        st.session_state.vida_idx = max(1, int(df_filt[df_filt["codigo_equipo"] == cod_sel].iloc[0]["ciclo_actual"] or 1))

    comp = df_filt[df_filt["codigo_equipo"] == cod_sel].iloc[0]
    total_vidas = max(0, int(comp["ciclo_actual"] or 1))
    nav_state = get_vidas_navigation_state(total_vidas, st.session_state.vida_idx)

    # Hoja de vida + botón directo a Actualizar
    c_det, c_go = st.columns([3.5, 1.2])
    with c_det:
        st.markdown(
            f"""
            <div class="mac-glass-card">
                <div style="display:flex; gap:0.5rem; flex-wrap:wrap; margin-bottom:0.35rem;">
                    {get_flota_badge_html(comp['flota'])}
                    {get_status_badge_html(comp['estado_actual'])}
                    <span class="badge badge-neutral">🔄 {nav_state['label']}</span>
                </div>
                <h3 style="margin:0.1rem 0;">{comp['descripcion']}</h3>
                <div style="font-size:0.84rem; color:var(--text-secondary);">
                    <b>SAP:</b> {comp['codigo_equipo']} · <b>Equipo:</b> {comp['equipo_principal']} · <b>Serie:</b> {comp['serie']} · <b>Fabricante:</b> {comp['fabricante']}
                </div>
                <div style="font-size:0.84rem; color:var(--text-primary); margin-top:0.4rem;">
                    ⏱️ <b>Hrs Ciclo:</b> {float(comp['horas_ciclo_actual']):,.1f} hrs &nbsp;|&nbsp;
                    ⏳ <b>Hrs Totales:</b> {float(comp['horas_acumuladas_totales']):,.1f} hrs &nbsp;|&nbsp;
                    💵 <b>Costo Acumulado:</b> ${float(comp['gasto_acumulado_total_usd']):,.2f} USD
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c_go:
        st.write("")
        if st.button("✏️ Actualizar este componente", type="primary", use_container_width=True):
            st.session_state.current_view = "ACTUALIZAR"
            st.rerun()
        if nav_state["show_nav"]:
            b1, b2 = st.columns(2)
            with b1:
                if st.button("◀", disabled=nav_state["disable_prev"], use_container_width=True, help="Vida anterior"):
                    st.session_state.vida_idx = max(1, nav_state["current"] - 1)
                    st.rerun()
            with b2:
                if st.button("▶", disabled=nav_state["disable_next"], use_container_width=True, help="Vida siguiente"):
                    st.session_state.vida_idx = min(total_vidas, nav_state["current"] + 1)
                    st.rerun()

    # 2 Últimas Fotos
    p1, p2 = st.columns(2)
    with p1:
        st.markdown('<div class="mac-glass-card"><b>📷 Foto #1 (Previa)</b>', unsafe_allow_html=True)
        if comp["foto_1_path"] and os.path.exists(comp["foto_1_path"]):
            st.image(comp["foto_1_path"], use_container_width=True)
            st.caption(f"{comp['foto_1_fecha'] or ''} — {comp['foto_1_comentario'] or ''}")
        else:
            st.caption("Sin fotografía en espacio #1.")
        st.markdown("</div>", unsafe_allow_html=True)
    with p2:
        st.markdown('<div class="mac-glass-card"><b>📷 Foto #2 (Más Reciente)</b>', unsafe_allow_html=True)
        if comp["foto_2_path"] and os.path.exists(comp["foto_2_path"]):
            st.image(comp["foto_2_path"], use_container_width=True)
            st.caption(f"{comp['foto_2_fecha'] or ''} — {comp['foto_2_comentario'] or ''}")
        else:
            st.caption("Sin fotografía en espacio #2.")
        st.markdown("</div>", unsafe_allow_html=True)

    df_eventos = pd.read_sql_query("SELECT * FROM historial_eventos WHERE codigo_equipo = ? ORDER BY id_evento DESC", conn, params=(cod_sel,))
    conn.close()
    if nav_state["show_nav"] and not df_eventos.empty:
        df_eventos = df_eventos[df_eventos["numero_ciclo"] == nav_state["current"]]

    if not df_eventos.empty:
        st.markdown("#### 📓 Historial de Costos y Bitácora")
        st.dataframe(
            df_eventos[["fecha_hora", "numero_ciclo", "tipo_operacion", "estado_asignado", "numero_ot", "numero_oc", "costo_total_evento_usd", "comentario_bitacora"]],
            use_container_width=True,
            hide_index=True,
        )


# ==============================================================================
# 6. VENTANA 2: ACTUALIZAR COMPONENTE (ESTADOS + FOTOS FIFO + HORAS)
# ==============================================================================
def render_view_actualizar() -> None:
    conn = get_connection()
    df_all = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)
    conn.close()
    if df_all.empty:
        st.info("No hay componentes registrados.")
        return

    df_flota = render_flota_selector(df_all)
    if df_flota.empty:
        st.warning("No hay componentes en esta flota.")
        return

    opciones = [f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']}) [{r['estado_actual']}]" for _, r in df_flota.iterrows()]
    cod_list = df_flota["codigo_equipo"].tolist()
    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0
    sel_comp = st.selectbox("Componente a actualizar:", opciones, index=idx_def)
    cod_sel = sel_comp.split(" | ")[0].strip()
    st.session_state.selected_component = cod_sel
    comp_row = df_flota[df_flota["codigo_equipo"] == cod_sel].iloc[0]

    st.markdown(
        f"""
        <div class="mac-glass-card">
            <div style="display:flex; gap:0.5rem; flex-wrap:wrap; margin-bottom:0.25rem;">
                {get_flota_badge_html(comp_row['flota'])}
                {get_status_badge_html(comp_row['estado_actual'])}
                <span class="badge badge-neutral">🔄 {format_vidas_label(comp_row['ciclo_actual'])}</span>
            </div>
            <b>{comp_row['codigo_equipo']}</b> — {comp_row['descripcion']} · Equipo: <b>{comp_row['equipo_principal']}</b> · Hrs Ciclo: <b>{float(comp_row['horas_ciclo_actual']):,.1f}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    role = st.session_state.user_role
    if role == "Campo-Tablet":
        tab_foto, = st.tabs(["📸 Foto y Comentario (FIFO)"])
        tab_estado = None
        tab_horas = None
    else:
        tab_foto, tab_estado, tab_horas = st.tabs(["📸 Foto y Comentario", "🔄 Estado, OT/OC y Costos", "⏱️ Cargar Horas (Todas las Flotas)"])

    # --- PESTAÑA DIRECTA: FOTO Y COMENTARIO (FIFO 2 FOTOS) ---
    with tab_foto:
        p1, p2 = st.columns(2)
        with p1:
            if comp_row["foto_1_path"] and os.path.exists(comp_row["foto_1_path"]):
                st.image(comp_row["foto_1_path"], caption=f"Foto #1: {comp_row['foto_1_comentario'] or ''}", use_container_width=True)
            else:
                st.info("Espacio Foto #1 libre")
        with p2:
            if comp_row["foto_2_path"] and os.path.exists(comp_row["foto_2_path"]):
                st.image(comp_row["foto_2_path"], caption=f"Foto #2 (Reciente): {comp_row['foto_2_comentario'] or ''}", use_container_width=True)
            else:
                st.info("Espacio Foto #2 libre")

        metodo = st.radio("Origen de imagen:", ["📁 Subir desde galería/archivo", "📸 Tomar foto con cámara"], horizontal=True)
        img_input = st.file_uploader("Seleccionar fotografía:", type=["jpg", "jpeg", "png", "webp"], key="upd_foto_up") if metodo.startswith("📁") else st.camera_input("Capturar fotografía:")
        com_foto = st.text_area("Comentario técnico:", placeholder="Observación de campo o taller...", height=85, key="upd_foto_com")

        c_b1, c_b2 = st.columns(2)
        with c_b1:
            if st.button("Guardar Foto y Comentario", type="primary", use_container_width=True):
                if img_input is None:
                    st.error("Selecciona o toma una foto primero.")
                else:
                    ok, msg, _ = save_component_photo_fifo(cod_sel, img_input, com_foto, role, True)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
        with c_b2:
            if st.button("Solo Guardar Comentario", type="secondary", use_container_width=True):
                if not com_foto.strip():
                    st.warning("Escribe un comentario.")
                else:
                    conn = get_connection()
                    cur = conn.cursor()
                    cur.execute(
                        "INSERT INTO historial_eventos (codigo_equipo, numero_ciclo, fecha_hora, usuario_rol, tipo_operacion, estado_asignado, equipo_principal, horas_componente_evento, comentario_bitacora) VALUES (?, ?, ?, ?, 'Nota / Comentario', ?, ?, ?, ?)",
                        (cod_sel, comp_row["ciclo_actual"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), role, comp_row["estado_actual"], comp_row["equipo_principal"], comp_row["horas_ciclo_actual"], com_foto.strip()),
                    )
                    conn.commit()
                    conn.close()
                    st.success("Comentario guardado.")

    # --- PESTAÑA ESTADO, OT/OC Y COSTOS ---
    if tab_estado is not None:
        with tab_estado:
            c1, c2 = st.columns(2)
            with c1:
                nuevo_estado = st.selectbox("Nuevo Estado:", ESTADOS_COMPONENTE, index=ESTADOS_COMPONENTE.index(comp_row["estado_actual"]) if comp_row["estado_actual"] in ESTADOS_COMPONENTE else 0)
            with c2:
                equipo_dest = st.text_input("Equipo Principal:", value=str(comp_row["equipo_principal"]))

            instalado_por, reiniciar_ciclo = "N/A", False
            motivo_cambio, modo_falla = "N/A", "N/A"
            numero_ot, numero_oc, proveedor, tipo_interv, detalle_kit = "", "", "", "Ninguna", ""
            costo_kit_txt, costo_rep_txt = "0", "0"

            if nuevo_estado == "Montado":
                m1, m2 = st.columns(2)
                with m1:
                    instalado_por = st.radio("Instalado por:", ["Southern", "Contratista"], horizontal=True)
                with m2:
                    reiniciar_ciclo = st.checkbox(f"Iniciar nueva vida ({int(comp_row['ciclo_actual']) + 1} vidas)", value=(comp_row["estado_actual"] != "Montado"))

            if nuevo_estado in ("Desmontado", "Por Reparar", "Salida por Reclamo", "De Baja"):
                d1, d2 = st.columns(2)
                with d1:
                    motivo_cambio = st.selectbox("Motivo del cambio:", ["PCR por Horas", "Falla por Condición", "Reclamo Garantía"])
                with d2:
                    modo_falla = st.selectbox("Modo de falla:", MODOS_FALLA_COMUNES, index=0 if motivo_cambio == "PCR por Horas" else 1)

            if nuevo_estado in ("Reparado", "Por Reparar", "Salida por Reclamo") or st.checkbox("Registrar OT / OC / Costos", value=False):
                r1, r2, r3 = st.columns(3)
                with r1:
                    numero_ot = st.text_input("N° OT SAP:")
                with r2:
                    numero_oc = st.text_input("N° Orden de Compra (OC):")
                with r3:
                    proveedor = st.text_input("Taller / Proveedor:")
                k1, k2, k3 = st.columns(3)
                with k1:
                    tipo_interv = st.selectbox("Intervención:", ["Kit + Reparación General", "Kit de Mantenimiento", "Reparación General", "Ninguna"])
                with k2:
                    costo_kit_txt = st.text_input("Costo Kit (USD):", value="0")
                with k3:
                    costo_rep_txt = st.text_input("Costo Reparación (USD):", value="0")
                detalle_kit = st.text_input("Detalle de Kit / Reparación:")

            ok_k, c_kit, err_k = validate_numeric_input(costo_kit_txt, "Costo Kit", allow_zero=True, allow_negative=False, allow_decimal=True)
            ok_r, c_rep, err_r = validate_numeric_input(costo_rep_txt, "Costo Reparación", allow_zero=True, allow_negative=False, allow_decimal=True)
            if not ok_k:
                st.error(err_k)
            if not ok_r:
                st.error(err_r)

            com_est = st.text_input("Comentario del movimiento:")
            confirm_baja = True
            if nuevo_estado == "De Baja":
                st.markdown("<div class='danger-zone'>⚠️ Dar de baja retirará el componente de la operación activa.</div>", unsafe_allow_html=True)
                confirm_baja = st.checkbox("Confirmo dar de baja este componente.")

            if st.button("Guardar Estado y Costos", type="primary", use_container_width=True, disabled=(not ok_k or not ok_r or not confirm_baja)):
                costo_tot = round(float(c_kit or 0.0) + float(c_rep or 0.0), 2)
                conn = get_connection()
                cur = conn.cursor()
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ciclo_db = int(comp_row["ciclo_actual"] or 1)
                if nuevo_estado == "Montado" and reiniciar_ciclo:
                    n_ciclo, n_hrs, n_gasto_c = ciclo_db + 1, 0.0, costo_tot
                else:
                    n_ciclo, n_hrs, n_gasto_c = ciclo_db, float(comp_row["horas_ciclo_actual"] or 0.0), float(comp_row["gasto_ciclo_actual_usd"] or 0.0) + costo_tot
                n_gasto_t = float(comp_row["gasto_acumulado_total_usd"] or 0.0) + costo_tot

                cur.execute(
                    "UPDATE componentes SET estado_actual = ?, equipo_principal = ?, ciclo_actual = ?, horas_ciclo_actual = ?, gasto_ciclo_actual_usd = ?, gasto_acumulado_total_usd = ? WHERE codigo_equipo = ?",
                    (nuevo_estado, equipo_dest.strip(), n_ciclo, round(n_hrs, 2), round(n_gasto_c, 2), round(n_gasto_t, 2), cod_sel),
                )
                cur.execute(
                    """
                    INSERT INTO historial_eventos (
                        codigo_equipo, numero_ciclo, fecha_hora, usuario_rol, tipo_operacion,
                        estado_asignado, equipo_principal, horas_componente_evento, horas_sumadas,
                        instalado_por, motivo_cambio, modo_falla, numero_ot, numero_oc,
                        proveedor_reparador, tipo_intervencion, detalle_kit, costo_kit_usd,
                        costo_reparacion_usd, costo_total_evento_usd, comentario_bitacora
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (cod_sel, n_ciclo, now_str, role, f"Cambio de Estado ({nuevo_estado})", nuevo_estado, equipo_dest.strip(), round(n_hrs, 2), instalado_por, motivo_cambio, modo_falla, numero_ot.strip(), numero_oc.strip(), proveedor.strip(), tipo_interv, detalle_kit.strip(), round(float(c_kit or 0.0), 2), round(float(c_rep or 0.0), 2), costo_tot, (com_est or f"Cambio a {nuevo_estado}").strip()),
                )
                conn.commit()
                conn.close()
                st.success(f"Componente `{cod_sel}` actualizado a **{nuevo_estado}**.")
                st.rerun()

    # --- PESTAÑA CARGA DE HORAS GLOBAL ---
    if tab_horas is not None:
        with tab_horas:
            st.caption("Una sola carga alimenta y actualiza las horas de **todos los componentes Montados de todas las flotas**.")
            st.download_button(
                "📥 Descargar Formato Excel para Carga de Horas",
                data=generate_hours_template_excel(),
                file_name="Formato_Carga_Horas_Montados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            archivo_hrs = st.file_uploader("Subir Excel/CSV con [Equipo_Principal o Codigo_Componente, Horas_a_Sumar o Horometro_Actual]:", type=["xlsx", "xls", "csv"], key="up_hrs_global")
            if archivo_hrs is not None:
                df_up = pd.read_csv(archivo_hrs) if archivo_hrs.name.lower().endswith(".csv") else pd.read_excel(archivo_hrs)
                st.dataframe(df_up.head(10), use_container_width=True, hide_index=True)
                cols_norm = {c.strip().lower(): c for c in df_up.columns}
                t_col = cols_norm.get("equipo_principal") or cols_norm.get("codigo_componente") or cols_norm.get("codigo_equipo")
                t_type = "equipo_principal" if "equipo_principal" in cols_norm else "codigo_equipo"
                v_col = cols_norm.get("horas_a_sumar") or cols_norm.get("horometro_actual")
                m_op = "add" if "horas_a_sumar" in cols_norm else "set_cycle"
                if t_col and v_col:
                    if st.button("Actualizar Horas a Todos los Montados", type="primary", use_container_width=True):
                        upd_list = []
                        for _, r in df_up.iterrows():
                            ok_n, num_h, _ = validate_numeric_input("" if pd.isna(r[v_col]) else str(r[v_col]), allow_zero=False, allow_negative=False, allow_decimal=True)
                            if pd.notna(r[t_col]) and ok_n:
                                upd_list.append({"target_type": t_type, "target_value": str(r[t_col]).strip(), "mode": m_op, "value": num_h})
                        n_upd, det = apply_hours_to_mounted_components(upd_list, role, f"Archivo ({archivo_hrs.name})")
                        st.success(f"Se actualizaron {n_upd} componente(s) montado(s).")
                        if det:
                            st.dataframe(pd.DataFrame(det), use_container_width=True, hide_index=True)
                else:
                    st.error("Faltan columnas requeridas ('Equipo_Principal'/'Codigo_Componente' y 'Horas_a_Sumar'/'Horometro_Actual').")

            st.markdown("---")
            eq_quick = st.selectbox("O suma horas rápidas a un Equipo Principal Montado:", sorted(df_all[df_all["estado_actual"] == "Montado"]["equipo_principal"].dropna().unique().tolist()) or ["SIN_EQUIPO"])
            hrs_q_txt = st.text_input("Horas operadas a sumar (+ hrs):", value="0", key="hrs_q_txt")
            ok_hq, h_q_val, err_hq = validate_numeric_input(hrs_q_txt, "Horas", allow_zero=False, allow_negative=False, allow_decimal=True)
            if st.button("Sumar Horas al Equipo", type="secondary", use_container_width=True):
                if not ok_hq:
                    st.error(err_hq)
                else:
                    n_upd, det = apply_hours_to_mounted_components([{"target_type": "equipo_principal", "target_value": eq_quick, "mode": "add", "value": h_q_val}], role, f"Rápida ({eq_quick})")
                    st.success(f"Se sumaron +{h_q_val} hrs a {n_upd} componente(s) de {eq_quick}.")


# ==============================================================================
# 7. VENTANA 3: AGREGAR COMPONENTE + SINCRONIZACIÓN SAP
# ==============================================================================
def render_view_agregar() -> None:
    tab_manual, tab_sap = st.tabs(["➕ Nuevo Componente", "🔄 Cargar Lista SAP (.xlsx)"])

    with tab_manual:
        st.markdown('<div class="mac-glass-card">', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            cod_m = st.text_input("Código Equipo SAP *:", placeholder="10045990")
            desc_m = st.text_input("Descripción del Componente *:", placeholder="MOTOR DE GIRO TRASERO")
            ubic_m = st.text_input("Ubicación Técnica SAP *:", value="SPCC-TOQ-MIN-495H-SHO009-GIRO")
            flota_prev, eq_prev = parse_ubicacion_tecnica(ubic_m)
            st.markdown(f"{get_flota_badge_html(flota_prev)} &nbsp; Equipo: **{eq_prev}**", unsafe_allow_html=True)
            crit_m = st.selectbox("Criticidad ABC:", ["A", "B", "C"])
        with c2:
            fab_m = st.text_input("Fabricante:", placeholder="KOMATSU / CAT")
            serie_m = st.text_input("N° de Serie:", placeholder="SN-884920")
            est_m = st.selectbox("Estado Inicial:", ESTADOS_COMPONENTE)
            hrs_txt = st.text_input("Horas Iniciales:", value="0")
            costo_txt = st.text_input("Costo Inicial (USD):", value="0")
            vidas_txt = st.text_input("Vidas Iniciales:", value="1")

        ok_h, h_val, err_h = validate_numeric_input(hrs_txt, "Horas Iniciales", allow_zero=True, allow_negative=False, allow_decimal=True)
        ok_c, c_val, err_c = validate_numeric_input(costo_txt, "Costo Inicial", allow_zero=True, allow_negative=False, allow_decimal=True)
        ok_v, v_val, err_v = validate_numeric_input(vidas_txt, "Vidas", allow_zero=False, allow_negative=False, allow_decimal=False)
        for ok_i, err_i in ((ok_h, err_h), (ok_c, err_c), (ok_v, err_v)):
            if not ok_i:
                st.error(err_i)

        if st.button("Guardar Componente", type="primary", use_container_width=True):
            if not cod_m.strip() or not desc_m.strip():
                st.error("Código Equipo SAP y Descripción son obligatorios.")
            elif ok_h and ok_c and ok_v:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                conn = get_connection()
                cur = conn.cursor()
                try:
                    cur.execute(
                        "INSERT INTO componentes (codigo_equipo, descripcion, flota, equipo_principal, ubicacion_tecnica, criticidad_abc, fabricante, serie, estado_actual, ciclo_actual, horas_acumuladas_totales, horas_ciclo_actual, gasto_acumulado_total_usd, gasto_ciclo_actual_usd, ultima_actualizacion_horas) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (cod_m.strip(), desc_m.strip().upper(), flota_prev, eq_prev, ubic_m.strip().upper(), crit_m, (fab_m.strip() or "S/I").upper(), (serie_m.strip() or "S/N").upper(), est_m, int(v_val), round(h_val, 2), round(h_val, 2), round(c_val, 2), round(c_val, 2), now_str),
                    )
                    conn.commit()
                    st.success(f"Componente `{cod_m.strip()}` agregado a Flota {flota_prev}.")
                except sqlite3.IntegrityError:
                    st.error(f"El código `{cod_m.strip()}` ya existe.")
                finally:
                    conn.close()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_sap:
        st.download_button(
            "📥 Descargar Plantilla Excel SAP ('Inventario SAP Nov.xlsx')",
            data=generate_sap_template_excel(),
            file_name="Inventario SAP Nov.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        sap_file = st.file_uploader("Subir Excel de SAP:", type=["xlsx", "xls"], key="sap_up_v3")
        if sap_file is not None:
            xl = openpyxl.load_workbook(sap_file, read_only=True)
            sheet_to_use = "Sheet1" if "Sheet1" in xl.sheetnames else xl.sheetnames[0]
            sap_file.seek(0)
            df_sap_up = pd.read_excel(sap_file, sheet_name=sheet_to_use)
            st.dataframe(df_sap_up.head(10), use_container_width=True, hide_index=True)
            if st.button("Sincronizar Lista SAP", type="primary", use_container_width=True):
                n_new, n_upd, msg = import_sap_dataframe(df_sap_up, st.session_state.user_role)
                st.success(f"{msg} (Nuevos: {n_new} | Actualizados: {n_upd})")


# ==============================================================================
# 8. VENTANA 4: ESTRATEGIA POR FLOTA Y RESPALDO
# ==============================================================================
def render_view_estrategia() -> None:
    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes", conn)
    conn.close()
    if df_comp.empty:
        st.info("No hay datos para graficar.")
        return

    df_comp["costo_por_hora_usd"] = df_comp.apply(
        lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2) if r["horas_acumuladas_totales"] > 0 else 0.0,
        axis=1,
    )

    c1, c2 = st.columns(2)
    with c1:
        df_fl = df_comp.groupby("flota", as_index=False).agg(gasto=("gasto_acumulado_total_usd", "sum"), hrs=("horas_acumuladas_totales", "sum"))
        df_fl["costo_hr"] = df_fl.apply(lambda r: round(r["gasto"] / r["hrs"], 2) if r["hrs"] > 0 else 0.0, axis=1)
        fig = px.bar(df_fl, x="flota", y="costo_hr", color="flota", text="costo_hr", title="Costo por Hora ($/hr) por Flota")
        fig.update_traces(texttemplate="$%{text:.2f}/hr", textposition="outside")
        fig.update_layout(height=320, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown("**🚨 Top Componentes por Costo/Hora ($/hr)**")
        st.dataframe(df_comp.sort_values("costo_por_hora_usd", ascending=False).head(8)[["flota", "codigo_equipo", "descripcion", "costo_por_hora_usd", "estado_actual"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### 🛡️ Respaldo de Base de Datos")
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    b1, b2 = st.columns(2)
    with b1:
        st.download_button(
            "📥 Descargar Base Completa (.xlsx)",
            data=export_full_database_to_excel(),
            file_name=f"PIM_Backup_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with b2:
        if os.path.exists(DB_PATH):
            with open(DB_PATH, "rb") as f_db:
                st.download_button(
                    "🗄️ Descargar Archivo SQLite (.db)",
                    data=f_db.read(),
                    file_name=f"inventario_mina_{ts}.db",
                    mime="application/octet-stream",
                    use_container_width=True,
                )


# ==============================================================================
# 9. PUNTO DE ENTRADA PRINCIPAL
# ==============================================================================
def main() -> None:
    init_ui_session_state()
    inject_ios_macos_css(st.session_state.theme_mode)
    init_db()
    render_compact_header_and_nav()

    view = st.session_state.current_view
    if st.session_state.user_role == "Campo-Tablet" and view not in ("VISUALIZAR", "ACTUALIZAR"):
        view = "VISUALIZAR"

    if view == "VISUALIZAR":
        render_view_visualizar()
    elif view == "ACTUALIZAR":
        render_view_actualizar()
    elif view == "AGREGAR":
        render_view_agregar()
    elif view == "ESTRATEGIA":
        render_view_estrategia()


if __name__ == "__main__":
    main()
