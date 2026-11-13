# -*- coding: utf-8 -*-
"""
Sistema Integral de Gestión de Componentes, Trazabilidad SAP PM y Costo por Ciclo
Arquitectura: Argon Design System UI (100% Offline Corporate Ready)
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
# 0. CONFIGURACIÓN INICIAL Y CONSTANTES CORPORATIVAS
# ==============================================================================
st.set_page_config(
    page_title="Argon Fleet PM | Control de Componentes",
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

# Asegurar carpeta local de fotos
os.makedirs(PHOTOS_DIR, exist_ok=True)


# ==============================================================================
# 1. ESTÉTICA ARGON UI (CSS 100% OFFLINE SIN CDNs BLOQUEABLES)
# ==============================================================================
def inject_argon_css() -> None:
    st.markdown(
        """
        <style>
        /* Tipografía nativa corporativa segura para firewall */
        html, body, [class*="css"] {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: #f8f9fe !important;
            color: #32325d;
        }
        .stApp {
            background-color: #f8f9fe !important;
        }
        /* Contenedor principal y cabeceras */
        .argon-header {
            background: linear-gradient(87deg, #5e72e4 0, #825ee4 100%);
            padding: 1.4rem 1.8rem;
            border-radius: 0.5rem;
            color: #ffffff;
            box-shadow: 0 0 2rem 0 rgba(136, 152, 170, .15);
            margin-bottom: 1.25rem;
        }
        .argon-header h1 {
            color: #ffffff !important;
            font-size: 1.55rem;
            font-weight: 700;
            margin: 0 0 0.25rem 0;
            letter-spacing: -0.02em;
        }
        .argon-header p {
            color: rgba(255, 255, 255, 0.88);
            font-size: 0.92rem;
            margin: 0;
        }
        /* Tarjetas blancas Argon */
        .argon-card {
            background-color: #ffffff;
            border-radius: 0.5rem;
            box-shadow: 0 0 2rem 0 rgba(136, 152, 170, .15);
            padding: 1.25rem 1.35rem;
            margin-bottom: 1rem;
            border: 1px solid rgba(0, 0, 0, 0.04);
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .argon-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 0.5rem 2.2rem 0 rgba(136, 152, 170, .22);
        }
        .argon-kpi-title {
            color: #8898aa;
            font-size: 0.74rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 0.3rem;
        }
        .argon-kpi-value {
            color: #32325d;
            font-size: 1.55rem;
            font-weight: 800;
            line-height: 1.2;
            margin-bottom: 0.25rem;
        }
        .argon-kpi-sub {
            font-size: 0.8rem;
            color: #525f7f;
        }
        /* Tarjetas de módulo del Lobby */
        .module-card {
            background-color: #ffffff;
            border-radius: 0.5rem;
            box-shadow: 0 0 2rem 0 rgba(136, 152, 170, .15);
            padding: 1.3rem 1.4rem;
            margin-bottom: 0.65rem;
            border-left: 5px solid #5e72e4;
            min-height: 145px;
        }
        .module-card h3 {
            color: #32325d;
            font-size: 1.05rem;
            font-weight: 700;
            margin: 0 0 0.45rem 0;
        }
        .module-card p {
            color: #525f7f;
            font-size: 0.86rem;
            line-height: 1.45;
            margin: 0 0 0.5rem 0;
        }
        /* Etiquetas de estado semánticas Argon */
        .badge {
            display: inline-block;
            padding: 0.28em 0.68em;
            font-size: 0.73rem;
            font-weight: 700;
            line-height: 1;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
            border-radius: 0.375rem;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .badge-success { background-color: rgba(45, 206, 137, 0.15); color: #2dce89; }
        .badge-primary { background-color: rgba(94, 114, 228, 0.15); color: #5e72e4; }
        .badge-warning { background-color: rgba(251, 99, 64, 0.16); color: #fb6340; }
        .badge-danger  { background-color: rgba(245, 54, 92, 0.15); color: #f5365c; }
        .badge-info    { background-color: rgba(17, 205, 239, 0.16); color: #11cdef; }
        .badge-dark    { background-color: rgba(50, 50, 93, 0.12); color: #32325d; }

        /* Botones primarios y secundarios estilo Argon */
        div.stButton > button[kind="primary"] {
            background-color: #5e72e4 !important;
            border-color: #5e72e4 !important;
            color: #ffffff !important;
            border-radius: 0.5rem !important;
            font-weight: 600 !important;
            box-shadow: 0 4px 6px rgba(50, 50, 93, .11), 0 1px 3px rgba(0, 0, 0, .08) !important;
        }
        div.stButton > button[kind="primary"]:hover {
            background-color: #485fe0 !important;
            border-color: #485fe0 !important;
        }
        div.stButton > button[kind="secondary"] {
            background-color: #ffffff !important;
            color: #5e72e4 !important;
            border: 1px solid #5e72e4 !important;
            border-radius: 0.5rem !important;
            font-weight: 600 !important;
        }
        /* Sidebar limpia */
        section[data-testid="stSidebar"] {
            background-color: #ffffff !important;
            border-right: 1px solid rgba(0,0,0,0.06);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_status_badge_html(estado: str) -> str:
    mapping = {
        "Montado": ("badge-success", "🟢 Montado"),
        "Reparado": ("badge-primary", "🔵 Reparado (Listo)"),
        "Nuevo": ("badge-info", "✨ Nuevo"),
        "Desmontado": ("badge-dark", "⚪ Desmontado"),
        "Por Reparar": ("badge-warning", "🟠 Por Reparar"),
        "Salida por Reclamo": ("badge-warning", "🛡️ Salida por Reclamo"),
        "De Baja": ("badge-danger", "🔴 De Baja"),
    }
    css_cls, label = mapping.get(estado, ("badge-dark", estado))
    return f'<span class="badge {css_cls}">{label}</span>'


# ==============================================================================
# 2. CAPA DE PERSISTENCIA SQLITE3 Y ARQUITECTURA DE DATOS SAP PM
# ==============================================================================
def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def parse_ubicacion_tecnica(ubic_tec: str) -> Tuple[str, str]:
    """
    Extrae del 4to bloque separado por guiones la `flota` (ej. 495H, 410A, 994F)
    y del 5to bloque el `equipo_principal` (ej. FLO005, SHO009).
    Ejemplo SAP PM: SPCC-TOQ-MIN-495H-SHO009-SIST -> flota='495H', equipo_principal='SHO009'
    """
    if not ubic_tec or pd.isna(ubic_tec):
        return ("SIN_FLOTA", "SIN_EQUIPO")
    bloques = [b.strip() for b in str(ubic_tec).strip().split("-") if b.strip()]
    flota = bloques[3] if len(bloques) >= 4 else "SIN_FLOTA"
    equipo_principal = bloques[4] if len(bloques) >= 5 else "SIN_EQUIPO"
    return flota, equipo_principal


def init_db() -> None:
    conn = get_connection()
    cur = conn.cursor()

    # 1. Tabla componentes
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

    # 2. Tabla historial_eventos
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

    # 3. Tabla log_carga_horas
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

    # Si la base está vacía, verificar si existe "Inventario SAP Nov.xlsx" o cargar semilla realista de Palas/Flota
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
    """
    Carga un conjunto de componentes representativos de Palas y Flota Minera (formato SAP PM)
    para que el sistema sea inmediatamente testeable al abrirse por primera vez.
    """
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    demo_components = [
        (
            "10045801",
            "MOTOR DE GIRO DELANTERO AC P&H",
            "495H",
            "SHO009",
            "SPCC-TOQ-MIN-495H-SHO009-GIRO",
            "A",
            "KOMATSU / P&H",
            "SN-495H-8821",
            "Montado",
            2,
            18450.0,
            4250.0,
            68500.0,
            24500.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045802",
            "TRANSMISION DE GIRO DERECHA",
            "495H",
            "SHO009",
            "SPCC-TOQ-MIN-495H-SHO009-GIRO",
            "A",
            "KOMATSU / P&H",
            "SN-495H-9104",
            "Montado",
            1,
            6120.0,
            6120.0,
            12400.0,
            12400.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045803",
            "MOTOR DE EMPUJE (CROWD) PRINCIPAL",
            "495H",
            "SHO008",
            "SPCC-TOQ-MIN-495H-SHO008-EMPU",
            "A",
            "SIEMENS",
            "SN-SIE-33019",
            "Montado",
            3,
            27800.0,
            3100.0,
            94200.0,
            31200.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045804",
            "CILINDRO HIDRAULICO DE LEVANTE IZQ",
            "994F",
            "FLO005",
            "SPCC-TOQ-MIN-994F-FLO005-HIDR",
            "A",
            "CATERPILLAR",
            "CAT-994-7741",
            "Montado",
            2,
            15900.0,
            5400.0,
            48000.0,
            18500.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045805",
            "BOMBA HIDRAULICA DE IMPLEMENTOS",
            "994F",
            "FLO005",
            "SPCC-TOQ-MIN-994F-FLO005-HIDR",
            "A",
            "REXROTH / CAT",
            "REX-55821",
            "Por Reparar",
            2,
            14200.0,
            7100.0,
            39500.0,
            14500.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045806",
            "REDUCTOR DE TRASLACION (PROPEL) DER",
            "410A",
            "SHO004",
            "SPCC-TOQ-MIN-410A-SHO004-PROP",
            "A",
            "BUCYRUS / CAT",
            "BUC-410-1120",
            "Reparado",
            2,
            19500.0,
            0.0,
            82000.0,
            29000.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045807",
            "COMPRESOR DE AIRE PRINCIPAL",
            "410A",
            "SHO004",
            "SPCC-TOQ-MIN-410A-SHO004-NEUM",
            "B",
            "SULLAIR",
            "SUL-88392",
            "Montado",
            1,
            4890.0,
            4890.0,
            6200.0,
            6200.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
        (
            "10045808",
            "FRENO DISCO DE IZAJE (HOIST) IZQ",
            "495H",
            "SHO009",
            "SPCC-TOQ-MIN-495H-SHO009-IZAJ",
            "B",
            "KOMATSU / P&H",
            "PH-BRK-601",
            "Desmontado",
            1,
            7850.0,
            7850.0,
            9800.0,
            9800.0,
            None,
            None,
            None,
            None,
            None,
            None,
            now_str,
        ),
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

    # Eventos semilla para trazabilidad financiera y hoja de vida
    demo_events = [
        (
            "10045801",
            1,
            "2025-11-15 08:30:00",
            "Planner",
            "Intervención / Reparación",
            "Reparado",
            "SHO009",
            14200.0,
            0.0,
            "N/A",
            "PCR por Horas",
            "N/A - Cambio Programado (PCR)",
            "OT-40089120",
            "4500192831",
            "KOMATSU MITSUI MAQUINARIAS PERU",
            "Kit + Reparación General",
            "Kit de rodamientos aislados SKF + Sellos Viton + Rebobinado rotor",
            6500.0,
            37500.0,
            44000.0,
            "Overhaul mayor ciclo 1 completado según estándar OEM.",
            None,
        ),
        (
            "10045801",
            2,
            "2026-02-10 14:15:00",
            "Planner",
            "Montaje en Equipo",
            "Montado",
            "SHO009",
            14200.0,
            0.0,
            "Southern",
            "PCR por Horas",
            "N/A - Cambio Programado (PCR)",
            "OT-40095410",
            "4500198004",
            "TALLER ELECTRICO SPCC",
            "Kit de Mantenimiento",
            "Kit de montaje, acoplamiento y sensores PT100",
            24500.0,
            0.0,
            24500.0,
            "Montado en Pala SHO009 durante parada programada. Inicia Ciclo 2.",
            None,
        ),
        (
            "10045802",
            1,
            "2026-04-02 10:00:00",
            "Planner",
            "Intervención / Reparación",
            "Montado",
            "SHO009",
            3500.0,
            0.0,
            "Southern",
            "N/A",
            "N/A - Cambio Programado (PCR)",
            "OT-40101220",
            "4500201145",
            "FERREYROS / TALLER CENTRAL",
            "Kit de Mantenimiento",
            "Kit de sellos de alta presión y respiradero",
            12400.0,
            0.0,
            12400.0,
            "Mantenimiento preventivo intermedio en PM de pala.",
            None,
        ),
        (
            "10045806",
            2,
            "2026-09-18 16:40:00",
            "Planner",
            "Intervención / Reparación",
            "Reparado",
            "SHO004",
            19500.0,
            0.0,
            "Contratista",
            "Falla por Condición",
            "Desgaste prematuro de rodamientos / bujes",
            "OT-40118900",
            "4500219087",
            "FLSMIDTH / MAESTRANZA SUR",
            "Kit + Reparación General",
            "Kit planetario etapa 1 y 2 + Rectificado de alojamiento",
            9000.0,
            20000.0,
            29000.0,
            "Componente reparado y preservado en bahía de listos.",
            None,
        ),
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


# ==============================================================================
# 3. REGLA ESTRICTA DE LAS 2 ÚLTIMAS FOTOS (FIFO CON BORRADO FÍSICO)
# ==============================================================================
def save_component_photo_fifo(
    codigo_equipo: str,
    uploaded_file,
    comentario_foto: str,
    usuario_rol: str,
    registrar_en_bitacora: bool = True,
) -> Tuple[bool, str, Optional[str]]:
    """
    Regla FIFO Estricta:
    - Cada componente guarda máximo 2 fotos (`foto_1_path` y `foto_2_path`).
    - Si ya existen 2 fotos, borra físicamente del disco `foto_1_path`,
      mueve `foto_2` a `foto_1` y guarda la nueva imagen comprimida en `foto_2`.
    """
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

        # Procesar y comprimir la imagen con Pillow para optimizar almacenamiento local
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
            # Slot 1 libre: se guarda directamente en foto_1
            cur.execute(
                """
                UPDATE componentes
                SET foto_1_path = ?, foto_1_fecha = ?, foto_1_comentario = ?
                WHERE codigo_equipo = ?
                """,
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Guardada en Espacio Fotográfico #1."
        elif not f2_path:
            # Slot 2 libre: se guarda en foto_2
            cur.execute(
                """
                UPDATE componentes
                SET foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ?
                WHERE codigo_equipo = ?
                """,
                (new_photo_path, now_display, comentario_limpio, codigo_equipo),
            )
            msg_rotacion = "Guardada en Espacio Fotográfico #2."
        else:
            # Ambos ocupados -> FIFO ESTRICTO: borrar físicamente foto_1_path del disco
            if f1_path and os.path.exists(f1_path):
                try:
                    os.remove(f1_path)
                except OSError:
                    pass

            # Mover foto_2 -> foto_1 y asentar nueva imagen en foto_2
            cur.execute(
                """
                UPDATE componentes
                SET foto_1_path = ?, foto_1_fecha = ?, foto_1_comentario = ?,
                    foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ?
                WHERE codigo_equipo = ?
                """,
                (
                    f2_path,
                    f2_fecha,
                    f2_com,
                    new_photo_path,
                    now_display,
                    comentario_limpio,
                    codigo_equipo,
                ),
            )
            msg_rotacion = (
                "Rotación FIFO aplicada: se eliminó físicamente la foto más antigua (#1), "
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


# ==============================================================================
# 4. FUNCIONES DE NEGOCIO: IMPORTACIÓN SAP, CARGA MASIVA HORAS Y BACKUP EXCEL
# ==============================================================================
def import_sap_dataframe(df: pd.DataFrame, usuario_rol: str) -> Tuple[int, int, str]:
    """
    Importa/Sincroniza el archivo Excel de SAP ('Inventario SAP Nov.xlsx' u homólogo)
    con columnas: ['Equipo', 'Denominación', 'Indicador ABC', 'Ubicac.técnica',
    'Denominación3', 'Fabricante', 'Fabr. Nº-serie', 'Válido de', 'Fin de validez'].
    Extrae del 4to bloque de 'Ubicac.técnica' la `flota` y del 5to bloque el `equipo_principal`.
    """
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
        # Limpiar formato numérico de Excel (ej. 10045801.0 -> '10045801')
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

        cur.execute(
            "SELECT codigo_equipo FROM componentes WHERE codigo_equipo = ?",
            (codigo_equipo,),
        )
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
                (
                    codigo_equipo,
                    descripcion,
                    flota,
                    equipo_principal,
                    ubic_tec,
                    criticidad_abc,
                    fabricante,
                    serie,
                    now_str,
                ),
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
                (
                    descripcion,
                    flota,
                    equipo_principal,
                    ubic_tec,
                    criticidad_abc,
                    fabricante,
                    serie,
                    codigo_equipo,
                ),
            )
            actualizados += 1

    cur.execute(
        """
        INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            now_str,
            usuario_rol,
            "Sincronización Maestro SAP PM",
            nuevos + actualizados,
            f"Nuevos dados de alta: {nuevos} | Actualizados en datos maestros: {actualizados}",
        ),
    )

    conn.commit()
    conn.close()
    return nuevos, actualizados, "Sincronización SAP completada con éxito."


def apply_hours_to_mounted_components(
    updates: List[Dict],
    usuario_rol: str,
    metodo_carga: str,
) -> Tuple[int, List[Dict]]:
    """
    Actualiza las horas exclusivamente a los componentes en estado 'Montado'.
    `updates` puede contener llaves:
      - {'target_type': 'equipo_principal' | 'codigo_equipo', 'target_value': str,
         'mode': 'add' | 'set_cycle', 'value': float}
    Suma las horas tanto en `horas_ciclo_actual` como en `horas_acumuladas_totales`,
    registra `ultima_actualizacion_horas` y asienta el log en `log_carga_horas` e `historial_eventos`.
    """
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
                SET horas_ciclo_actual = ?,
                    horas_acumuladas_totales = ?,
                    ultima_actualizacion_horas = ?
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
        resumen_str = (
            f"Se actualizaron {len(afectados_detalle)} componentes montados. "
            f"Ejemplos: {', '.join(x['Código Componente'] for x in afectados_detalle[:5])}"
        )
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
    """
    Genera en memoria un archivo Excel (.xlsx) con múltiples pestañas:
    1. Inventario_Actual
    2. Historial_Eventos_Costos
    3. Bitacora_Carga_Horas
    """
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

        # Auto-ajuste de ancho de columnas en las 3 hojas
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=12)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    output.seek(0)
    return output.getvalue()


def generate_sap_template_excel() -> bytes:
    """
    Crea una plantilla Excel exacta con las columnas de 'Inventario SAP Nov.xlsx' (hoja Sheet1)
    para facilitar pruebas o cargas masivas offline.
    """
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


# ==============================================================================
# 5. COMPONENTES VISUALES COMUNES: CABECERA, KPIs EJECUTIVOS Y NAVEGACIÓN
# ==============================================================================
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


def render_top_kpis() -> None:
    kpis = fetch_executive_kpis()
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">📦 Total Componentes</div>
                <div class="argon-kpi-value">{kpis['total']:,}</div>
                <div class="argon-kpi-sub"><span style="color:#5e72e4;font-weight:700;">SAP PM</span> · Flota Minera</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">🟢 En Operación (Montados)</div>
                <div class="argon-kpi-value" style="color:#2dce89;">{kpis['montados']:,}</div>
                <div class="argon-kpi-sub">Acumulando horas activas</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">🔵 Listos (Reparado / Nuevo)</div>
                <div class="argon-kpi-value" style="color:#5e72e4;">{kpis['reparados_listos']:,}</div>
                <div class="argon-kpi-sub">Disponibles para cambio</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">🟠 En Taller / Por Reparar</div>
                <div class="argon-kpi-value" style="color:#fb6340;">{kpis['por_reparar']:,}</div>
                <div class="argon-kpi-sub">Pendientes OT / OC / Garantía</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c5:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">💰 Gasto Acumulado / Costo Hr</div>
                <div class="argon-kpi-value">${kpis['gasto_acum_usd']:,.0f}</div>
                <div class="argon-kpi-sub"><span style="color:#f5365c;font-weight:700;">${kpis['costo_hora_global']:,.2f} USD/hr</span> global</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_sidebar_and_header() -> None:
    if "current_module" not in st.session_state:
        st.session_state.current_module = "LOBBY"
    if "user_role" not in st.session_state:
        st.session_state.user_role = "Planner"
    if "selected_component" not in st.session_state:
        st.session_state.selected_component = None

    with st.sidebar:
        st.markdown(
            """
            <div style="padding: 0.6rem 0.2rem 1rem 0.2rem; border-bottom: 1px solid #e9ecef; margin-bottom: 1rem;">
                <h2 style="margin:0; color:#5e72e4; font-size:1.3rem; font-weight:800;">⚙️ ARGON PM SUITE</h2>
                <span style="font-size:0.78rem; color:#8898aa; font-weight:600;">PLANEAMIENTO MANTENIMIENTO MINA</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Botón Permanente para volver al Inicio (Lobby)
        if st.button("🏠 Volver al Inicio (Lobby)", use_container_width=True, type="primary"):
            st.session_state.current_module = "LOBBY"
            st.rerun()

        st.markdown("### 👤 Perfil de Usuario Activo")
        selected_role = st.selectbox(
            "Seleccionar Rol Operativo:",
            options=ROLES_SISTEMA,
            index=ROLES_SISTEMA.index(st.session_state.user_role),
            help="Filtra automáticamente los módulos habilitados según los permisos del rol.",
        )
        if selected_role != st.session_state.user_role:
            st.session_state.user_role = selected_role
            # Si cambia a Campo-Tablet y estaba en un módulo restringido, redirigir al Lobby
            if selected_role == "Campo-Tablet" and st.session_state.current_module not in ("LOBBY", "MOD_1", "MOD_4"):
                st.session_state.current_module = "LOBBY"
            st.rerun()

        role_badge_color = {
            "Administrador": "#f5365c",
            "Planner": "#5e72e4",
            "Campo-Tablet": "#2dce89",
        }[st.session_state.user_role]

        st.markdown(
            f"""
            <div style="background:#f8f9fe; padding:0.65rem 0.85rem; border-radius:0.5rem; border-left:4px solid {role_badge_color}; margin-bottom:1rem;">
                <div style="font-size:0.72rem; color:#8898aa; text-transform:uppercase; font-weight:700;">Rol Activo</div>
                <div style="font-size:0.95rem; font-weight:700; color:#32325d;">{st.session_state.user_role}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### 🧭 Acceso Rápido a Módulos")
        if st.button("🔍 1. Buscador 360° y Hoja de Vida", use_container_width=True):
            st.session_state.current_module = "MOD_1"
            st.rerun()

        if st.session_state.user_role in ("Planner", "Administrador"):
            if st.button("⏱️ 2. Carga Masiva de Horas", use_container_width=True):
                st.session_state.current_module = "MOD_2"
                st.rerun()
            if st.button("⚡ 3. Gestión Estados, Kits y OC", use_container_width=True):
                st.session_state.current_module = "MOD_3"
                st.rerun()

        if st.button("📱 4. Modo Tablet Campo (Fotos)", use_container_width=True):
            st.session_state.current_module = "MOD_4"
            st.rerun()

        if st.session_state.user_role in ("Planner", "Administrador"):
            if st.button("➕ 5. Alta SAP y Backup Excel", use_container_width=True):
                st.session_state.current_module = "MOD_5"
                st.rerun()
            if st.button("📊 6. Estrategia y Costo por Ciclo", use_container_width=True):
                st.session_state.current_module = "MOD_6"
                st.rerun()

        st.markdown("---")
        st.caption(f"📁 DB Local: `{os.path.basename(DB_PATH)}`")
        st.caption(f"🖼️ Fotos FIFO: `{os.path.basename(PHOTOS_DIR)}/`")

    # Cabecera Principal Superior con Botón Rápido al Inicio si estamos dentro de un módulo
    col_hdr, col_btn = st.columns([5, 1.2])
    with col_hdr:
        st.markdown(
            f"""
            <div class="argon-header">
                <h1>⚙️ Sistema de Control de Componentes y Trazabilidad SAP PM</h1>
                <p>Gestión de Ciclos · Kits y Reparaciones · Órdenes de Compra (OC/OT) · Horómetros de Componentes Montados · Rol activo: <b>{st.session_state.user_role}</b></p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_btn:
        st.write("")
        if st.session_state.current_module != "LOBBY":
            if st.button("⬅️ Ir al Inicio", use_container_width=True, type="primary", key="top_home_btn"):
                st.session_state.current_module = "LOBBY"
                st.rerun()


# ==============================================================================
# 6. PANTALLA PRINCIPAL (LOBBY) CON TARJETAS DE SELECCIÓN POR ROL
# ==============================================================================
def render_lobby() -> None:
    render_top_kpis()

    role = st.session_state.user_role
    st.markdown(
        f"""
        <div style="margin: 0.5rem 0 1rem 0;">
            <h3 style="margin:0; color:#32325d;">🎛️ Panel de Navegación Principal (Lobby)</h3>
            <span style="color:#525f7f; font-size:0.9rem;">Selecciona un módulo de trabajo habilitado para el perfil <b>{role}</b>:</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    modules_config = [
        {
            "id": "MOD_1",
            "title": "🔍 Módulo 1: Buscador 360° y Hoja de Vida",
            "desc": "Consulta integral por flota, equipo principal o código SAP. KPIs de horas del ciclo, costo por hora ($/hr), galería comparativa de las 2 últimas fotos, desglose financiero por ciclo y bitácora completa.",
            "roles": ["Administrador", "Planner", "Campo-Tablet"],
            "color": "#5e72e4",
            "badge": "Todos los roles",
        },
        {
            "id": "MOD_2",
            "title": "⏱️ Módulo 2: Botón de Carga Masiva de Horas (Componentes Montados)",
            "desc": "Actualización masiva por archivo Excel/CSV o asignación rápida en pantalla por Flota/Equipo (ej. SHO009 +150 hrs) sumando automáticamente horas al ciclo actual y al acumulado total.",
            "roles": ["Administrador", "Planner"],
            "color": "#2dce89",
            "badge": "Planner y Admin",
        },
        {
            "id": "MOD_3",
            "title": "⚡ Módulo 3: Gestión de Estados, Kits, OC y Trazabilidad",
            "desc": "Formulario inteligente adaptable: Montaje (Southern/Contratista + incremento automático de ciclo), Desmontaje (PCR por horas vs. Falla por condición) y Reparación (OT, OC, Proveedor, Kit y Costos USD).",
            "roles": ["Administrador", "Planner"],
            "color": "#fb6340",
            "badge": "Planner y Admin",
        },
        {
            "id": "MOD_4",
            "title": "📱 Módulo 4: Modo Tablet Campo: Fotos y Comentarios",
            "desc": "Interfaz ágil para inspección en pampa o taller: captura/carga fotográfica con compresión automática y rotación estricta FIFO de las 2 últimas fotos con eliminación física de la anterior.",
            "roles": ["Administrador", "Planner", "Campo-Tablet"],
            "color": "#11cdef",
            "badge": "Campo, Planner y Admin",
        },
        {
            "id": "MOD_5",
            "title": "➕ Módulo 5: Alta de Componentes, Sincronización SAP y Backup",
            "desc": "Importación directa de 'Inventario SAP Nov.xlsx' con extracción automática de Flota (4to bloque) y Equipo Principal (5to bloque), alta manual y descarga de Base de Datos Completa a Excel (.xlsx).",
            "roles": ["Administrador", "Planner"],
            "color": "#5e72e4",
            "badge": "Planner y Admin",
        },
        {
            "id": "MOD_6",
            "title": "📊 Módulo 6: Estrategia de Componentes y Costo por Ciclo",
            "desc": "Analítica ejecutiva Plotly: Costo por hora ($/hr) por flota y fabricante, curva de horas vs. gasto por ciclo, Pareto de modos de falla y priorización de reemplazo o baja técnica.",
            "roles": ["Administrador", "Planner"],
            "color": "#f5365c",
            "badge": "Planner y Admin",
        },
    ]

    allowed_modules = [m for m in modules_config if role in m["roles"]]

    # Renderizar en grilla de 2 columnas
    for idx in range(0, len(allowed_modules), 2):
        cols = st.columns(2)
        for col_idx in range(2):
            if idx + col_idx < len(allowed_modules):
                mod = allowed_modules[idx + col_idx]
                with cols[col_idx]:
                    st.markdown(
                        f"""
                        <div class="module-card" style="border-left-color: {mod['color']};">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.4rem;">
                                <h3 style="margin:0;">{mod['title']}</h3>
                                <span class="badge badge-primary">{mod['badge']}</span>
                            </div>
                            <p>{mod['desc']}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button(f"Ingresar a {mod['title'].split(':')[0]} ➡️", key=f"btn_{mod['id']}", use_container_width=True, type="primary"):
                        st.session_state.current_module = mod["id"]
                        st.rerun()


# ==============================================================================
# 7. MÓDULO 1: BUSCADOR 360° Y HOJA DE VIDA DEL COMPONENTE
# ==============================================================================
def render_module_1_buscador_360() -> None:
    st.markdown("### 🔍 Módulo 1: Buscador 360° y Hoja de Vida del Componente")

    conn = get_connection()
    df_comp = pd.read_sql_query(
        "SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn
    )

    if df_comp.empty:
        conn.close()
        st.warning("No hay componentes registrados en la base de datos.")
        return

    # Barra de filtros 360°
    with st.container():
        f1, f2, f3, f4 = st.columns([1.3, 1.3, 1.3, 2.1])
        with f1:
            flotas = ["Todas"] + sorted(df_comp["flota"].dropna().unique().tolist())
            sel_flota = st.selectbox("Filtrar por Flota:", flotas)
        with f2:
            df_sub = df_comp if sel_flota == "Todas" else df_comp[df_comp["flota"] == sel_flota]
            equipos = ["Todos"] + sorted(df_sub["equipo_principal"].dropna().unique().tolist())
            sel_equipo = st.selectbox("Equipo Principal:", equipos)
        with f3:
            estados = ["Todos"] + ESTADOS_COMPONENTE
            sel_estado = st.selectbox("Estado Actual:", estados)
        with f4:
            texto_busqueda = st.text_input(
                "Búsqueda rápida (Código SAP, Serie, Descripción, Fabricante):",
                placeholder="Ej. 10045801, MOTOR DE GIRO, SHO009, SIEMENS...",
            )

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
        )
        df_filtered = df_filtered[mask]

    st.caption(f"Mostrando **{len(df_filtered)}** componentes encontrados en el inventario.")

    # Tabla resumen interactiva
    df_vista = df_filtered[
        [
            "codigo_equipo",
            "descripcion",
            "flota",
            "equipo_principal",
            "estado_actual",
            "ciclo_actual",
            "horas_ciclo_actual",
            "horas_acumuladas_totales",
            "gasto_ciclo_actual_usd",
            "gasto_acumulado_total_usd",
            "ultima_actualizacion_horas",
        ]
    ].copy()
    df_vista["costo_por_hora_usd"] = df_vista.apply(
        lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2)
        if r["horas_acumuladas_totales"] > 0
        else 0.0,
        axis=1,
    )

    st.dataframe(
        df_vista,
        use_container_width=True,
        hide_index=True,
        column_config={
            "codigo_equipo": "Código SAP",
            "descripcion": "Descripción del Componente",
            "flota": "Flota",
            "equipo_principal": "Equipo Principal",
            "estado_actual": "Estado",
            "ciclo_actual": "Ciclo #",
            "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo Actual", format="%.1f hrs"),
            "horas_acumuladas_totales": st.column_config.NumberColumn("Hrs Acum. Totales", format="%.1f hrs"),
            "gasto_ciclo_actual_usd": st.column_config.NumberColumn("Costo Ciclo ($)", format="$%.2f"),
            "gasto_acumulado_total_usd": st.column_config.NumberColumn("Costo Acum. ($)", format="$%.2f"),
            "costo_por_hora_usd": st.column_config.NumberColumn("Costo/Hr ($/hr)", format="$%.2f/hr"),
            "ultima_actualizacion_horas": "Últ. Act. Horas",
        },
    )

    if df_filtered.empty:
        conn.close()
        return

    st.markdown("---")
    st.markdown("#### 📋 Hoja de Vida Detallada del Componente Seleccionado")

    opciones_comp = [
        f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']}) [{r['estado_actual']}]"
        for _, r in df_filtered.iterrows()
    ]
    codigos_lista = df_filtered["codigo_equipo"].tolist()

    default_idx = 0
    if st.session_state.selected_component in codigos_lista:
        default_idx = codigos_lista.index(st.session_state.selected_component)

    seleccion_str = st.selectbox(
        "Seleccionar Componente para abrir su Hoja de Vida 360°:",
        options=opciones_comp,
        index=default_idx,
    )
    cod_sel = seleccion_str.split(" | ")[0].strip()
    st.session_state.selected_component = cod_sel

    comp = df_filtered[df_filtered["codigo_equipo"] == cod_sel].iloc[0]

    # Ficha Maestra del Componente
    badge_html = get_status_badge_html(comp["estado_actual"])
    st.markdown(
        f"""
        <div class="argon-card">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;">
                <div>
                    <span style="font-size:0.8rem; color:#8898aa; font-weight:700;">CÓDIGO EQUIPO SAP: {comp['codigo_equipo']} · SERIE: {comp['serie']}</span>
                    <h2 style="margin:0.15rem 0 0.3rem 0; color:#32325d; font-size:1.35rem;">{comp['descripcion']}</h2>
                    <span style="font-size:0.88rem; color:#525f7f;">
                        <b>Flota:</b> {comp['flota']} &nbsp;|&nbsp;
                        <b>Equipo Principal:</b> {comp['equipo_principal']} &nbsp;|&nbsp;
                        <b>Ubicación Técnica:</b> <code>{comp['ubicacion_tecnica']}</code> &nbsp;|&nbsp;
                        <b>Fabricante:</b> {comp['fabricante']} &nbsp;|&nbsp;
                        <b>Criticidad ABC:</b> {comp['criticidad_abc']}
                    </span>
                </div>
                <div style="text-align:right;">
                    {badge_html}
                    <div style="font-size:0.8rem; color:#8898aa; margin-top:0.35rem;">
                        Ciclo Actual: <b>#{int(comp['ciclo_actual'])}</b><br>
                        Últ. Carga Horas: {comp['ultima_actualizacion_horas'] or 'Sin registro'}
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 5 KPIs obligatorios de la Hoja de Vida
    hrs_ciclo = float(comp["horas_ciclo_actual"] or 0.0)
    hrs_total = float(comp["horas_acumuladas_totales"] or 0.0)
    costo_ciclo = float(comp["gasto_ciclo_actual_usd"] or 0.0)
    costo_total = float(comp["gasto_acumulado_total_usd"] or 0.0)
    costo_por_hr = (costo_total / hrs_total) if hrs_total > 0 else 0.0
    costo_por_hr_ciclo = (costo_ciclo / hrs_ciclo) if hrs_ciclo > 0 else 0.0

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">⏱️ Horas del Ciclo (#{int(comp['ciclo_actual'])})</div>
                <div class="argon-kpi-value" style="color:#5e72e4;">{hrs_ciclo:,.1f} hrs</div>
                <div class="argon-kpi-sub">Operadas en el montaje actual</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">⏳ Horas Totales Acumuladas</div>
                <div class="argon-kpi-value">{hrs_total:,.1f} hrs</div>
                <div class="argon-kpi-sub">Vida histórica total del activo</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">💵 Costo del Ciclo Actual</div>
                <div class="argon-kpi-value" style="color:#fb6340;">${costo_ciclo:,.2f}</div>
                <div class="argon-kpi-sub">Kits + Reparaciones ciclo #{int(comp['ciclo_actual'])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">🏦 Costo Acumulado Total</div>
                <div class="argon-kpi-value" style="color:#2dce89;">${costo_total:,.2f}</div>
                <div class="argon-kpi-sub">Inversión acumulada todos los ciclos</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k5:
        st.markdown(
            f"""
            <div class="argon-card">
                <div class="argon-kpi-title">📈 Costo por Hora ($/hr)</div>
                <div class="argon-kpi-value" style="color:#f5365c;">${costo_por_hr:,.2f}/hr</div>
                <div class="argon-kpi-sub">Ciclo actual: ${costo_por_hr_ciclo:,.2f}/hr</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Galería lado a lado de las 2 últimas fotos (FIFO)
    st.markdown("#### 📸 Registro Visual del Componente (2 Últimas Fotos Activas - FIFO)")
    p_col1, p_col2 = st.columns(2)

    with p_col1:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("**🖼️ Fotografía #1 (Anterior / Previa en rotación)**")
        f1_path = comp["foto_1_path"]
        if f1_path and os.path.exists(f1_path):
            st.image(f1_path, use_container_width=True)
            st.caption(f"📅 **Fecha:** {comp['foto_1_fecha'] or 'S/F'}")
            st.info(f"💬 **Comentario:** {comp['foto_1_comentario'] or 'Sin comentario'}")
        else:
            st.markdown(
                """
                <div style="background:#f8f9fe; border:1px dashed #cbd5e0; border-radius:0.5rem; padding:2.2rem; text-align:center; color:#8898aa;">
                    📷 Sin fotografía registrada en el Espacio #1
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with p_col2:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("**🖼️ Fotografía #2 (Más Reciente)**")
        f2_path = comp["foto_2_path"]
        if f2_path and os.path.exists(f2_path):
            st.image(f2_path, use_container_width=True)
            st.caption(f"📅 **Fecha:** {comp['foto_2_fecha'] or 'S/F'}")
            st.info(f"💬 **Comentario:** {comp['foto_2_comentario'] or 'Sin comentario'}")
        else:
            st.markdown(
                """
                <div style="background:#f8f9fe; border:1px dashed #cbd5e0; border-radius:0.5rem; padding:2.2rem; text-align:center; color:#8898aa;">
                    📷 Sin fotografía registrada en el Espacio #2 (se poblará al subir una 2da foto)
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # Historial Financiero por Ciclo y Bitácora Completa
    df_eventos = pd.read_sql_query(
        """
        SELECT * FROM historial_eventos
        WHERE codigo_equipo = ?
        ORDER BY id_evento DESC
        """,
        conn,
        params=(cod_sel,),
    )
    conn.close()

    st.markdown("#### 💳 Desglose Financiero por Ciclo (Kits + Reparación General + OC + OT + Quién Reparó)")
    if not df_eventos.empty:
        df_fin = df_eventos[
            (df_eventos["costo_total_evento_usd"] > 0)
            | (df_eventos["numero_ot"].fillna("") != "")
            | (df_eventos["numero_oc"].fillna("") != "")
            | (df_eventos["tipo_intervencion"].fillna("Ninguna") != "Ninguna")
        ].copy()

        if not df_fin.empty:
            st.dataframe(
                df_fin[
                    [
                        "numero_ciclo",
                        "fecha_hora",
                        "tipo_intervencion",
                        "detalle_kit",
                        "costo_kit_usd",
                        "costo_reparacion_usd",
                        "costo_total_evento_usd",
                        "numero_ot",
                        "numero_oc",
                        "proveedor_reparador",
                        "instalado_por",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "numero_ciclo": "Ciclo #",
                    "fecha_hora": "Fecha / Hora",
                    "tipo_intervencion": "Tipo Intervención",
                    "detalle_kit": "Detalle de Kit / Alcance",
                    "costo_kit_usd": st.column_config.NumberColumn("Costo Kit ($)", format="$%.2f"),
                    "costo_reparacion_usd": st.column_config.NumberColumn("Reparación Gral ($)", format="$%.2f"),
                    "costo_total_evento_usd": st.column_config.NumberColumn("Costo Total ($)", format="$%.2f"),
                    "numero_ot": "N° OT SAP",
                    "numero_oc": "N° Orden Compra (OC)",
                    "proveedor_reparador": "Taller / Quién Reparó",
                    "instalado_por": "Instalado Por",
                },
            )
        else:
            st.info("Este componente aún no registra costos de Kits ni Órdenes de Compra/Reparación.")

        st.markdown("#### 📓 Bitácora Completa de Comentarios, Movimientos y Trazabilidad")
        st.dataframe(
            df_eventos[
                [
                    "id_evento",
                    "fecha_hora",
                    "numero_ciclo",
                    "usuario_rol",
                    "tipo_operacion",
                    "estado_asignado",
                    "equipo_principal",
                    "horas_componente_evento",
                    "horas_sumadas",
                    "motivo_cambio",
                    "modo_falla",
                    "comentario_bitacora",
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "id_evento": "ID",
                "fecha_hora": "Fecha y Hora",
                "numero_ciclo": "Ciclo #",
                "usuario_rol": "Rol",
                "tipo_operacion": "Operación",
                "estado_asignado": "Estado",
                "equipo_principal": "Equipo",
                "horas_componente_evento": st.column_config.NumberColumn("Horómetro Comp.", format="%.1f hrs"),
                "horas_sumadas": st.column_config.NumberColumn("Hrs Sumadas", format="+%.1f hrs"),
                "motivo_cambio": "Motivo Cambio",
                "modo_falla": "Modo de Falla",
                "comentario_bitacora": "Comentario / Observación en Bitácora",
            },
        )
    else:
        st.info("No existen eventos registrados en la bitácora para este componente.")


# ==============================================================================
# 8. MÓDULO 2: BOTÓN DE CARGA MASIVA DE HORAS (COMPONENTES MONTADOS)
# ==============================================================================
def render_module_2_carga_horas() -> None:
    st.markdown("### ⏱️ Módulo 2: Botón de Carga Masiva de Horas (Componentes Montados)")
    st.markdown(
        """
        <div class="argon-card" style="border-left: 5px solid #2dce89;">
            <b>Regla de Integridad de Horómetros:</b> Toda actualización suma horas exclusivamente a los componentes cuyo
            <code>estado_actual == 'Montado'</code>, impactando simultáneamente en <b>Horas del Ciclo Actual</b> y en
            <b>Horas Acumuladas Totales</b>, dejando trazabilidad completa de fecha y usuario.
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab_a, tab_b, tab_log = st.tabs(
        [
            "📂 Opción A: Carga Masiva por Archivo Excel / CSV",
            "⚡ Opción B: Actualización Rápida por Flota / Equipo en Pantalla",
            "📜 Historial de Cargas de Horas",
        ]
    )

    # --------------------------------------------------------------------------
    # OPCIÓN A: Carga por archivo Excel o CSV
    # --------------------------------------------------------------------------
    with tab_a:
        st.markdown("#### 📂 Opción A: Subir Archivo Excel (`.xlsx`) o CSV (`.csv`)")
        st.write(
            "El archivo debe incluir una columna identificadora (**`Equipo_Principal`** o **`Codigo_Componente`**) "
            "y una columna de horas (**`Horas_a_Sumar`** o **`Horometro_Actual`**)."
        )

        # Botón de descarga de plantilla de ejemplo para Opción A
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
            "📥 Descargar Plantilla Excel de Ejemplo (Carga de Horas)",
            data=buf_tpl.getvalue(),
            file_name="Plantilla_Carga_Horas_Montados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        archivo_horas = st.file_uploader(
            "Seleccionar archivo Excel (.xlsx) o CSV (.csv):",
            type=["xlsx", "xls", "csv"],
            key="uploader_horas_montados",
        )

        if archivo_horas is not None:
            try:
                if archivo_horas.name.lower().endswith(".csv"):
                    df_up = pd.read_csv(archivo_horas)
                else:
                    df_up = pd.read_excel(archivo_horas)

                st.markdown("**Vista previa del archivo cargado:**")
                st.dataframe(df_up.head(15), use_container_width=True, hide_index=True)

                cols_norm = {c.strip().lower(): c for c in df_up.columns}

                target_col = None
                target_type = None
                if "equipo_principal" in cols_norm:
                    target_col = cols_norm["equipo_principal"]
                    target_type = "equipo_principal"
                elif "codigo_componente" in cols_norm:
                    target_col = cols_norm["codigo_componente"]
                    target_type = "codigo_equipo"
                elif "codigo_equipo" in cols_norm:
                    target_col = cols_norm["codigo_equipo"]
                    target_type = "codigo_equipo"

                value_col = None
                mode_op = None
                if "horas_a_sumar" in cols_norm:
                    value_col = cols_norm["horas_a_sumar"]
                    mode_op = "add"
                elif "horometro_actual" in cols_norm:
                    value_col = cols_norm["horometro_actual"]
                    mode_op = "set_cycle"

                if not target_col or not value_col:
                    st.error(
                        "El archivo debe contener la columna 'Equipo_Principal' (o 'Codigo_Componente') "
                        "y la columna 'Horas_a_Sumar' (o 'Horometro_Actual')."
                    )
                else:
                    st.info(
                        f"Columna detectada: **{target_col}** (`{target_type}`) | "
                        f"Columna de horas: **{value_col}** (Modo: `{'Sumar Horas' if mode_op == 'add' else 'Asignar Horómetro Ciclo'}`)"
                    )
                    if st.button(
                        "🚀 Actualizar Horas de Componentes Montados",
                        type="primary",
                        use_container_width=True,
                        key="btn_ejecutar_carga_archivo",
                    ):
                        updates_list = []
                        for _, r in df_up.iterrows():
                            val_id = r[target_col]
                            val_hr = r[value_col]
                            if pd.notna(val_id) and pd.notna(val_hr):
                                updates_list.append(
                                    {
                                        "target_type": target_type,
                                        "target_value": str(val_id).strip(),
                                        "mode": mode_op,
                                        "value": float(val_hr),
                                    }
                                )
                        total_upd, detalle_upd = apply_hours_to_mounted_components(
                            updates_list,
                            usuario_rol=st.session_state.user_role,
                            metodo_carga=f"Archivo ({archivo_horas.name})",
                        )
                        if total_upd > 0:
                            st.success(
                                f"✅ ¡Actualización completada! Se actualizaron **{total_upd}** componentes en estado 'Montado'."
                            )
                            st.dataframe(pd.DataFrame(detalle_upd), use_container_width=True, hide_index=True)
                        else:
                            st.warning(
                                "No se encontraron componentes en estado 'Montado' que coincidan con los registros o las horas a sumar fueron <= 0."
                            )
            except Exception as exc:
                st.error(f"Error al leer el archivo: {exc}")

    # --------------------------------------------------------------------------
    # OPCIÓN B: Actualización rápida por Flota / Equipo en pantalla
    # --------------------------------------------------------------------------
    with tab_b:
        st.markdown("#### ⚡ Opción B: Actualización Rápida por Flota / Equipo Principal o Tabla Editable")
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
            st.warning("Actualmente no hay componentes en estado 'Montado'.")
        else:
            col_f1, col_f2, col_f3 = st.columns([1.4, 1.4, 1.6])
            with col_f1:
                flotas_m = ["Todas"] + sorted(df_montados["flota"].dropna().unique().tolist())
                sel_flota_b = st.selectbox("1. Filtrar Flota:", flotas_m, key="flota_op_b")
            with col_f2:
                df_sub_b = (
                    df_montados
                    if sel_flota_b == "Todas"
                    else df_montados[df_montados["flota"] == sel_flota_b]
                )
                equipos_m = ["Todos"] + sorted(df_sub_b["equipo_principal"].dropna().unique().tolist())
                sel_eq_b = st.selectbox("2. Filtrar Equipo Principal (ej. SHO009):", equipos_m, key="eq_op_b")
            with col_f3:
                horas_bloque = st.number_input(
                    "3. Horas operadas en el periodo a sumar (+ hrs):",
                    min_value=0.0,
                    max_value=5000.0,
                    value=0.0,
                    step=10.0,
                    help="Si ingresas un valor > 0 (ej. +150 hrs), se pre-cargará para todos los componentes montados filtrados.",
                )

            df_edit_source = df_sub_b if sel_eq_b == "Todos" else df_sub_b[df_sub_b["equipo_principal"] == sel_eq_b]
            df_editor_input = df_edit_source.copy()
            df_editor_input["Horas_a_Sumar"] = float(horas_bloque)

            st.caption(
                "Puedes ajustar manualmente la columna **`Horas_a_Sumar`** en cualquier fila de la tabla editable antes de aplicar:"
            )
            edited_df = st.data_editor(
                df_editor_input,
                use_container_width=True,
                hide_index=True,
                disabled=[
                    "codigo_equipo",
                    "descripcion",
                    "flota",
                    "equipo_principal",
                    "ciclo_actual",
                    "horas_ciclo_actual",
                    "horas_acumuladas_totales",
                    "ultima_actualizacion_horas",
                ],
                column_config={
                    "codigo_equipo": "Código SAP",
                    "descripcion": "Descripción",
                    "flota": "Flota",
                    "equipo_principal": "Equipo Principal",
                    "ciclo_actual": "Ciclo #",
                    "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo Actual", format="%.1f"),
                    "horas_acumuladas_totales": st.column_config.NumberColumn("Hrs Totales", format="%.1f"),
                    "ultima_actualizacion_horas": "Últ. Actualización",
                    "Horas_a_Sumar": st.column_config.NumberColumn(
                        "➕ Horas a Sumar (+)",
                        min_value=0.0,
                        max_value=10000.0,
                        step=5.0,
                        format="%.1f hrs",
                    ),
                },
                key="editor_horas_montados",
            )

            if st.button("⚡ Aplicar Horas a Montados", type="primary", use_container_width=True):
                updates_b = []
                for _, row_ed in edited_df.iterrows():
                    hrs_add = float(row_ed.get("Horas_a_Sumar", 0.0) or 0.0)
                    if hrs_add > 0:
                        updates_b.append(
                            {
                                "target_type": "codigo_equipo",
                                "target_value": str(row_ed["codigo_equipo"]),
                                "mode": "add",
                                "value": hrs_add,
                            }
                        )
                if not updates_b:
                    st.warning("Ingresa una cantidad de 'Horas a Sumar' mayor a 0 en al menos un componente montado.")
                else:
                    n_upd, det_upd = apply_hours_to_mounted_components(
                        updates_b,
                        usuario_rol=st.session_state.user_role,
                        metodo_carga=f"Pantalla Rápida (Flota: {sel_flota_b} | Equipo: {sel_eq_b})",
                    )
                    st.success(f"✅ Se aplicaron las horas correctamente a **{n_upd}** componentes montados.")
                    st.dataframe(pd.DataFrame(det_upd), use_container_width=True, hide_index=True)

    # --------------------------------------------------------------------------
    # PESTAÑA 3: Log de auditoría de cargas de horas
    # --------------------------------------------------------------------------
    with tab_log:
        conn = get_connection()
        df_log = pd.read_sql_query("SELECT * FROM log_carga_horas ORDER BY id_carga DESC LIMIT 100", conn)
        conn.close()
        if df_log.empty:
            st.info("Aún no se han registrado cargas masivas de horas.")
        else:
            st.dataframe(df_log, use_container_width=True, hide_index=True)


# ==============================================================================
# 9. MÓDULO 3: GESTIÓN DE ESTADOS, KITS, OC Y TRAZABILIDAD
# ==============================================================================
def render_module_3_gestion_estados() -> None:
    st.markdown("### ⚡ Módulo 3: Gestión de Estados, Kits, OC y Trazabilidad por Ciclo")

    conn = get_connection()
    df_comp = pd.read_sql_query(
        "SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn
    )
    conn.close()

    if df_comp.empty:
        st.warning("No existen componentes cargados.")
        return

    opciones = [
        f"{r['codigo_equipo']} | {r['descripcion']} | Equipo: {r['equipo_principal']} | Estado: {r['estado_actual']} (Ciclo #{r['ciclo_actual']})"
        for _, r in df_comp.iterrows()
    ]
    sel_comp_str = st.selectbox("1. Seleccionar Componente a Gestionar:", opciones)
    cod_sel = sel_comp_str.split(" | ")[0].strip()
    comp_row = df_comp[df_comp["codigo_equipo"] == cod_sel].iloc[0]

    # Resumen actual del componente
    st.markdown(
        f"""
        <div class="argon-card">
            <b>Componente:</b> <code>{comp_row['codigo_equipo']}</code> — {comp_row['descripcion']} &nbsp;|&nbsp;
            <b>Estado Actual:</b> {get_status_badge_html(comp_row['estado_actual'])} &nbsp;|&nbsp;
            <b>Equipo Principal Actual:</b> <code>{comp_row['equipo_principal']}</code> &nbsp;|&nbsp;
            <b>Ciclo Actual:</b> #{int(comp_row['ciclo_actual'])} &nbsp;|&nbsp;
            <b>Horas Ciclo:</b> {float(comp_row['horas_ciclo_actual']):,.1f} hrs &nbsp;|&nbsp;
            <b>Horas Totales:</b> {float(comp_row['horas_acumuladas_totales']):,.1f} hrs
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_est, col_op = st.columns(2)
    with col_est:
        nuevo_estado = st.selectbox(
            "2. Nuevo Estado a Asignar al Componente:",
            options=ESTADOS_COMPONENTE,
            index=ESTADOS_COMPONENTE.index(comp_row["estado_actual"])
            if comp_row["estado_actual"] in ESTADOS_COMPONENTE
            else 0,
        )
    with col_op:
        equipo_destino = st.text_input(
            "Equipo Principal Asociado (ej. SHO009, FLO005, TALLER):",
            value=str(comp_row["equipo_principal"]),
        )

    # Variables por defecto del formulario inteligente
    instalado_por = "N/A"
    reiniciar_ciclo = False
    motivo_cambio = "N/A"
    modo_falla = "N/A"
    numero_ot = ""
    numero_oc = ""
    proveedor_reparador = ""
    tipo_intervencion = "Ninguna"
    detalle_kit = ""
    costo_kit_usd = 0.0
    costo_reparacion_usd = 0.0

    st.markdown("#### 🧠 Parámetros Inteligentes según el Estado Seleccionado")

    # A) Si se MONTA
    if nuevo_estado == "Montado":
        st.info("🔧 **Modo Montaje de Componente:** Registra quién ejecutó el montaje y el control de inicio de nuevo ciclo.")
        m1, m2 = st.columns(2)
        with m1:
            instalado_por = st.radio(
                "¿Quién instaló el componente en el equipo?",
                options=["Southern", "Contratista"],
                horizontal=True,
            )
        with m2:
            reiniciar_ciclo = st.checkbox(
                f"Iniciar Nuevo Ciclo (Incrementar a Ciclo #{int(comp_row['ciclo_actual']) + 1} y reiniciar Horas/Costo de Ciclo a 0)",
                value=(comp_row["estado_actual"] != "Montado"),
            )

    # B) Si se DESMONTA o pasa a POR REPARAR / SALIDA POR RECLAMO / DE BAJA
    if nuevo_estado in ("Desmontado", "Por Reparar", "Salida por Reclamo", "De Baja"):
        st.warning("🛠️ **Modo Desmontaje / Salida de Operación:** Especifica la causa técnica del retiro del componente.")
        d1, d2 = st.columns(2)
        with d1:
            motivo_cambio = st.selectbox(
                "Motivo del Desmontaje / Cambio:",
                options=["PCR por Horas", "Falla por Condición", "Reclamo Garantía"],
            )
        with d2:
            modo_falla = st.selectbox(
                "Modo de Falla / Condición Detectada:",
                options=MODOS_FALLA_COMUNES,
                index=0 if motivo_cambio == "PCR por Horas" else 1,
            )

    # C) Si se REPARA o se registra intervención técnica/financiera (Reparado, Por Reparar, Salida por Reclamo o Kit en Montado)
    mostrar_bloque_reparacion = nuevo_estado in ("Reparado", "Por Reparar", "Salida por Reclamo") or st.checkbox(
        "➕ Agregar Costos de Kit / OT / Orden de Compra (OC) en este evento", value=False
    )

    if mostrar_bloque_reparacion:
        st.markdown("#### 💰 Trazabilidad de OT, Orden de Compra (OC), Kits y Costos de Reparación")
        r1, r2, r3 = st.columns(3)
        with r1:
            numero_ot = st.text_input("Número de OT SAP:", placeholder="Ej. OT-40125901")
        with r2:
            numero_oc = st.text_input("Número de Orden de Compra (OC):", placeholder="Ej. 4500231099")
        with r3:
            proveedor_reparador = st.text_input(
                "Proveedor / Taller Reparador:",
                placeholder="Ej. KOMATSU / FERREYROS / TALLER SPCC",
            )

        k1, k2, k3 = st.columns([1.4, 1.3, 1.3])
        with k1:
            tipo_intervencion = st.selectbox(
                "Tipo de Intervención Financiera:",
                options=[
                    "Kit + Reparación General",
                    "Kit de Mantenimiento",
                    "Reparación General",
                    "Ninguna",
                ],
            )
        with k2:
            costo_kit_usd = st.number_input(
                "Costo de Kit Aplicado (USD):",
                min_value=0.0,
                max_value=2000000.0,
                value=0.0,
                step=100.0,
            )
        with k3:
            costo_reparacion_usd = st.number_input(
                "Costo Reparación General (USD):",
                min_value=0.0,
                max_value=2000000.0,
                value=0.0,
                step=250.0,
            )

        detalle_kit = st.text_input(
            "Detalle de Kit Aplicado / Alcance Técnico de la Reparación:",
            placeholder="Ej. Kit de sellos viton, rodamientos serie pesada, rectificado de alojamiento...",
        )

    st.markdown("#### 📝 Bitácora y Evidencia Fotográfica Opcional (Regla FIFO 2 Fotos)")
    col_com, col_fot = st.columns([1.6, 1.4])
    with col_com:
        comentario_bitacora = st.text_area(
            "Comentario Técnico para la Bitácora:",
            placeholder="Describe hallazgos técnicos, presiones, holguras, motivo del cambio o acuerdos con proveedor...",
            height=110,
        )
    with col_fot:
        foto_evento = st.file_uploader(
            "Adjuntar Fotografía del Evento (Opcional - aplica FIFO 2 fotos):",
            type=["jpg", "jpeg", "png", "webp"],
            key="foto_mod3",
        )

    costo_total_evento = round(float(costo_kit_usd) + float(costo_reparacion_usd), 2)
    st.markdown(
        f"**💵 Costo Total Calculado de este Evento:** `${costo_total_evento:,.2f} USD` "
        f"(Kit: `${costo_kit_usd:,.2f}` + Reparación General: `${costo_reparacion_usd:,.2f}`)"
    )

    if st.button("💾 Guardar Movimiento, Estado y Costos en Base de Datos", type="primary", use_container_width=True):
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
            SET estado_actual = ?,
                equipo_principal = ?,
                ciclo_actual = ?,
                horas_ciclo_actual = ?,
                gasto_ciclo_actual_usd = ?,
                gasto_acumulado_total_usd = ?
            WHERE codigo_equipo = ?
            """,
            (
                nuevo_estado,
                equipo_destino.strip(),
                nuevo_ciclo_num,
                round(nuevas_horas_ciclo, 2),
                round(nuevo_gasto_ciclo, 2),
                round(nuevo_gasto_total, 2),
                cod_sel,
            ),
        )

        # Si subió foto, procesarla con FIFO sin duplicar evento en bitácora
        saved_photo_path = None
        msg_fifo = ""
        if foto_evento is not None:
            conn.commit()
            conn.close()
            ok_p, msg_fifo, saved_photo_path = save_component_photo_fifo(
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
                cod_sel,
                nuevo_ciclo_num,
                now_str,
                st.session_state.user_role,
                f"Cambio de Estado / Gestión ({nuevo_estado})",
                nuevo_estado,
                equipo_destino.strip(),
                round(nuevas_horas_ciclo, 2),
                instalado_por,
                motivo_cambio,
                modo_falla,
                numero_ot.strip(),
                numero_oc.strip(),
                proveedor_reparador.strip(),
                tipo_intervencion,
                detalle_kit.strip(),
                round(float(costo_kit_usd), 2),
                round(float(costo_reparacion_usd), 2),
                costo_total_evento,
                (comentario_bitacora or f"Actualización a estado {nuevo_estado}").strip(),
                saved_photo_path,
            ),
        )

        conn.commit()
        conn.close()
        st.success(
            f"✅ Componente `{cod_sel}` actualizado a **{nuevo_estado}** (Ciclo #{nuevo_ciclo_num}). "
            f"Costo evento registrado: **${costo_total_evento:,.2f} USD**. {msg_fifo}"
        )


# ==============================================================================
# 10. MÓDULO 4: MODO TABLET CAMPO: FOTOS Y COMENTARIOS (FIFO ESTRICTO)
# ==============================================================================
def render_module_4_tablet_campo() -> None:
    st.markdown("### 📱 Módulo 4: Modo Tablet Campo — Inspección Visual, Fotos FIFO y Comentarios")
    st.markdown(
        """
        <div class="argon-card" style="border-left: 5px solid #11cdef;">
            <b>Optimizado para Terreno / Pampa / Bahía de Mantenimiento:</b> Cada componente conserva como máximo sus
            <b>2 últimas fotografías</b> (<code>foto_1_path</code> y <code>foto_2_path</code>). Al registrar una 3ra foto,
            el sistema elimina físicamente del disco la foto #1, desplaza la #2 a #1 y guarda la nueva imagen comprimida en #2.
        </div>
        """,
        unsafe_allow_html=True,
    )

    conn = get_connection()
    df_comp = pd.read_sql_query(
        "SELECT * FROM componentes ORDER BY equipo_principal, codigo_equipo", conn
    )
    conn.close()

    if df_comp.empty:
        st.warning("No hay componentes disponibles.")
        return

    c_filtro1, c_filtro2 = st.columns([1.2, 2.0])
    with c_filtro1:
        equipos_disp = ["Todos"] + sorted(df_comp["equipo_principal"].dropna().unique().tolist())
        eq_tablet = st.selectbox("Filtrar por Equipo Principal:", equipos_disp)
    with c_filtro2:
        df_t = df_comp if eq_tablet == "Todos" else df_comp[df_comp["equipo_principal"] == eq_tablet]
        opciones_t = [
            f"{r['codigo_equipo']} — {r['descripcion']} ({r['equipo_principal']}) [{r['estado_actual']}]"
            for _, r in df_t.iterrows()
        ]
        sel_t = st.selectbox("Seleccionar Componente en Campo:", opciones_t)

    cod_tablet = sel_t.split(" — ")[0].strip()
    comp_t = df_comp[df_comp["codigo_equipo"] == cod_tablet].iloc[0]

    # Mostrar estado actual de las 2 ranuras fotográficas
    col_actual1, col_actual2 = st.columns(2)
    with col_actual1:
        st.markdown("##### 🖼️ Ranura #1 Actual (Foto Anterior)")
        if comp_t["foto_1_path"] and os.path.exists(comp_t["foto_1_path"]):
            st.image(comp_t["foto_1_path"], use_container_width=True)
            st.caption(f"📅 {comp_t['foto_1_fecha']} | 💬 {comp_t['foto_1_comentario']}")
        else:
            st.info("Ranura #1 Libre")

    with col_actual2:
        st.markdown("##### 🖼️ Ranura #2 Actual (Foto Más Reciente)")
        if comp_t["foto_2_path"] and os.path.exists(comp_t["foto_2_path"]):
            st.image(comp_t["foto_2_path"], use_container_width=True)
            st.caption(f"📅 {comp_t['foto_2_fecha']} | 💬 {comp_t['foto_2_comentario']}")
        else:
            st.info("Ranura #2 Libre")

    st.markdown("---")
    st.markdown("#### 📷 Capturar o Subir Nueva Evidencia Fotográfica / Observación de Campo")

    metodo_captura = st.radio(
        "Origen de la Fotografía:",
        options=["📁 Subir imagen desde Galería / Tablet", "📸 Tomar Foto con Cámara del Dispositivo"],
        horizontal=True,
    )

    if metodo_captura.startswith("📁"):
        img_input = st.file_uploader(
            "Seleccionar fotografía (JPG, PNG, WEBP):",
            type=["jpg", "jpeg", "png", "webp"],
            key="tablet_uploader",
        )
    else:
        img_input = st.camera_input("Capturar fotografía directa del componente:")

    comentario_campo = st.text_area(
        "Comentario de Inspección en Campo (se guarda junto a la foto y en la Bitácora):",
        placeholder="Ej. Se inspecciona componente en pampa; dentado sin fisuras, tapones magnéticos limpios, cubierto con plástico protector...",
        height=100,
    )

    col_btn_foto, col_btn_solo_txt = st.columns(2)
    with col_btn_foto:
        if st.button("📸 Guardar Nueva Foto (Rotación FIFO) + Comentario", type="primary", use_container_width=True):
            if img_input is None:
                st.error("Por favor selecciona o captura una fotografía primero.")
            else:
                ok, msg, _ = save_component_photo_fifo(
                    codigo_equipo=cod_tablet,
                    uploaded_file=img_input,
                    comentario_foto=comentario_campo,
                    usuario_rol=st.session_state.user_role,
                    registrar_en_bitacora=True,
                )
                if ok:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(msg)

    with col_btn_solo_txt:
        if st.button("📝 Registrar Solo Comentario en Bitácora (Sin Foto)", use_container_width=True):
            if not comentario_campo.strip():
                st.warning("Escribe un comentario antes de registrar en la bitácora.")
            else:
                conn = get_connection()
                cur = conn.cursor()
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cur.execute(
                    """
                    INSERT INTO historial_eventos (
                        codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                        tipo_operacion, estado_asignado, equipo_principal,
                        horas_componente_evento, comentario_bitacora
                    ) VALUES (?, ?, ?, ?, 'Nota de Campo (Tablet)', ?, ?, ?, ?)
                    """,
                    (
                        cod_tablet,
                        comp_t["ciclo_actual"],
                        now_str,
                        st.session_state.user_role,
                        comp_t["estado_actual"],
                        comp_t["equipo_principal"],
                        comp_t["horas_ciclo_actual"],
                        comentario_campo.strip(),
                    ),
                )
                conn.commit()
                conn.close()
                st.success("✅ Comentario de campo asentado en el historial del componente.")


# ==============================================================================
# 11. MÓDULO 5: ALTA DE COMPONENTES, SINCRONIZACIÓN SAP Y BACKUP EXCEL
# ==============================================================================
def render_module_5_alta_sap_backup() -> None:
    st.markdown("### ➕ Módulo 5: Alta de Componentes, Sincronización SAP PM y Backup Completo")

    # Bloque superior de Respaldo Inmediato a Excel (.xlsx)
    st.markdown(
        """
        <div class="argon-card" style="border-left: 5px solid #5e72e4;">
            <h4 style="margin:0 0 0.35rem 0; color:#32325d;">🛡️ Respaldo Corporativo Inmediato (.xlsx Multi-Pestaña)</h4>
            <p style="margin:0; font-size:0.88rem; color:#525f7f;">
                Descarga en un solo clic toda la base de datos SQLite en formato Excel estructurado con 3 hojas:
                <b>Inventario_Actual</b>, <b>Historial_Eventos_Costos</b> y <b>Bitacora_Carga_Horas</b>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    excel_bytes = export_full_database_to_excel()
    timestamp_dl = datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        label="📥 Descargar Base de Datos Completa a Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"Backup_Inventario_Mina_SAP_{timestamp_dl}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.markdown("---")
    tab_sap, tab_alta = st.tabs(
        [
            "🔄 Sincronización Maestro SAP ('Inventario SAP Nov.xlsx')",
            "➕ Alta Manual de Nuevo Componente",
        ]
    )

    with tab_sap:
        st.markdown("#### 🔄 Importar / Sincronizar Excel de SAP PM (`Inventario SAP Nov.xlsx`)")
        st.write(
            "Columnas esperadas en la hoja **`Sheet1`**: "
            "`['Equipo', 'Denominación', 'Indicador ABC', 'Ubicac.técnica', 'Denominación3', 'Fabricante', 'Fabr. Nº-serie', 'Válido de', 'Fin de validez']`."
        )
        st.write(
            "Regla de negocio aplicada automáticamente al importar: "
            "del **4to bloque** de `Ubicac.técnica` separado por guiones se extrae la **`flota`** (ej. `495H`, `410A`, `994F`) "
            "y del **5to bloque** el **`equipo_principal`** (ej. `FLO005`, `SHO009`)."
        )

        c_dl, c_local = st.columns(2)
        with c_dl:
            st.download_button(
                "📄 Descargar Plantilla Muestra 'Inventario SAP Nov.xlsx'",
                data=generate_sap_template_excel(),
                file_name="Inventario SAP Nov.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c_local:
            if os.path.exists(DEFAULT_SAP_EXCEL):
                if st.button("⚡ Sincronizar desde 'Inventario SAP Nov.xlsx' en Carpeta Local", use_container_width=True):
                    try:
                        df_local = pd.read_excel(DEFAULT_SAP_EXCEL, sheet_name="Sheet1")
                        n_new, n_upd, msg = import_sap_dataframe(df_local, st.session_state.user_role)
                        st.success(f"✅ {msg} (Nuevos: {n_new} | Actualizados: {n_upd})")
                    except Exception as exc:
                        st.error(f"Error al leer '{DEFAULT_SAP_EXCEL}': {exc}")

        sap_file = st.file_uploader(
            "O sube aquí tu archivo Excel de SAP ('Inventario SAP Nov.xlsx'):",
            type=["xlsx", "xls"],
            key="sap_excel_uploader",
        )
        if sap_file is not None:
            try:
                xl = openpyxl.load_workbook(sap_file, read_only=True)
                sheet_to_use = "Sheet1" if "Sheet1" in xl.sheetnames else xl.sheetnames[0]
                sap_file.seek(0)
                df_sap_up = pd.read_excel(sap_file, sheet_name=sheet_to_use)
                st.caption(f"Hoja leída: `{sheet_to_use}` ({len(df_sap_up)} filas detectadas)")
                st.dataframe(df_sap_up.head(10), use_container_width=True, hide_index=True)

                if st.button("🚀 Ejecutar Sincronización SAP Ahora", type="primary", use_container_width=True):
                    n_new, n_upd, msg = import_sap_dataframe(df_sap_up, st.session_state.user_role)
                    if n_new + n_upd > 0:
                        st.success(f"✅ {msg} — Nuevos componentes: **{n_new}** | Componentes actualizados: **{n_upd}**.")
                    else:
                        st.error(msg)
            except Exception as exc:
                st.error(f"Error procesando archivo SAP: {exc}")

    with tab_alta:
        st.markdown("#### ➕ Registro Manual de Componente Individual")
        with st.form("form_alta_manual", clear_on_submit=True):
            a1, a2, a3 = st.columns(3)
            with a1:
                cod_m = st.text_input("Código Equipo SAP *:", placeholder="Ej. 10045990")
                ubic_m = st.text_input(
                    "Ubicación Técnica SAP *:",
                    value="SPCC-TOQ-MIN-495H-SHO009-GIRO",
                    help="El 4to bloque define la Flota y el 5to bloque el Equipo Principal.",
                )
                crit_m = st.selectbox("Criticidad ABC:", options=["A", "B", "C"], index=0)
            with a2:
                desc_m = st.text_input("Descripción del Componente *:", placeholder="Ej. MOTOR DE GIRO TRASERO")
                fab_m = st.text_input("Fabricante:", placeholder="Ej. KOMATSU / SIEMENS / CAT")
                serie_m = st.text_input("Número de Serie Fabricante:", placeholder="Ej. SN-884920")
            with a3:
                est_m = st.selectbox("Estado Inicial:", options=ESTADOS_COMPONENTE, index=0)
                hrs_ini = st.number_input("Horas Iniciales (Ciclo y Acumuladas):", min_value=0.0, value=0.0, step=50.0)
                costo_ini = st.number_input("Costo Inicial Acumulado (USD):", min_value=0.0, value=0.0, step=100.0)

            submitted = st.form_submit_button("➕ Dar de Alta Componente", use_container_width=True)
            if submitted:
                if not cod_m.strip() or not desc_m.strip():
                    st.error("El Código de Equipo SAP y la Descripción son obligatorios.")
                else:
                    flota_auto, eq_auto = parse_ubicacion_tecnica(ubic_m)
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
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
                            """,
                            (
                                cod_m.strip(),
                                desc_m.strip().upper(),
                                flota_auto,
                                eq_auto,
                                ubic_m.strip().upper(),
                                crit_m,
                                (fab_m.strip() or "S/I").upper(),
                                (serie_m.strip() or "S/N").upper(),
                                est_m,
                                round(float(hrs_ini), 2),
                                round(float(hrs_ini), 2),
                                round(float(costo_ini), 2),
                                round(float(costo_ini), 2),
                                now_str,
                            ),
                        )
                        cur.execute(
                            """
                            INSERT INTO historial_eventos (
                                codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,
                                tipo_operacion, estado_asignado, equipo_principal,
                                horas_componente_evento, costo_total_evento_usd, comentario_bitacora
                            ) VALUES (?, 1, ?, ?, 'Alta Inicial de Componente', ?, ?, ?, ?, ?)
                            """,
                            (
                                cod_m.strip(),
                                now_str,
                                st.session_state.user_role,
                                est_m,
                                eq_auto,
                                round(float(hrs_ini), 2),
                                round(float(costo_ini), 2),
                                f"Alta manual. Flota extraída: {flota_auto} | Equipo: {eq_auto}",
                            ),
                        )
                        conn.commit()
                        st.success(
                            f"✅ Componente `{cod_m.strip()}` creado exitosamente (Flota detectada: **{flota_auto}**, Equipo: **{eq_auto}**)."
                        )
                    except sqlite3.IntegrityError:
                        st.error(f"El código de equipo `{cod_m.strip()}` ya existe en la base de datos.")
                    finally:
                        conn.close()


# ==============================================================================
# 12. MÓDULO 6: ESTRATEGIA DE COMPONENTES Y COSTO POR CICLO (PLOTLY ANALYTICS)
# ==============================================================================
def render_module_6_estrategia_costos() -> None:
    st.markdown("### 📊 Módulo 6: Estrategia de Componentes, Rendimiento y Costo por Ciclo ($/hr)")

    conn = get_connection()
    df_comp = pd.read_sql_query("SELECT * FROM componentes", conn)
    df_hist = pd.read_sql_query("SELECT * FROM historial_eventos", conn)
    conn.close()

    if df_comp.empty:
        st.warning("No hay datos suficientes para generar analítica.")
        return

    df_comp["costo_por_hora_usd"] = df_comp.apply(
        lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2)
        if r["horas_acumuladas_totales"] > 0
        else 0.0,
        axis=1,
    )

    col_g1, col_g2 = st.columns(2)

    with col_g1:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("#### 💵 Costo Promedio por Hora ($/hr) y Gasto por Flota")
        df_flota = (
            df_comp.groupby("flota", as_index=False)
            .agg(
                gasto_total=("gasto_acumulado_total_usd", "sum"),
                horas_totales=("horas_acumuladas_totales", "sum"),
                equipos=("codigo_equipo", "count"),
            )
        )
        df_flota["costo_hr"] = df_flota.apply(
            lambda r: round(r["gasto_total"] / r["horas_totales"], 2) if r["horas_totales"] > 0 else 0.0,
            axis=1,
        )
        fig_flota = px.bar(
            df_flota,
            x="flota",
            y="costo_hr",
            color="flota",
            text="costo_hr",
            hover_data=["gasto_total", "horas_totales", "equipos"],
            color_discrete_sequence=["#5e72e4", "#2dce89", "#fb6340", "#f5365c", "#11cdef"],
            labels={"flota": "Flota Minera", "costo_hr": "Costo por Hora ($/hr)"},
        )
        fig_flota.update_traces(texttemplate="$%{text:.2f}/hr", textposition="outside")
        fig_flota.update_layout(
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            showlegend=False,
        )
        st.plotly_chart(fig_flota, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_g2:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("#### 🎯 Matriz Horas Acumuladas vs. Gasto Total (Tamaño = Ciclos)")
        fig_scatter = px.scatter(
            df_comp,
            x="horas_acumuladas_totales",
            y="gasto_acumulado_total_usd",
            color="estado_actual",
            size="ciclo_actual",
            hover_name="descripcion",
            hover_data=["codigo_equipo", "flota", "equipo_principal", "costo_por_hora_usd"],
            color_discrete_map={
                "Montado": "#2dce89",
                "Reparado": "#5e72e4",
                "Por Reparar": "#fb6340",
                "Desmontado": "#8898aa",
                "De Baja": "#f5365c",
                "Nuevo": "#11cdef",
                "Salida por Reclamo": "#ffd600",
            },
            labels={
                "horas_acumuladas_totales": "Horas Acumuladas Totales (hrs)",
                "gasto_acumulado_total_usd": "Gasto Acumulado Total ($ USD)",
                "estado_actual": "Estado",
            },
        )
        fig_scatter.update_layout(
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Segunda fila analítica: Desglose Kits vs Reparación General por Ciclo y Ranking Crítico
    col_g3, col_g4 = st.columns(2)

    with col_g3:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("#### 🛠️ Inversión en Kits vs. Reparación General por Número de Ciclo")
        if not df_hist.empty and df_hist["costo_total_evento_usd"].sum() > 0:
            df_ciclo_cost = (
                df_hist.groupby("numero_ciclo", as_index=False)
                .agg(
                    Costo_Kits=("costo_kit_usd", "sum"),
                    Costo_Reparacion_General=("costo_reparacion_usd", "sum"),
                )
            )
            fig_ciclo = go.Figure()
            fig_ciclo.add_trace(
                go.Bar(
                    x=df_ciclo_cost["numero_ciclo"].astype(str),
                    y=df_ciclo_cost["Costo_Kits"],
                    name="Kits de Mantenimiento ($)",
                    marker_color="#5e72e4",
                )
            )
            fig_ciclo.add_trace(
                go.Bar(
                    x=df_ciclo_cost["numero_ciclo"].astype(str),
                    y=df_ciclo_cost["Costo_Reparacion_General"],
                    name="Reparación General ($)",
                    marker_color="#fb6340",
                )
            )
            fig_ciclo.update_layout(
                barmode="stack",
                height=340,
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis_title="Número de Ciclo del Componente",
                yaxis_title="Costo Acumulado ($ USD)",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_ciclo, use_container_width=True)
        else:
            st.info("Sin costos históricos suficientes para graficar por ciclo.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_g4:
        st.markdown('<div class="argon-card">', unsafe_allow_html=True)
        st.markdown("#### 🚨 Top Componentes con Mayor Costo por Hora ($/hr) — Evaluación PCR / Baja")
        df_ranking = df_comp.sort_values(by="costo_por_hora_usd", ascending=False).head(10)
        st.dataframe(
            df_ranking[
                [
                    "codigo_equipo",
                    "descripcion",
                    "flota",
                    "equipo_principal",
                    "ciclo_actual",
                    "horas_acumuladas_totales",
                    "gasto_acumulado_total_usd",
                    "costo_por_hora_usd",
                    "estado_actual",
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "codigo_equipo": "Código SAP",
                "descripcion": "Componente",
                "flota": "Flota",
                "equipo_principal": "Equipo",
                "ciclo_actual": "Ciclo #",
                "horas_acumuladas_totales": st.column_config.NumberColumn("Hrs Totales", format="%.0f"),
                "gasto_acumulado_total_usd": st.column_config.NumberColumn("Gasto Total ($)", format="$%.0f"),
                "costo_por_hora_usd": st.column_config.NumberColumn("$/hr", format="$%.2f/hr"),
                "estado_actual": "Estado",
            },
        )
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# 13. PUNTO DE ENTRADA PRINCIPAL (MAIN ROUTER)
# ==============================================================================
def main() -> None:
    inject_argon_css()
    init_db()
    render_sidebar_and_header()

    mod = st.session_state.current_module
    role = st.session_state.user_role

    # Protección de permisos por rol
    if role == "Campo-Tablet" and mod in ("MOD_2", "MOD_3", "MOD_5", "MOD_6"):
        st.warning("El rol 'Campo-Tablet' tiene acceso exclusivo al Buscador 360° (Módulo 1) y Modo Tablet Campo (Módulo 4).")
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