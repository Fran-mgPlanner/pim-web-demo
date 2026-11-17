# -*- coding: utf-8 -*-

"""

P.I.M. - CONTROL DE INVENTARIOS | PLANEAMIENTO MANTENIMIENTO MINA

Versión Producción Auditada:

- Cero errores de renderizado CSS/HTML (todo CSS/HTML inyectado en línea única sin saltos ni sangrías).

- Navegación directa de un solo camino mediante 4 pestañas nativas:

  1. Visualizar (Flota, KPIs, Tarjetas y Hoja de Vida con Fotos FIFO)

  2. Actualizar y Fotos (Foto FIFO, Cambio de Estado/OC/Kit y Carga Masiva de Horas)

  3. Agregar y SAP (Alta individual e importación de Inventario SAP Nov.xlsx)

  4. Estrategia y Backup (Gráfico $/hr por flota, Ranking y Descarga Excel .xlsx / SQLite .db)

"""



import io

import os

import sqlite3

from datetime import datetime

from typing import Dict, List, Optional, Tuple



import openpyxl

import pandas as pd

import plotly.express as px

import streamlit as st

from PIL import Image



# ==============================================================================

# 0. CONFIGURACIÓN INICIAL Y CONSTANTES

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

# 1. ESTILO LIMPIO EN LÍNEA ÚNICA (IMPOSIBLE QUE SE ROMPA EN MARKDOWN)

# ==============================================================================

def inject_clean_css() -> None:

    css_rules = (

        ".block-container{padding-top:3.2rem !important;padding-bottom:2rem !important;max-width:1280px !important;}"

        ".pim-header{background:linear-gradient(90deg,#1e293b 0%,#334155 100%);color:#ffffff;padding:1rem 1.4rem;border-radius:12px;margin-bottom:1rem;box-shadow:0 4px 12px rgba(0,0,0,0.08);}"

        ".pim-header h2{color:#ffffff !important;margin:0;font-size:1.35rem;font-weight:800;letter-spacing:0.02em;}"

        ".pim-header p{color:#cbd5e1 !important;margin:0.15rem 0 0 0;font-size:0.84rem;}"

        ".flota-tag{display:inline-block;padding:0.18rem 0.55rem;border-radius:999px;font-size:0.74rem;font-weight:700;margin-right:0.35rem;}"

        ".estado-tag{display:inline-block;padding:0.18rem 0.55rem;border-radius:999px;font-size:0.74rem;font-weight:700;}"

        "div.stButton > button, div.stDownloadButton > button{border-radius:10px !important;font-weight:600 !important;}"

    )

    st.markdown(f"&lt;style&gt;{css_rules}&lt;/style&gt;", unsafe_allow_html=True)





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

        return False, None, f"El campo '{field_name}' debe ser numérico."

    if not allow_negative and num < 0:

        return False, None, f"El campo '{field_name}' no permite negativos."

    if not allow_zero and num == 0:

        return False, None, f"El campo '{field_name}' debe ser mayor a 0."

    if not allow_decimal and not num.is_integer():

        return False, None, f"El campo '{field_name}' debe ser un número entero."

    return True, num, ""





def get_badges_line_html(flota: str, estado: str, vidas: int) -> str:

    """Devuelve una sola línea HTML sin saltos para identificar Flota, Estado y Vidas."""

    flota_clean = str(flota or "SIN_FLOTA").strip()

    flota_colors = {

        "495H": ("#fee2e2", "#b91c1c"),

        "410A": ("#dbeafe", "#1d4ed8"),

        "994F": ("#dcfce7", "#15803d"),

        "793F": ("#fef3c7", "#b45309"),

    }

    f_bg, f_fg = flota_colors.get(flota_clean, ("#f1f5f9", "#334155"))



    est_colors = {

        "Montado": ("#dcfce7", "#15803d"),

        "Reparado": ("#dbeafe", "#1d4ed8"),

        "Nuevo": ("#e0f2fe", "#0369a1"),

        "Por Reparar": ("#ffedd5", "#c2410c"),

        "Salida por Reclamo": ("#fef3c7", "#b45309"),

        "Desmontado": ("#f1f5f9", "#475569"),

        "De Baja": ("#fee2e2", "#b91c1c"),

    }

    e_bg, e_fg = est_colors.get(estado, ("#f1f5f9", "#475569"))

    v_lbl = format_vidas_label(vidas)



    return (

        f'<span class="flota-tag" style="background:{f_bg};color:{f_fg};">🚜 Flota {flota_clean}</span>'

        f'<span class="estado-tag" style="background:{e_bg};color:{e_fg};">● {estado}</span> '

        f'<span class="estado-tag" style="background:#f1f5f9;color:#334155;">🔄 {v_lbl}</span>'

    )





# ==============================================================================

# 2. BASE DE DATOS SQLITE3 (MODO WAL PARA CONCURRENCIA EN RED LOCAL)

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





# ==============================================================================

# 3. FUNCIONES OPERATIVAS (FOTOS FIFO, CARGA HORAS, SAP Y BACKUP)

# ==============================================================================

def save_component_photo_fifo(

    codigo_equipo: str,

    uploaded_file,

    comentario_foto: str,

    usuario_rol: str,

    registrar_en_bitacora: bool = True,

) -> Tuple[bool, str, Optional[str]]:

    if uploaded_file is None:

        return False, "No se seleccionó ninguna imagen.", None



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

            msg_rotacion = "Fotografía guardada en Espacio #1."

        elif not f2_path:

            cur.execute(

                "UPDATE componentes SET foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ? WHERE codigo_equipo = ?",

                (new_photo_path, now_display, comentario_limpio, codigo_equipo),

            )

            msg_rotacion = "Fotografía guardada en Espacio #2."

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

            msg_rotacion = "Rotación FIFO completada: se eliminó la foto #1, la #2 pasó a #1 y la nueva quedó en #2."



        if registrar_en_bitacora:

            cur.execute(

                """

                INSERT INTO historial_eventos (

                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                    tipo_operacion, estado_asignado, equipo_principal,

                    horas_componente_evento, comentario_bitacora, foto_path

                ) VALUES (?, ?, ?, ?, 'Inspección Fotográfica', ?, ?, ?, ?, ?)

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

        if not cur.fetchone():

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

            (now_str, usuario_rol, metodo_carga, len(afectados_detalle), f"Actualizados {len(afectados_detalle)} componentes montados."),

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

# 4. CABECERA Y FILTRO GLOBAL POR FLOTA

# ==============================================================================

def init_session_state() -> None:

    defaults = {

        "user_role": "Planner",

        "active_flota": "Todas",

        "selected_component": None,

        "vida_idx": 1,

        "vida_comp_ref": None,

    }

    for k, v in defaults.items():

        if k not in st.session_state:

            st.session_state[k] = v





def render_top_bar(df_all: pd.DataFrame) -> pd.DataFrame:

    st.markdown(

        '&lt;div class=&quot;pim-header&quot;&gt;&lt;h2&gt;⚙️ P.I.M. — CONTROL DE INVENTARIOS&lt;/h2&gt;&lt;p&gt;PLANEAMIENTO MANTENIMIENTO MINA · Control por Flota y Trazabilidad SAP PM&lt;/p&gt;&lt;/div&gt;',

        unsafe_allow_html=True,

    )



    flotas = ["Todas"] + sorted(df_all["flota"].dropna().unique().tolist())

    if st.session_state.active_flota not in flotas:

        st.session_state.active_flota = "Todas"



    c_flota, c_rol = st.columns([2.2, 1.2])

    with c_flota:

        st.selectbox("🚜 Flota de trabajo (Espacio del Planner):", options=flotas, key="active_flota")

    with c_rol:

        st.selectbox("👤 Perfil:", options=ROLES_SISTEMA, key="user_role")



    return df_all if st.session_state.active_flota == "Todas" else df_all[df_all["flota"] == st.session_state.active_flota]





# ==============================================================================

# 5. SECCIÓN 1: VISUALIZAR (TARJETAS NATIVAS + HOJA DE VIDA + FOTOS)

# ==============================================================================

def render_tab_visualizar(df_flota: pd.DataFrame) -> None:

    if df_flota.empty:

        st.info("No hay componentes en esta flota.")

        return



    total = len(df_flota)

    montados = int((df_flota["estado_actual"] == "Montado").sum())

    taller = int(df_flota["estado_actual"].isin(["Por Reparar", "Desmontado", "Salida por Reclamo"]).sum())

    gasto = float(df_flota["gasto_acumulado_total_usd"].sum())

    hrs = float(df_flota["horas_acumuladas_totales"].sum())

    c_hr = (gasto / hrs) if hrs > 0 else 0.0



    k1, k2, k3, k4 = st.columns(4)

    k1.metric("Componentes", f"{total}", f"Flota: {st.session_state.active_flota}")

    k2.metric("Montados", f"{montados}", "En operación")

    k3.metric("Por Reparar / Taller", f"{taller}", "Pendientes")

    k4.metric("Gasto Acumulado", f"${gasto:,.0f}", f"${c_hr:,.2f} USD/hr")



    f1, f2, f3 = st.columns([1.2, 1.2, 2.0])

    with f1:

        sel_eq = st.selectbox("Equipo Principal:", ["Todos"] + sorted(df_flota["equipo_principal"].dropna().unique().tolist()), key="vis_eq")

    with f2:

        sel_est = st.selectbox("Estado:", ["Todos"] + ESTADOS_COMPONENTE, key="vis_est")

    with f3:

        q = st.text_input("Buscar componente:", placeholder="Código SAP, descripción o serie...", key="vis_q")



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

        st.warning("Sin coincidencias para este filtro.")

        return



    # Tarjetas limpias usando contenedores nativos de Streamlit (100% libres de errores HTML)

    rows_list = [r for _, r in df_filt.head(15).iterrows()]

    for idx in range(0, len(rows_list), 3):

        cols = st.columns(3)

        for c_i in range(3):

            if idx + c_i < len(rows_list):

                r = rows_list[idx + c_i]

                cod_item = str(r["codigo_equipo"])

                c_hr_item = (r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"]) if r["horas_acumuladas_totales"] > 0 else 0.0

                with cols[c_i]:

                    with st.container(border=True):

                        st.markdown(get_badges_line_html(r["flota"], r["estado_actual"], r["ciclo_actual"]), unsafe_allow_html=True)

                        st.markdown(f"**{r['descripcion']}**")

                        st.caption(f"SAP: `{cod_item}` · Equipo: **{r['equipo_principal']}** · ⏱️ {float(r['horas_ciclo_actual']):,.0f} hrs · 💵 ${c_hr_item:,.2f}/hr")

                        if st.button("🔍 Ver Hoja de Vida", key=f"sel_card_{cod_item}", use_container_width=True):

                            st.session_state.selected_component = cod_item

                            st.rerun()



    st.divider()

    opciones = [f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']})" for _, r in df_filt.iterrows()]

    cod_list = df_filt["codigo_equipo"].tolist()

    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0

    sel_comp = st.selectbox("📋 Hoja de Vida del Componente Seleccionado:", opciones, index=idx_def, key="vis_sel_comp")

    cod_sel = sel_comp.split(" | ")[0].strip()

    st.session_state.selected_component = cod_sel



    comp = df_filt[df_filt["codigo_equipo"] == cod_sel].iloc[0]

    total_vidas = max(0, int(comp["ciclo_actual"] or 1))

    if st.session_state.vida_comp_ref != cod_sel:

        st.session_state.vida_comp_ref = cod_sel

        st.session_state.vida_idx = max(1, total_vidas)

    nav_state = get_vidas_navigation_state(total_vidas, st.session_state.vida_idx)



    with st.container(border=True):

        c_info, c_vidas = st.columns([3.2, 1.2])

        with c_info:

            st.markdown(get_badges_line_html(comp["flota"], comp["estado_actual"], total_vidas), unsafe_allow_html=True)

            st.subheader(comp["descripcion"])

            st.write(f"**Código SAP:** `{comp['codigo_equipo']}` · **Equipo:** `{comp['equipo_principal']}` · **Fabricante:** {comp['fabricante']} · **Serie:** {comp['serie']} · **ABC:** {comp['criticidad_abc']}")

            st.write(f"⏱️ **Hrs Ciclo:** {float(comp['horas_ciclo_actual']):,.1f} hrs | ⏳ **Hrs Totales:** {float(comp['horas_acumuladas_totales']):,.1f} hrs | 💵 **Costo Ciclo:** ${float(comp['gasto_ciclo_actual_usd']):,.2f} | 🏦 **Costo Total:** ${float(comp['gasto_acumulado_total_usd']):,.2f}")

        with c_vidas:

            st.metric("Vidas Registradas", nav_state["label"])

            if nav_state["show_nav"]:

                b1, b2 = st.columns(2)

                with b1:

                    if st.button("◀ Ant.", disabled=nav_state["disable_prev"], use_container_width=True):

                        st.session_state.vida_idx = max(1, nav_state["current"] - 1)

                        st.rerun()

                with b2:

                    if st.button("Sig. ▶", disabled=nav_state["disable_next"], use_container_width=True):

                        st.session_state.vida_idx = min(total_vidas, nav_state["current"] + 1)

                        st.rerun()



    p1, p2 = st.columns(2)

    with p1:

        with st.container(border=True):

            st.markdown("**📷 Fotografía #1 (Previa)**")

            if comp["foto_1_path"] and os.path.exists(comp["foto_1_path"]):

                st.image(comp["foto_1_path"], use_container_width=True)

                st.caption(f"{comp['foto_1_fecha'] or ''} — {comp['foto_1_comentario'] or ''}")

            else:

                st.caption("Sin fotografía en espacio #1.")

    with p2:

        with st.container(border=True):

            st.markdown("**📷 Fotografía #2 (Más Reciente)**")

            if comp["foto_2_path"] and os.path.exists(comp["foto_2_path"]):

                st.image(comp["foto_2_path"], use_container_width=True)

                st.caption(f"{comp['foto_2_fecha'] or ''} — {comp['foto_2_comentario'] or ''}")

            else:

                st.caption("Sin fotografía en espacio #2.")



    conn = get_connection()

    df_eventos = pd.read_sql_query("SELECT * FROM historial_eventos WHERE codigo_equipo = ? ORDER BY id_evento DESC", conn, params=(cod_sel,))

    conn.close()

    if nav_state["show_nav"] and not df_eventos.empty:

        df_eventos = df_eventos[df_eventos["numero_ciclo"] == nav_state["current"]]



    if not df_eventos.empty:

        st.markdown("**📓 Historial de Eventos, OC/OT y Bitácora**")

        st.dataframe(

            df_eventos[["fecha_hora", "numero_ciclo", "tipo_operacion", "estado_asignado", "numero_ot", "numero_oc", "costo_total_evento_usd", "comentario_bitacora"]],

            use_container_width=True,

            hide_index=True,

        )





# ==============================================================================

# 6. SECCIÓN 2: ACTUALIZAR COMPONENTE (FOTOS FIFO + ESTADO/OC + CARGA HORAS)

# ==============================================================================

def render_tab_actualizar(df_all: pd.DataFrame, df_flota: pd.DataFrame) -> None:

    if df_flota.empty:

        st.warning("No hay componentes en esta flota.")

        return



    opciones = [f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} - {r['equipo_principal']}) [{r['estado_actual']}]" for _, r in df_flota.iterrows()]

    cod_list = df_flota["codigo_equipo"].tolist()

    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0

    sel_comp = st.selectbox("Seleccionar Componente a Actualizar:", opciones, index=idx_def, key="upd_sel_comp")

    cod_sel = sel_comp.split(" | ")[0].strip()

    st.session_state.selected_component = cod_sel

    comp_row = df_flota[df_flota["codigo_equipo"] == cod_sel].iloc[0]



    with st.container(border=True):

        st.markdown(get_badges_line_html(comp_row["flota"], comp_row["estado_actual"], comp_row["ciclo_actual"]), unsafe_allow_html=True)

        st.write(f"**{comp_row['codigo_equipo']} — {comp_row['descripcion']}** · Equipo: `{comp_row['equipo_principal']}` · Hrs Ciclo: `{float(comp_row['horas_ciclo_actual']):,.1f}`")



    role = st.session_state.user_role

    if role == "Campo-Tablet":

        sub_foto, = st.tabs(["📸 Subir Foto y Comentario"])

        sub_estado, sub_horas = None, None

    else:

        sub_foto, sub_estado, sub_horas = st.tabs(["📸 Foto y Comentario", "🔄 Estado, OT/OC y Costos", "⏱️ Cargar Horas (Todas las Flotas)"])



    with sub_foto:

        c_f1, c_f2 = st.columns(2)

        with c_f1:

            if comp_row["foto_1_path"] and os.path.exists(comp_row["foto_1_path"]):

                st.image(comp_row["foto_1_path"], caption=f"Foto #1: {comp_row['foto_1_comentario'] or ''}", use_container_width=True)

            else:

                st.caption("Espacio Foto #1 libre")

        with c_f2:

            if comp_row["foto_2_path"] and os.path.exists(comp_row["foto_2_path"]):

                st.image(comp_row["foto_2_path"], caption=f"Foto #2 (Reciente): {comp_row['foto_2_comentario'] or ''}", use_container_width=True)

            else:

                st.caption("Espacio Foto #2 libre")



        origen = st.radio("Capturar desde:", ["📁 Archivo / Galería", "📸 Cámara"], horizontal=True)

        img_file = st.file_uploader("Seleccionar imagen:", type=["jpg", "jpeg", "png", "webp"], key="up_img") if origen.startswith("📁") else st.camera_input("Tomar foto:")

        com_txt = st.text_area("Comentario u observación:", height=80, key="up_com")



        b1, b2 = st.columns(2)

        with b1:

            if st.button("💾 Guardar Foto (FIFO) y Comentario", type="primary", use_container_width=True):

                if img_file is None:

                    st.error("Selecciona o toma una fotografía primero.")

                else:

                    ok, msg, _ = save_component_photo_fifo(cod_sel, img_file, com_txt, role, True)

                    if ok:

                        st.success(msg)

                        st.rerun()

                    else:

                        st.error(msg)

        with b2:

            if st.button("📝 Guardar Solo Comentario", use_container_width=True):

                if not com_txt.strip():

                    st.warning("Escribe un comentario primero.")

                else:

                    conn = get_connection()

                    cur = conn.cursor()

                    cur.execute(

                        "INSERT INTO historial_eventos (codigo_equipo, numero_ciclo, fecha_hora, usuario_rol, tipo_operacion, estado_asignado, equipo_principal, horas_componente_evento, comentario_bitacora) VALUES (?, ?, ?, ?, 'Nota / Comentario', ?, ?, ?, ?)",

                        (cod_sel, comp_row["ciclo_actual"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), role, comp_row["estado_actual"], comp_row["equipo_principal"], comp_row["horas_ciclo_actual"], com_txt.strip()),

                    )

                    conn.commit()

                    conn.close()

                    st.success("Comentario registrado en bitácora.")



    if sub_estado is not None:

        with sub_estado:

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



            com_est = st.text_input("Observación del cambio de estado:")

            confirm_baja = True

            if nuevo_estado == "De Baja":

                st.warning("⚠️ Dar de baja retirará el componente de operación activa.")

                confirm_baja = st.checkbox("Confirmo dar de baja este componente.")



            if st.button("💾 Guardar Estado y Costos", type="primary", use_container_width=True, disabled=(not ok_k or not ok_r or not confirm_baja)):

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

                st.success(f"Componente `{cod_sel}` guardado en estado **{nuevo_estado}**.")

                st.rerun()



    if sub_horas is not None:

        with sub_horas:

            st.write("Sube un solo archivo Excel/CSV para actualizar las horas de **todos los componentes Montados de todas las flotas**.")

            st.download_button(

                "📥 Descargar Plantilla Excel de Horas",

                data=generate_hours_template_excel(),

                file_name="Plantilla_Carga_Horas_Montados.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            )

            archivo_hrs = st.file_uploader("Archivo Excel/CSV ([Equipo_Principal o Codigo_Componente, Horas_a_Sumar o Horometro_Actual]):", type=["xlsx", "xls", "csv"], key="up_hrs_file")

            if archivo_hrs is not None:

                df_up = pd.read_csv(archivo_hrs) if archivo_hrs.name.lower().endswith(".csv") else pd.read_excel(archivo_hrs)

                st.dataframe(df_up.head(10), use_container_width=True, hide_index=True)

                cols_norm = {c.strip().lower(): c for c in df_up.columns}

                t_col = cols_norm.get("equipo_principal") or cols_norm.get("codigo_componente") or cols_norm.get("codigo_equipo")

                t_type = "equipo_principal" if "equipo_principal" in cols_norm else "codigo_equipo"

                v_col = cols_norm.get("horas_a_sumar") or cols_norm.get("horometro_actual")

                m_op = "add" if "horas_a_sumar" in cols_norm else "set_cycle"

                if t_col and v_col:

                    if st.button("🚀 Actualizar Horas de Componentes Montados", type="primary", use_container_width=True):

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

                    st.error("El archivo debe tener las columnas 'Equipo_Principal' (o 'Codigo_Componente') y 'Horas_a_Sumar' (o 'Horometro_Actual').")



            st.divider()

            c_eq, c_hr = st.columns(2)

            with c_eq:

                eq_quick = st.selectbox("O actualizar por Equipo Principal en pantalla:", sorted(df_all[df_all["estado_actual"] == "Montado"]["equipo_principal"].dropna().unique().tolist()) or ["SIN_EQUIPO"])

            with c_hr:

                hrs_q_txt = st.text_input("Horas a sumar (+ hrs):", value="0", key="hrs_q_input")

            ok_hq, h_q_val, err_hq = validate_numeric_input(hrs_q_txt, "Horas", allow_zero=False, allow_negative=False, allow_decimal=True)

            if st.button("⚡ Aplicar Horas al Equipo Seleccionado", use_container_width=True):

                if not ok_hq:

                    st.error(err_hq)

                else:

                    n_upd, _ = apply_hours_to_mounted_components([{"target_type": "equipo_principal", "target_value": eq_quick, "mode": "add", "value": h_q_val}], role, f"Rápida ({eq_quick})")

                    st.success(f"Se sumaron +{h_q_val} hrs a {n_upd} componente(s) montados en {eq_quick}.")





# ==============================================================================

# 7. SECCIÓN 3: AGREGAR COMPONENTE Y SINCRONIZAR EXCEL SAP

# ==============================================================================

def render_tab_agregar() -> None:

    sub_alta, sub_sap = st.tabs(["➕ Agregar Componente Individual", "🔄 Importar Lista Excel SAP"])



    with sub_alta:

        with st.container(border=True):

            c1, c2 = st.columns(2)

            with c1:

                cod_m = st.text_input("Código Equipo SAP *:", placeholder="10045990")

                desc_m = st.text_input("Descripción del Componente *:", placeholder="MOTOR DE GIRO TRASERO")

                ubic_m = st.text_input("Ubicación Técnica SAP *:", value="SPCC-TOQ-MIN-495H-SHO009-GIRO")

                flota_prev, eq_prev = parse_ubicacion_tecnica(ubic_m)

                st.caption(f"Flota detectada: **{flota_prev}** · Equipo Principal: **{eq_prev}**")

                crit_m = st.selectbox("Criticidad ABC:", ["A", "B", "C"])

            with c2:

                fab_m = st.text_input("Fabricante:", placeholder="KOMATSU / CAT")

                serie_m = st.text_input("N° de Serie:", placeholder="SN-884920")

                est_m = st.selectbox("Estado Inicial:", ESTADOS_COMPONENTE)

                hrs_txt = st.text_input("Horas Iniciales:", value="0")

                costo_txt = st.text_input("Costo Inicial (USD):", value="0")

                vidas_txt = st.text_input("Vidas Iniciales (entero >= 1):", value="1")



            ok_h, h_val, err_h = validate_numeric_input(hrs_txt, "Horas Iniciales", allow_zero=True, allow_negative=False, allow_decimal=True)

            ok_c, c_val, err_c = validate_numeric_input(costo_txt, "Costo Inicial", allow_zero=True, allow_negative=False, allow_decimal=True)

            ok_v, v_val, err_v = validate_numeric_input(vidas_txt, "Vidas", allow_zero=False, allow_negative=False, allow_decimal=False)

            for ok_i, err_i in ((ok_h, err_h), (ok_c, err_c), (ok_v, err_v)):

                if not ok_i:

                    st.error(err_i)



            if st.button("💾 Guardar Nuevo Componente", type="primary", use_container_width=True):

                if not cod_m.strip() or not desc_m.strip():

                    st.error("El Código SAP y la Descripción son obligatorios.")

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

                        st.success(f"Componente `{cod_m.strip()}` registrado en Flota {flota_prev}.")

                    except sqlite3.IntegrityError:

                        st.error(f"El código `{cod_m.strip()}` ya existe.")

                    finally:

                        conn.close()



    with sub_sap:

        with st.container(border=True):

            st.write("Formato requerido (`Sheet1`): `Equipo`, `Denominación`, `Indicador ABC`, `Ubicac.técnica`, `Denominación3`, `Fabricante`, `Fabr. Nº-serie`, `Válido de`, `Fin de validez`.")

            st.download_button(

                "📥 Descargar Plantilla 'Inventario SAP Nov.xlsx'",

                data=generate_sap_template_excel(),

                file_name="Inventario SAP Nov.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            )

            sap_file = st.file_uploader("Subir Excel de SAP:", type=["xlsx", "xls"], key="sap_up_file")

            if sap_file is not None:

                xl = openpyxl.load_workbook(sap_file, read_only=True)

                sheet_to_use = "Sheet1" if "Sheet1" in xl.sheetnames else xl.sheetnames[0]

                sap_file.seek(0)

                df_sap_up = pd.read_excel(sap_file, sheet_name=sheet_to_use)

                st.dataframe(df_sap_up.head(10), use_container_width=True, hide_index=True)

                if st.button("🚀 Sincronizar Lista SAP", type="primary", use_container_width=True):

                    n_new, n_upd, msg = import_sap_dataframe(df_sap_up, st.session_state.user_role)

                    st.success(f"{msg} (Nuevos: {n_new} | Actualizados: {n_upd})")





# ==============================================================================

# 8. SECCIÓN 4: ESTRATEGIA Y BACKUP

# ==============================================================================

def render_tab_estrategia(df_all: pd.DataFrame) -> None:

    if df_all.empty:

        st.info("No hay datos suficientes para graficar.")

        return



    df_comp = df_all.copy()

    df_comp["costo_por_hora_usd"] = df_comp.apply(

        lambda r: round(r["gasto_acumulado_total_usd"] / r["horas_acumuladas_totales"], 2) if r["horas_acumuladas_totales"] > 0 else 0.0,

        axis=1,

    )



    c1, c2 = st.columns(2)

    with c1:

        with st.container(border=True):

            df_fl = df_comp.groupby("flota", as_index=False).agg(gasto=("gasto_acumulado_total_usd", "sum"), hrs=("horas_acumuladas_totales", "sum"))

            df_fl["costo_hr"] = df_fl.apply(lambda r: round(r["gasto"] / r["hrs"], 2) if r["hrs"] > 0 else 0.0, axis=1)

            fig = px.bar(df_fl, x="flota", y="costo_hr", color="flota", text="costo_hr", title="Costo por Hora ($/hr) por Flota")

            fig.update_traces(texttemplate="$%{text:.2f}/hr", textposition="outside")

            fig.update_layout(height=320, showlegend=False)

            st.plotly_chart(fig, use_container_width=True)

    with c2:

        with st.container(border=True):

            st.markdown("**🚨 Top Componentes por Costo/Hora ($/hr)**")

            st.dataframe(

                df_comp.sort_values("costo_por_hora_usd", ascending=False).head(8)[["flota", "codigo_equipo", "descripcion", "costo_por_hora_usd", "estado_actual"]],

                use_container_width=True,

                hide_index=True,

            )



    st.divider()

    st.markdown("#### 🛡️ Respaldo Completo de Base de Datos")

    ts = datetime.now().strftime("%Y%m%d_%H%M")

    b1, b2 = st.columns(2)

    with b1:

        st.download_button(

            "📥 Descargar Base Completa a Excel (.xlsx)",

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

# 9. PUNTO DE ENTRADA PRINCIPAL (FLUJO ÚNICO Y DIRECTO)

# ==============================================================================

def main() -> None:

    init_session_state()

    inject_clean_css()

    init_db()



    conn = get_connection()

    df_all = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, equipo_principal, codigo_equipo", conn)

    conn.close()



    df_flota = render_top_bar(df_all)



    if st.session_state.user_role == "Campo-Tablet":

        t_vis, t_upd = st.tabs(["👁️ 1. Visualizar Componentes", "✏️ 2. Actualizar y Subir Foto"])

        with t_vis:

            render_tab_visualizar(df_flota)

        with t_upd:

            render_tab_actualizar(df_all, df_flota)

    else:

        t_vis, t_upd, t_add, t_est = st.tabs(

            [

                "👁️ 1. Visualizar",

                "✏️ 2. Actualizar y Fotos",

                "➕ 3. Agregar y SAP",

                "📊 4. Estrategia y Backup",

            ]

        )

        with t_vis:

            render_tab_visualizar(df_flota)

        with t_upd:

            render_tab_actualizar(df_all, df_flota)

        with t_add:

            render_tab_agregar()

        with t_est:

            render_tab_estrategia(df_all)





if __name__ == "__main__":

    main()

