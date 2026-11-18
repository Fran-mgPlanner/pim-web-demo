# -*- coding: utf-8 -*-

"""

P.I.M. - CONTROL DE INVENTARIOS | PLANEAMIENTO MANTENIMIENTO MINA

Arquitectura CMMS / EAM Moderna (Inspirada en MaintainX, UpKeep, Limble CMMS y Fiori SAP PM)

Regla Técnica: 100% Componentes Nativos de Streamlit (CERO inyección de estilos externos).

Garantiza visualización limpia en cualquier versión de Streamlit Cloud, PC corporativa, tablet y móvil.



Estructura Operativa Directa (4 Pestañas):

  1. 👁️ Inventario y Reemplazos: KPIs por Flota, Filtros en 1 línea, Vista Tarjetas / Tabla con barras de progreso,

     Hoja de Vida con Reemplazos Compatibles del mismo Modelo/Flota, 2 Fotos FIFO, Resumen por Ciclo y Log Cronológico.

  2. ✏️ Editar, Foto FIFO y Horas:

     - Subpestaña 1: Acción rápida Campo/Planner (Visor 2 fotos + Cámara/Galería con rotación FIFO + Estado/Equipo + Comentario).

     - Subpestaña 2: Gestión Técnica y Financiera SAP (Ubicación Técnica SAP con deducción automática de estado,

       control de vidas con navegación, OT, OC, Proveedor, Kits y Costos USD).

     - Subpestaña 3: Corrida Masiva de Horas (Archivo Excel/CSV, tabla editable en vivo o suma rápida por equipo).

  3. ➕ Ingresar / Subir SAP: Alta guiada en cascada (Flota -> Modelo -> Equipo -> Ubicación SAP -> Estado sugerido)

     e importador directo de Excel SAP (Sheet1).

  4. 📊 Estrategia y Backup: Gráficos interactivos $/hr por Flota/Modelo, matriz Horas vs Costo,

     ranking crítico con barras de desgaste y respaldo inmediato en .xlsx (multi-pestaña) y .db (SQLite).

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

# 0. CONFIGURACIÓN INICIAL Y CATÁLOGOS DE LAS 4 FLOTAS FIJAS

# ==============================================================================

st.set_page_config(

    page_title="P.I.M. | Control de Inventarios CMMS",

    page_icon="⚙️",

    layout="wide",

    initial_sidebar_state="collapsed",

)



BASE_DIR = os.getcwd()

DB_PATH = os.path.join(BASE_DIR, "inventario_mina.db")

PHOTOS_DIR = os.path.join(BASE_DIR, "fotos_componentes")

DEFAULT_SAP_EXCEL = os.path.join(BASE_DIR, "Inventario SAP Nov.xlsx")



# LAS 4 FLOTAS ÚNICAS Y DEFINIDAS DEL SISTEMA

FLOTAS_MACRO = ["PALAS", "PERFOS", "VOLQUETES", "AUXILIARES"]



# Referencia de meta de horas de ciclo (PCR referencial) para barras visuales de progreso

PCR_META_HORAS_POR_FLOTA: Dict[str, float] = {

    "PALAS": 12000.0,

    "PERFOS": 8000.0,

    "VOLQUETES": 15000.0,

    "AUXILIARES": 10000.0,

}



CATALOGO_FLOTAS: Dict[str, Dict[str, List[str]]] = {

    "PALAS": {

        "4100A": ["PALA4 (SHO004 / PATQ4)", "PALA5 (SHO005 / PATQ5)", "PALA6 (SHO006 / PATQ6)"],

        "495HR": ["PALA7 (SHO007 / PATQ7)", "PALA8 (SHO008 / PATQ8)", "PALA9 (SHO009 / PATQ9)"],

        "4100XPC": ["PALA10 (SHO010 / PATQ10)", "PALA11 (SHO011 / PATQ11)"],

        "PC8000": ["EXC7001 (PALA HIDRÁULICA)"],

    },

    "PERFOS": {

        "49HR": ["PER001", "PER002", "PER003"],

        "PV271": ["PER004", "PER005", "PER006"],

        "PV351": ["PER007", "PER008"],

        "MD6640": ["PER009", "PER010"],

    },

    "VOLQUETES": {

        "797F": ["VOL001", "VOL002", "VOL003", "VOL004", "VOL005"],

        "793F": ["VOL010", "VOL011", "VOL012", "VOL015"],

        "930E": ["VOL020", "VOL021", "VOL022", "VOL025"],

        "980E": ["VOL030", "VOL031", "VOL032"],

    },

    "AUXILIARES": {

        "994F": ["FLO005 (CARGADOR)", "FLO006 (CARGADOR)"],

        "994K": ["FLO007 (CARGADOR)", "FLO008 (CARGADOR)"],

        "D11T": ["TRACT001", "TRACT002", "TRACT003"],

        "854K": ["WT001 (WHEELDOZER)", "WT002 (WHEELDOZER)"],

        "24M": ["MOTO001", "MOTO002"],

        "777G": ["CIST001", "CIST002"],

    },

}



ESTADOS_COMPONENTE = [

    "Montado",

    "Por Reparar",

    "Reparado",

    "Desmontado",

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

# 1. REGLAS DE CLASIFICACIÓN SAP: FLOTA (4 FIJAS), MODELO, EQUIPO Y ESTADO

# ==============================================================================

def normalize_model_name(raw_model: str) -> str:

    m = str(raw_model or "").upper().strip()

    if m in ("410A", "4100"):

        return "4100A"

    if m in ("495H", "495"):

        return "495HR"

    return m if m else "4100A"





def infer_flota_from_model_or_unit(modelo: str, equipo: str, ubic_tec: str = "") -> str:

    """Garantiza que SIEMPRE retorne una de las 4 flotas: PALAS, PERFOS, VOLQUETES o AUXILIARES."""

    m_up = normalize_model_name(modelo)

    e_up = str(equipo or "").upper().strip()

    u_up = str(ubic_tec or "").upper().strip()

    joined = f"{m_up} {e_up} {u_up}"



    for flota_name, modelos_dict in CATALOGO_FLOTAS.items():

        if m_up in modelos_dict:

            return flota_name



    if any(k in joined for k in ("PERF", "PER0", "PV271", "PV351", "49HR", "MD6")):

        return "PERFOS"

    if any(k in joined for k in ("VOL", "CAM", "TRK", "797", "793", "930E", "980E", "830E")):

        return "VOLQUETES"

    if any(k in joined for k in ("AUX", "FLO", "994", "D11", "D10", "854", "24M", "16M", "777", "CIST", "TRACT", "MOTO", "WT0")):

        return "AUXILIARES"

    return "PALAS"





def infer_estado_from_ubicacion_sap(ubic_tec: str, current_status: Optional[str] = None) -> str:

    """

    Deduce automáticamente el estado del componente a partir de su Ubicación Técnica SAP:

    - TALLER / REP / OVERHAUL / EVALUAC -> 'Por Reparar'

    - ALMACEN / BODEGA / LISTO / STOCK -> 'Reparado'

    - RECLAMO / GARANT -> 'Salida por Reclamo'

    - BAJA / SCRAP / CHATARRA -> 'De Baja'

    - DESMONT / PISO / PATIO -> 'Desmontado'

    - Ubicación de equipo (SHO, PALA, PATQ, PER, VOL, FLO, TRACT, MOTO, CIST) -> 'Montado'

    """

    u = str(ubic_tec or "").upper().strip()

    if not u:

        return current_status or "Montado"

    if any(k in u for k in ("BAJA", "SCRAP", "CHATARRA")):

        return "De Baja"

    if any(k in u for k in ("RECLAMO", "GARANT", "PROV")):

        return "Salida por Reclamo"

    if any(k in u for k in ("POR REPARAR", "TALLER", "REPARAC", "OVERHAUL", "EVALUAC")):

        return "Por Reparar"

    if any(k in u for k in ("REPARADO", "LISTO", "ALMACEN", "BODEGA", "STOCK", "STAND")):

        return "Reparado"

    if any(k in u for k in ("DESMONT", "PISO", "PATIO", "BAHIA")):

        return "Desmontado"

    if any(k in u for k in ("NUEVO",)):

        return "Nuevo"

    if any(k in u for k in ("SHO", "PALA", "PATQ", "PER", "VOL", "FLO", "TRACT", "MOTO", "CIST")):

        return "Montado"

    return current_status or "Montado"





def parse_ubicacion_tecnica(ubic_tec: str) -> Tuple[str, str, str, str]:

    """

    Extrae de la Ubicación Técnica SAP (ej. SPCC-TOQ-MIN-4100A-SHO004-PROP):

    - `flota`: PALAS | PERFOS | VOLQUETES | AUXILIARES

    - `modelo`: 4to bloque normalizado (ej. 4100A, 495HR, PV271, 797F, 994F)

    - `equipo_principal`: 5to bloque (ej. SHO004, SHO008, PER004, VOL002, FLO005)

    - `estado_sap`: estado deducido según la ubicación técnica

    """

    if not ubic_tec or pd.isna(ubic_tec):

        return ("PALAS", "4100A", "SIN_EQUIPO", "Montado")



    bloques = [b.strip() for b in str(ubic_tec).strip().split("-") if b.strip()]

    raw_mod = bloques[3].upper() if len(bloques) >= 4 else "4100A"

    modelo = normalize_model_name(raw_mod)

    equipo_principal = bloques[4].upper() if len(bloques) >= 5 else "SIN_EQUIPO"

    flota = infer_flota_from_model_or_unit(modelo, equipo_principal, str(ubic_tec))

    estado_sap = infer_estado_from_ubicacion_sap(str(ubic_tec))

    return flota, modelo, equipo_principal, estado_sap





# ==============================================================================

# 2. UTILIDADES DE VALIDACIÓN NUMÉRICA, ETIQUETAS NATIVAS Y VIDAS

# ==============================================================================

def format_vidas_label(vidas: int) -> str:

    n = int(vidas) if vidas is not None else 0

    return "1 vida" if n == 1 else f"{n} vidas"





def get_vidas_navigation_state(total_vidas: int, vida_actual: int) -> Dict:

    """

    - Con 0 o 1 vida: muestra '1 vida' (o '0 vidas') y oculta los botones de navegación.

    - Con varias vidas (>1): muestra '&lt;n&gt; vidas' en plural y activa '◀ Ant.' / 'Sig. ▶'.

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

    if raw_value is None or str(raw_value).strip() == "":

        return False, None, f"El campo '{field_name}' no puede estar vacío."

    cleaned = str(raw_value).strip().replace(",", ".")

    try:

        num = float(cleaned)

    except ValueError:

        return False, None, f"El campo '{field_name}' debe ser numérico (sin letras)."

    if not allow_negative and num < 0:

        return False, None, f"El campo '{field_name}' no permite valores negativos."

    if not allow_zero and num == 0:

        return False, None, f"El campo '{field_name}' debe ser mayor a cero."

    if not allow_decimal and not num.is_integer():

        return False, None, f"El campo '{field_name}' debe ser un número entero (sin decimales)."

    return True, num, ""





def get_native_badges_markdown(flota: str, modelo: str, estado: str, vidas: int) -> str:

    """Etiquetas semánticas 100% nativas de Streamlit (:color-background[...])."""

    flota_clean = str(flota or "PALAS").upper().strip()

    modelo_clean = str(modelo or "4100A").upper().strip()



    flota_color_map = {

        "PALAS": "red",

        "PERFOS": "blue",

        "VOLQUETES": "orange",

        "AUXILIARES": "green",

    }

    f_col = flota_color_map.get(flota_clean, "blue")



    est_color_map = {

        "Montado": "green",

        "Reparado": "blue",

        "Nuevo": "blue",

        "Por Reparar": "orange",

        "Salida por Reclamo": "orange",

        "Desmontado": "gray",

        "De Baja": "red",

    }

    e_col = est_color_map.get(estado, "gray")

    v_lbl = format_vidas_label(vidas)



    return (

        f":{f_col}-background[**🚜 {flota_clean}**] &nbsp; "

        f":violet-background[**⚙️ Mod. {modelo_clean}**] &nbsp; "

        f":{e_col}-background[**● {estado}**] &nbsp; "

        f":gray-background[**🔄 {v_lbl}**]"

    )





# ==============================================================================

# 3. BASE DE DATOS SQLITE3 CON AUTO-CORRECCIÓN DE FLOTAS Y MODELOS

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





def init_db() -> None:

    conn = get_connection()

    cur = conn.cursor()



    cur.execute(

        """

        CREATE TABLE IF NOT EXISTS componentes (

            codigo_equipo TEXT PRIMARY KEY,

            descripcion TEXT,

            flota TEXT,

            modelo TEXT,

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



    cur.execute("PRAGMA table_info(componentes)")

    existing_cols = [r["name"] for r in cur.fetchall()]

    if "modelo" not in existing_cols:

        cur.execute("ALTER TABLE componentes ADD COLUMN modelo TEXT DEFAULT '4100A'")



    # Auto-migración en cada arranque: si alguna fila tiene '410A', '495H' o '994F' en `flota`,

    # lo pasa a `modelo` y asigna la Flota real (PALAS, PERFOS, VOLQUETES, AUXILIARES).

    cur.execute("SELECT codigo_equipo, flota, modelo, equipo_principal, ubicacion_tecnica FROM componentes")

    for r in cur.fetchall():

        f_val = str(r["flota"] or "").upper().strip()

        m_val = str(r["modelo"] or "").upper().strip()

        ubic = str(r["ubicacion_tecnica"] or "")

        eq = str(r["equipo_principal"] or "")



        if f_val not in FLOTAS_MACRO or not m_val or m_val in ("SIN_MODELO", "NONE", "NAN"):

            _, mod_calc, _, _ = parse_ubicacion_tecnica(ubic)

            if f_val not in FLOTAS_MACRO and f_val not in ("", "SIN_FLOTA", "NONE"):

                mod_final = normalize_model_name(f_val)

            elif m_val and m_val not in ("SIN_MODELO", "NONE", "NAN"):

                mod_final = normalize_model_name(m_val)

            else:

                mod_final = mod_calc

            flota_final = f_val if f_val in FLOTAS_MACRO else infer_flota_from_model_or_unit(mod_final, eq, ubic)

            cur.execute(

                "UPDATE componentes SET flota = ?, modelo = ? WHERE codigo_equipo = ?",

                (flota_final, mod_final, r["codigo_equipo"]),

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

    conn.close()

    seed_initial_demo_data()





def seed_initial_demo_data() -> None:

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    demo_components = [

        ("BRASWIT231", "INTERRUPTOR / BRAKE SWITCH DE IZAJE", "PALAS", "495HR", "PALA8 (SHO008 / PATQ8)", "SPCC-TOQ-MIN-495HR-SHO008-TALLER", "A", "KOMATSU / P&H", "SN-BRK-231", "Por Reparar", 1, 7850.0, 7850.0, 9800.0, 9800.0, None, None, None, None, None, None, now_str),

        ("10045801", "MOTOR DE GIRO DELANTERO AC P&H", "PALAS", "495HR", "PALA9 (SHO009 / PATQ9)", "SPCC-TOQ-MIN-495HR-SHO009-GIRO", "A", "KOMATSU / P&H", "SN-495H-8821", "Montado", 2, 18450.0, 4250.0, 68500.0, 24500.0, None, None, None, None, None, None, now_str),

        ("10045806", "REDUCTOR DE TRASLACION (PROPEL) DER", "PALAS", "4100A", "PALA4 (SHO004 / PATQ4)", "SPCC-TOQ-MIN-4100A-SHO004-ALMACEN", "A", "BUCYRUS / CAT", "BUC-410-1120", "Reparado", 2, 19500.0, 0.0, 82000.0, 29000.0, None, None, None, None, None, None, now_str),

        ("10045807", "COMPRESOR DE AIRE PRINCIPAL", "PALAS", "4100A", "PALA4 (SHO004 / PATQ4)", "SPCC-TOQ-MIN-4100A-SHO004-NEUM", "B", "SULLAIR", "SUL-88392", "Montado", 1, 4890.0, 4890.0, 6200.0, 6200.0, None, None, None, None, None, None, now_str),

        ("PERCAB402", "CABEZAL DE ROTACION PRINCIPAL", "PERFOS", "PV271", "PER004", "SPCC-TOQ-MIN-PV271-PER004-ROTA", "A", "EPIROC / INGERSOLL", "EP-PV271-901", "Montado", 2, 12400.0, 3800.0, 54000.0, 21000.0, None, None, None, None, None, None, now_str),

        ("VOLMND797", "MANDO FINAL TRASERO DERECHO", "VOLQUETES", "797F", "VOL002", "SPCC-TOQ-MIN-797F-VOL002-TREN", "A", "CATERPILLAR", "CAT-797-4410", "Montado", 1, 9200.0, 9200.0, 115000.0, 115000.0, None, None, None, None, None, None, now_str),

        ("10045804", "CILINDRO HIDRAULICO DE LEVANTE IZQ", "AUXILIARES", "994F", "FLO005 (CARGADOR)", "SPCC-TOQ-MIN-994F-FLO005-HIDR", "A", "CATERPILLAR", "CAT-994-7741", "Montado", 2, 15900.0, 5400.0, 48000.0, 18500.0, None, None, None, None, None, None, now_str),

        ("10045805", "BOMBA HIDRAULICA DE IMPLEMENTOS", "AUXILIARES", "994F", "FLO005 (CARGADOR)", "SPCC-TOQ-MIN-994F-FLO005-TALLER", "A", "REXROTH / CAT", "REX-55821", "Por Reparar", 2, 14200.0, 7100.0, 39500.0, 14500.0, None, None, None, None, None, None, now_str),

    ]



    conn = get_connection()

    cur = conn.cursor()

    cur.executemany(

        """

        INSERT OR IGNORE INTO componentes (

            codigo_equipo, descripcion, flota, modelo, equipo_principal, ubicacion_tecnica,

            criticidad_abc, fabricante, serie, estado_actual, ciclo_actual,

            horas_acumuladas_totales, horas_ciclo_actual, gasto_acumulado_total_usd,

            gasto_ciclo_actual_usd, foto_1_path, foto_1_fecha, foto_1_comentario,

            foto_2_path, foto_2_fecha, foto_2_comentario, ultima_actualizacion_horas

        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

        """,

        demo_components,

    )



    cur.execute("SELECT COUNT(*) as c FROM historial_eventos WHERE codigo_equipo = 'BRASWIT231'")

    if cur.fetchone()["c"] == 0:

        cur.execute(

            """

            INSERT INTO historial_eventos (

                codigo_equipo, numero_ciclo, fecha_hora, usuario_rol, tipo_operacion,

                estado_asignado, equipo_principal, horas_componente_evento, horas_sumadas,

                instalado_por, motivo_cambio, modo_falla, numero_ot, numero_oc,

                proveedor_reparador, tipo_intervencion, detalle_kit, costo_kit_usd,

                costo_reparacion_usd, costo_total_evento_usd, comentario_bitacora, foto_path

            ) VALUES (

                'BRASWIT231', 1, ?, 'Planner (PALAS)', 'Desmontaje / Bitácora Planner',

                'Por Reparar', 'PALA8 (SHO008 / PATQ8)', 7850.0, 0.0,

                'Southern', 'Falla por Condición', 'Desgaste prematuro de rodamientos / bujes',

                'OT-40129901', '', 'TALLER ELECTRICO SPCC', 'Ninguna', '', 0.0, 0.0, 0.0,

                'Salió de la Pala 8 (PATQ8 / SHO008) y está por reparar en taller.', NULL

            )

            """,

            (now_str,),

        )



    conn.commit()

    conn.close()





# ==============================================================================

# 4. FUNCIONES DE NEGOCIO: CICLO FIFO 2 FOTOS, CORRIDA DE HORAS, SAP Y BACKUP

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

            msg_rotacion = "Fotografía guardada en Espacio #1 (1 de 2)."

        elif not f2_path:

            cur.execute(

                "UPDATE componentes SET foto_2_path = ?, foto_2_fecha = ?, foto_2_comentario = ? WHERE codigo_equipo = ?",

                (new_photo_path, now_display, comentario_limpio, codigo_equipo),

            )

            msg_rotacion = "Fotografía guardada en Espacio #2 (2 de 2)."

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

            msg_rotacion = "Ciclo FIFO de 2 fotos aplicado: se borró físicamente la foto #1, la #2 pasó a #1 y la nueva quedó en #2."



        if registrar_en_bitacora:

            cur.execute(

                """

                INSERT INTO historial_eventos (

                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                    tipo_operacion, estado_asignado, equipo_principal,

                    horas_componente_evento, comentario_bitacora, foto_path

                ) VALUES (?, ?, ?, ?, 'Actualización Fotográfica (FIFO)', ?, ?, ?, ?, ?)

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

        codigo_equipo = str(cod_raw).strip().upper()

        if codigo_equipo.endswith(".0"):

            codigo_equipo = codigo_equipo[:-2]



        descripcion = str(row.get("Denominación", "") or "").strip()

        denom3 = str(row.get("Denominación3", "") or "").strip()

        if denom3 and denom3 != "nan" and denom3 not in descripcion:

            descripcion = f"{descripcion} - {denom3}".strip(" -")



        ubic_tec = str(row.get("Ubicac.técnica", "") or "").strip().upper()

        flota_auto, modelo_auto, equipo_auto, estado_sap = parse_ubicacion_tecnica(ubic_tec)



        flota_col = str(row.get("Flota", "") or "").strip().upper()

        flota_final = flota_col if flota_col in FLOTAS_MACRO else flota_auto



        mod_col = str(row.get("Modelo", "") or "").strip().upper()

        modelo_final = normalize_model_name(mod_col) if (mod_col and mod_col != "NAN") else modelo_auto



        est_col = str(row.get("Estado", "") or "").strip()

        estado_final = est_col if est_col in ESTADOS_COMPONENTE else estado_sap



        criticidad_abc = str(row.get("Indicador ABC", "A") or "A").strip()

        if criticidad_abc == "nan" or not criticidad_abc:

            criticidad_abc = "A"

        fabricante = str(row.get("Fabricante", "S/I") or "S/I").strip()

        if fabricante == "nan":

            fabricante = "S/I"

        serie = str(row.get("Fabr. Nº-serie", "S/N") or "S/N").strip()

        if serie == "nan":

            serie = "S/N"



        comentario_sap = str(row.get("Comentario", "") or "").strip()

        if comentario_sap == "nan" or comentario_sap.startswith("[EJEMPLO") or comentario_sap.startswith("[Ejemplo"):

            comentario_sap = ""



        cur.execute("SELECT codigo_equipo FROM componentes WHERE UPPER(codigo_equipo) = UPPER(?)", (codigo_equipo,))

        if not cur.fetchone():

            cur.execute(

                """

                INSERT INTO componentes (

                    codigo_equipo, descripcion, flota, modelo, equipo_principal,

                    ubicacion_tecnica, criticidad_abc, fabricante, serie,

                    estado_actual, ciclo_actual, horas_acumuladas_totales,

                    horas_ciclo_actual, gasto_acumulado_total_usd,

                    gasto_ciclo_actual_usd, ultima_actualizacion_horas

                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 0.0, 0.0, 0.0, 0.0, ?)

                """,

                (codigo_equipo, descripcion, flota_final, modelo_final, equipo_auto, ubic_tec, criticidad_abc, fabricante, serie, estado_final, now_str),

            )

            cur.execute(

                """

                INSERT INTO historial_eventos (

                    codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                    tipo_operacion, estado_asignado, equipo_principal, comentario_bitacora

                ) VALUES (?, 1, ?, ?, 'Alta por Sincronización SAP', ?, ?, ?)

                """,

                (codigo_equipo, now_str, usuario_rol, estado_final, equipo_auto, comentario_sap or f"Importado de SAP en Flota {flota_final} · Modelo {modelo_final}"),

            )

            nuevos += 1

        else:

            cur.execute(

                """

                UPDATE componentes

                SET descripcion = ?, flota = ?, modelo = ?, equipo_principal = ?,

                    ubicacion_tecnica = ?, criticidad_abc = ?,

                    fabricante = ?, serie = ?, estado_actual = ?

                WHERE UPPER(codigo_equipo) = UPPER(?)

                """,

                (descripcion, flota_final, modelo_final, equipo_auto, ubic_tec, criticidad_abc, fabricante, serie, estado_final, codigo_equipo),

            )

            if comentario_sap:

                cur.execute(

                    """

                    INSERT INTO historial_eventos (

                        codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                        tipo_operacion, estado_asignado, equipo_principal, comentario_bitacora

                    ) VALUES (?, 1, ?, ?, 'Actualización SAP', ?, ?, ?)

                    """,

                    (codigo_equipo, now_str, usuario_rol, estado_final, equipo_auto, comentario_sap),

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

                """

                SELECT codigo_equipo, descripcion, flota, modelo, equipo_principal,

                       ciclo_actual, horas_ciclo_actual, horas_acumuladas_totales

                FROM componentes

                WHERE (UPPER(equipo_principal) = UPPER(?) OR UPPER(equipo_principal) LIKE UPPER(?))

                  AND estado_actual = 'Montado'

                """,

                (t_val, f"%{t_val}%"),

            )

        else:

            cur.execute(

                """

                SELECT codigo_equipo, descripcion, flota, modelo, equipo_principal,

                       ciclo_actual, horas_ciclo_actual, horas_acumuladas_totales

                FROM componentes

                WHERE UPPER(codigo_equipo) = UPPER(?) AND estado_actual = 'Montado'

                """,

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

                ) VALUES (?, ?, ?, ?, 'Corrida de Horas (Montado)', 'Montado', ?, ?, ?, ?)

                """,

                (cod, r["ciclo_actual"], now_str, usuario_rol, r["equipo_principal"], round(nuevo_ciclo, 2), round(delta, 2), f"Corrida de horas ({metodo_carga}): +{delta:,.1f} hrs (Horómetro ciclo: {nuevo_ciclo:,.1f} hrs)"),

            )

            afectados_detalle.append(

                {

                    "Flota": r["flota"],

                    "Modelo": r["modelo"],

                    "Código SAP": cod,

                    "Descripción": r["descripcion"],

                    "Equipo": r["equipo_principal"],

                    "Ciclo #": r["ciclo_actual"],

                    "Horas Sumadas (+)": round(delta, 1),

                    "Nuevas Hrs Ciclo": round(nuevo_ciclo, 1),

                    "Nuevas Hrs Totales": round(nuevo_total, 1),

                }

            )



    if afectados_detalle:

        cur.execute(

            "INSERT INTO log_carga_horas (fecha_hora, usuario, metodo_carga, componentes_afectados, detalle) VALUES (?, ?, ?, ?, ?)",

            (now_str, usuario_rol, metodo_carga, len(afectados_detalle), f"Corrida aplicada a {len(afectados_detalle)} componentes montados."),

        )



    conn.commit()

    conn.close()

    return len(afectados_detalle), afectados_detalle





def build_cycle_summary_dataframe(df_eventos_all: pd.DataFrame, comp: pd.Series) -> pd.DataFrame:

    total_ciclos = max(1, int(comp["ciclo_actual"] or 1))

    rows = []

    for c_num in range(1, total_ciclos + 1):

        sub = df_eventos_all[df_eventos_all["numero_ciclo"] == c_num] if not df_eventos_all.empty else pd.DataFrame()

        if not sub.empty:

            fecha_ini = sub["fecha_hora"].min()

            fecha_fin = sub["fecha_hora"].max()

            equipos = ", ".join(sorted(set(str(x) for x in sub["equipo_principal"].dropna() if str(x).strip())))

            hrs_max_ev = float(sub["horas_componente_evento"].max() or 0.0)

            hrs_sumadas = float(sub["horas_sumadas"].sum() or 0.0)

            corridas_count = int((sub["horas_sumadas"] > 0).sum())

            acciones_count = len(sub)

            costo_ciclo = float(sub["costo_total_evento_usd"].sum() or 0.0)

            ots = ", ".join(sorted(set(str(x) for x in sub["numero_ot"].dropna() if str(x).strip())))

            ocs = ", ".join(sorted(set(str(x) for x in sub["numero_oc"].dropna() if str(x).strip())))

            ult_com = str(sub.iloc[0]["comentario_bitacora"] or "")

        else:

            fecha_ini = comp["ultima_actualizacion_horas"] or "S/F"

            fecha_fin = comp["ultima_actualizacion_horas"] or "S/F"

            equipos = str(comp["equipo_principal"])

            hrs_max_ev = float(comp["horas_ciclo_actual"] or 0.0) if c_num == total_ciclos else 0.0

            hrs_sumadas = 0.0

            corridas_count = 0

            acciones_count = 0

            costo_ciclo = float(comp["gasto_ciclo_actual_usd"] or 0.0) if c_num == total_ciclos else 0.0

            ots, ocs, ult_com = "", "", "Ciclo activo"



        hrs_mostrar = max(hrs_max_ev, float(comp["horas_ciclo_actual"] or 0.0)) if c_num == total_ciclos else hrs_max_ev

        costo_hr_c = (costo_ciclo / hrs_mostrar) if hrs_mostrar > 0 else 0.0



        rows.append(

            {

                "Vida / Ciclo": format_vidas_label(c_num) if c_num == 1 else f"Ciclo #{c_num}",

                "Inicio Ciclo": fecha_ini,

                "Última Acción": fecha_fin,

                "Equipo(s)": equipos or str(comp["equipo_principal"]),

                "Horas Corridas": round(hrs_mostrar, 1),

                "Hrs Sumadas (Corridas)": round(hrs_sumadas, 1),

                "N° Corridas": corridas_count,

                "N° Eventos": acciones_count,

                "OT / OC": f"{ots} {ocs}".strip() or "S/N",

                "Costo Ciclo ($)": round(costo_ciclo, 2),

                "Costo/Hr ($/hr)": round(costo_hr_c, 2),

                "Último Comentario": ult_com,

            }

        )

    return pd.DataFrame(rows)





def export_full_database_to_excel() -> bytes:

    conn = get_connection()

    df_comp = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, modelo, equipo_principal, codigo_equipo", conn)

    df_hist = pd.read_sql_query("SELECT * FROM historial_eventos ORDER BY id_evento DESC", conn)

    df_logs = pd.read_sql_query("SELECT * FROM log_carga_horas ORDER BY id_carga DESC", conn)

    conn.close()



    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df_comp.to_excel(writer, sheet_name="Inventario_Actual", index=False)

        df_hist.to_excel(writer, sheet_name="Log_Acciones_Por_Ciclo", index=False)

        df_logs.to_excel(writer, sheet_name="Historial_Corridas_Horas", index=False)

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

                "Equipo": "BRASWIT231",

                "Denominación": "INTERRUPTOR / BRAKE SWITCH DE IZAJE",

                "Indicador ABC": "A",

                "Ubicac.técnica": "SPCC-TOQ-MIN-495HR-SHO008-TALLER",

                "Denominación3": "PALA 8 (SHO008 / PATQ8)",

                "Fabricante": "KOMATSU / P&H",

                "Fabr. Nº-serie": "SN-BRK-231",

                "Válido de": "2026-01-01",

                "Fin de validez": "9999-12-31",

                "Flota": "PALAS",

                "Modelo": "495HR",

                "Estado": "Por Reparar",

                "Comentario": "[Ejemplo] Salió de Pala 8",

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

            {"Equipo_Principal": "SHO009", "Horas_a_Sumar": 145.0, "Observacion_Referencia": "[Ejemplo] Corrida semanal"},

        ]

    )

    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        df_template_hrs.to_excel(writer, index=False, sheet_name="Carga_Horas")

    buf.seek(0)

    return buf.getvalue()





# ==============================================================================

# 5. BARRA DE CONTEXTO SUPERIOR (ESPACIO PROPIO POR PLANNER)

# ==============================================================================

def init_session_state() -> None:

    defaults = {

        "user_role": "Planner",

        "planner_nombre": "",

        "active_flota": "Todas",

        "selected_component": None,

        "vida_idx": 1,

        "vida_comp_ref": None,

        "vista_inventario": "Tarjetas",

    }

    for k, v in defaults.items():

        if k not in st.session_state:

            st.session_state[k] = v





def render_top_bar(df_all: pd.DataFrame) -> pd.DataFrame:

    with st.container(border=True):

        c_title, c_flota, c_rol, c_planner = st.columns([2.1, 1.4, 1.0, 1.2])

        with c_title:

            st.markdown("### ⚙️ P.I.M. — CONTROL DE INVENTARIOS")

            st.caption("PLANEAMIENTO MANTENIMIENTO MINA · PALAS · PERFOS · VOLQUETES · AUXILIARES")

        opciones_flota = ["Todas"] + FLOTAS_MACRO

        if st.session_state.active_flota not in opciones_flota:

            st.session_state.active_flota = "Todas"

        with c_flota:

            st.selectbox(

                "🚜 Flota de trabajo (4 Flotas):",

                options=opciones_flota,

                key="active_flota",

            )

        with c_rol:

            st.selectbox("👤 Rol Operativo:", options=ROLES_SISTEMA, key="user_role")

        with c_planner:

            st.text_input("✍️ Planner Responsable:", placeholder="Ej. F. Mendoza", key="planner_nombre")



    return df_all if st.session_state.active_flota == "Todas" else df_all[df_all["flota"] == st.session_state.active_flota]





# ==============================================================================

# 6. PESTAÑA 1: INVENTARIO, REEMPLAZOS POR MODELO Y LOG POR CICLO

# ==============================================================================

def render_tab_visualizar(df_all: pd.DataFrame, df_flota: pd.DataFrame) -> None:

    if df_flota.empty:

        st.info(f"No hay componentes registrados en la flota **{st.session_state.active_flota}**. Regístralos en la pestaña **➕ 3. Ingresar / Subir SAP**.")

        return



    total = len(df_flota)

    montados = int((df_flota["estado_actual"] == "Montado").sum())

    reemplazos = int(df_flota["estado_actual"].isin(["Reparado", "Nuevo"]).sum())

    taller = int(df_flota["estado_actual"].isin(["Por Reparar", "Desmontado", "Salida por Reclamo"]).sum())



    k1, k2, k3, k4 = st.columns(4)

    with k1:

        with st.container(border=True):

            st.metric("📦 Total Flota", f"{total}", f"Vista: {st.session_state.active_flota}")

    with k2:

        with st.container(border=True):

            st.metric("🟢 Montados Operando", f"{montados}", "Acumulando horas")

    with k3:

        with st.container(border=True):

            st.metric("🔵 Listos para Reemplazo", f"{reemplazos}", "Reparado / Nuevo")

    with k4:

        with st.container(border=True):

            st.metric("🟠 En Taller / Por Reparar", f"{taller}", "Pendientes de atención")



    # Filtros en una sola línea + selector de vista (Tarjetas vs Tabla CMMS)

    f1, f2, f3, f4, f5 = st.columns([1.05, 1.15, 1.15, 1.65, 0.9])

    with f1:

        modelos_disp = ["Todos"] + sorted(df_flota["modelo"].dropna().unique().tolist())

        sel_mod = st.selectbox("⚙️ Modelo Compatible:", modelos_disp, key="vis_mod")

    with f2:

        df_sub_eq = df_flota if sel_mod == "Todos" else df_flota[df_flota["modelo"] == sel_mod]

        equipos_disp = ["Todos"] + sorted(df_sub_eq["equipo_principal"].dropna().unique().tolist())

        sel_eq = st.selectbox("🚜 Equipo Principal:", equipos_disp, key="vis_eq")

    with f3:

        sel_est = st.selectbox("📌 Estado:", ["Todos", "Solo Reemplazos (Reparado/Nuevo)"] + ESTADOS_COMPONENTE, key="vis_est")

    with f4:

        q = st.text_input("🔍 Buscar Código SAP (ej. BRASWIT231), serie o descripción:", key="vis_q")

    with f5:

        st.selectbox("🗂️ Vista:", ["Tarjetas", "Tabla CMMS"], key="vista_inventario")



    df_filt = df_flota.copy()

    if sel_mod != "Todos":

        df_filt = df_filt[df_filt["modelo"] == sel_mod]

    if sel_eq != "Todos":

        df_filt = df_filt[df_filt["equipo_principal"] == sel_eq]

    if sel_est == "Solo Reemplazos (Reparado/Nuevo)":

        df_filt = df_filt[df_filt["estado_actual"].isin(["Reparado", "Nuevo"])]

    elif sel_est != "Todos":

        df_filt = df_filt[df_filt["estado_actual"] == sel_est]

    if q.strip():

        ql = q.strip().lower()

        df_filt = df_filt[

            df_filt["codigo_equipo"].astype(str).str.lower().str.contains(ql, na=False)

            | df_filt["descripcion"].astype(str).str.lower().str.contains(ql, na=False)

            | df_filt["modelo"].astype(str).str.lower().str.contains(ql, na=False)

            | df_filt["equipo_principal"].astype(str).str.lower().str.contains(ql, na=False)

            | df_filt["ubicacion_tecnica"].astype(str).str.lower().str.contains(ql, na=False)

        ]



    if df_filt.empty:

        st.warning("Sin coincidencias para este filtro.")

        return



    conn = get_connection()

    if st.session_state.vista_inventario == "Tarjetas":

        rows_list = [r for _, r in df_filt.head(15).iterrows()]

        for idx in range(0, len(rows_list), 3):

            cols = st.columns(3)

            for c_i in range(3):

                if idx + c_i < len(rows_list):

                    r = rows_list[idx + c_i]

                    cod_item = str(r["codigo_equipo"])

                    cur_ev = conn.cursor()

                    cur_ev.execute(

                        "SELECT comentario_bitacora FROM historial_eventos WHERE codigo_equipo = ? ORDER BY id_evento DESC LIMIT 1",

                        (cod_item,),

                    )

                    ult_ev = cur_ev.fetchone()

                    ult_com = ult_ev["comentario_bitacora"] if ult_ev and ult_ev["comentario_bitacora"] else "Sin comentarios registrados"



                    with cols[c_i]:

                        with st.container(border=True):

                            st.markdown(get_native_badges_markdown(r["flota"], r["modelo"], r["estado_actual"], r["ciclo_actual"]))

                            st.markdown(f"**{r['descripcion']}**")

                            st.caption(

                                f"**SAP:** `{cod_item}` · **Equipo:** `{r['equipo_principal']}`\n\n"

                                f"📍 **Ubic. SAP:** `{r['ubicacion_tecnica']}` · ⏱️ **{float(r['horas_ciclo_actual']):,.0f} hrs**\n\n"

                                f"💬 *{ult_com[:85]}*"

                            )

                            if st.button("🔍 Abrir Ficha y Compatibilidad", key=f"sel_card_{cod_item}", use_container_width=True):

                                st.session_state.selected_component = cod_item

                                st.rerun()

    else:

        df_grid = df_filt.copy()

        df_grid["vidas_lbl"] = df_grid["ciclo_actual"].apply(format_vidas_label)

        df_grid["pct_pcr"] = df_grid.apply(

            lambda r: min(100.0, round((float(r["horas_ciclo_actual"] or 0.0) / PCR_META_HORAS_POR_FLOTA.get(str(r["flota"]), 12000.0)) * 100.0, 1)),

            axis=1,

        )

        df_grid["costo_hr"] = df_grid.apply(

            lambda r: round(float(r["gasto_acumulado_total_usd"] or 0.0) / float(r["horas_acumuladas_totales"]), 2)

            if float(r["horas_acumuladas_totales"] or 0.0) > 0

            else 0.0,

            axis=1,

        )

        st.dataframe(

            df_grid[

                [

                    "codigo_equipo",

                    "descripcion",

                    "flota",

                    "modelo",

                    "equipo_principal",

                    "estado_actual",

                    "vidas_lbl",

                    "horas_ciclo_actual",

                    "pct_pcr",

                    "costo_hr",

                    "ubicacion_tecnica",

                ]

            ],

            use_container_width=True,

            hide_index=True,

            column_config={

                "codigo_equipo": "Código SAP",

                "descripcion": "Componente",

                "flota": "Flota",

                "modelo": "Modelo",

                "equipo_principal": "Equipo",

                "estado_actual": "Estado",

                "vidas_lbl": "Vidas",

                "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo", format="%.0f hrs"),

                "pct_pcr": st.column_config.ProgressColumn("% Avance Ciclo PCR", min_value=0, max_value=100, format="%.0f%%"),

                "costo_hr": st.column_config.NumberColumn("Costo/Hr", format="$%.2f/hr"),

                "ubicacion_tecnica": "Ubicación Técnica SAP",

            },

        )



    st.divider()

    opciones = [

        f"{r['codigo_equipo']} | {r['descripcion']} | Flota: {r['flota']} · Mod. {r['modelo']} · {r['equipo_principal']} [{r['estado_actual']}]"

        for _, r in df_filt.iterrows()

    ]

    cod_list = df_filt["codigo_equipo"].tolist()

    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0

    sel_comp = st.selectbox("📋 Ficha 360°, Reemplazos Compatibles y Log por Ciclo del Componente:", opciones, index=idx_def, key="vis_sel_comp")

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

            st.markdown(get_native_badges_markdown(comp["flota"], comp["modelo"], comp["estado_actual"], total_vidas))

            st.subheader(f"{comp['codigo_equipo']} — {comp['descripcion']}")

            st.write(

                f"**Flota:** `{comp['flota']}` · **Modelo Compatible:** `{comp['modelo']}` · **Equipo:** `{comp['equipo_principal']}`\n\n"

                f"📍 **Ubicación Técnica SAP:** `{comp['ubicacion_tecnica']}` · **Fabricante:** {comp['fabricante']} · **Serie:** {comp['serie']} · **ABC:** {comp['criticidad_abc']}"

            )

            st.write(

                f"⏱️ **Hrs Ciclo Actual:** {float(comp['horas_ciclo_actual']):,.1f} hrs | "

                f"⏳ **Hrs Totales:** {float(comp['horas_acumuladas_totales']):,.1f} hrs | "

                f"💵 **Costo Ciclo:** ${float(comp['gasto_ciclo_actual_usd']):,.2f} | "

                f"🏦 **Costo Acumulado:** ${float(comp['gasto_acumulado_total_usd']):,.2f}"

            )

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



    # Reemplazos compatibles del mismo Modelo y Flota

    df_reemplazos = df_all[

        (df_all["flota"] == comp["flota"])

        & (df_all["modelo"] == comp["modelo"])

        & (df_all["codigo_equipo"] != cod_sel)

    ]

    with st.expander(f"🔄 Componentes Compatibles del Modelo {comp['modelo']} en Flota {comp['flota']} para Reemplazo ({len(df_reemplazos)} encontrados)", expanded=True):

        if df_reemplazos.empty:

            st.caption(f"No hay otros componentes registrados con el modelo `{comp['modelo']}` en la flota `{comp['flota']}`.")

        else:

            st.dataframe(

                df_reemplazos[["codigo_equipo", "descripcion", "modelo", "estado_actual", "equipo_principal", "ubicacion_tecnica", "horas_ciclo_actual", "ciclo_actual"]],

                use_container_width=True,

                hide_index=True,

                column_config={

                    "codigo_equipo": "Código SAP",

                    "descripcion": "Descripción",

                    "modelo": "Modelo",

                    "estado_actual": "Estado",

                    "equipo_principal": "Equipo / Ubicación",

                    "ubicacion_tecnica": "Ubicación Técnica SAP",

                    "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo", format="%.0f hrs"),

                    "ciclo_actual": "Ciclo #",

                },

            )



    p1, p2 = st.columns(2)

    with p1:

        with st.container(border=True):

            st.markdown("**📷 Fotografía #1 (Previa en el ciclo de 2 fotos)**")

            if comp["foto_1_path"] and os.path.exists(comp["foto_1_path"]):

                st.image(comp["foto_1_path"], use_container_width=True)

                st.caption(f"📅 {comp['foto_1_fecha'] or ''} — 💬 {comp['foto_1_comentario'] or ''}")

            else:

                st.caption("Sin fotografía en espacio #1.")

    with p2:

        with st.container(border=True):

            st.markdown("**📷 Fotografía #2 (Más Reciente)**")

            if comp["foto_2_path"] and os.path.exists(comp["foto_2_path"]):

                st.image(comp["foto_2_path"], use_container_width=True)

                st.caption(f"📅 {comp['foto_2_fecha'] or ''} — 💬 {comp['foto_2_comentario'] or ''}")

            else:

                st.caption("Sin fotografía en espacio #2.")



    df_eventos_all = pd.read_sql_query("SELECT * FROM historial_eventos WHERE codigo_equipo = ? ORDER BY id_evento DESC", conn, params=(cod_sel,))

    conn.close()



    st.markdown("#### 📊 Resumen Consolidado por Ciclo / Vida")

    df_resumen_ciclos = build_cycle_summary_dataframe(df_eventos_all, comp)

    st.dataframe(

        df_resumen_ciclos,

        use_container_width=True,

        hide_index=True,

        column_config={

            "Horas Corridas": st.column_config.NumberColumn("Horas Corridas", format="%.1f hrs"),

            "Hrs Sumadas (Corridas)": st.column_config.NumberColumn("Hrs Sumadas", format="+%.1f hrs"),

            "Costo Ciclo ($)": st.column_config.NumberColumn("Costo Ciclo ($)", format="$%.2f"),

            "Costo/Hr ($/hr)": st.column_config.NumberColumn("$/hr", format="$%.2f/hr"),

        },

    )



    st.markdown("#### ⏱️ Log Cronológico de Acciones, Comentarios y Corridas por Tiempo")

    filtro_log = st.radio(

        "Alcance del log:",

        options=["Todas las vidas", f"Solo vida seleccionada (#{nav_state['current']})"] if nav_state["show_nav"] else ["Todas las vidas"],

        horizontal=True,

    )

    df_ev_mostrar = df_eventos_all[df_eventos_all["numero_ciclo"] == nav_state["current"]] if filtro_log.startswith("Solo") else df_eventos_all

    if not df_ev_mostrar.empty:

        st.dataframe(

            df_ev_mostrar[

                [

                    "fecha_hora",

                    "numero_ciclo",

                    "usuario_rol",

                    "tipo_operacion",

                    "estado_asignado",

                    "equipo_principal",

                    "horas_componente_evento",

                    "horas_sumadas",

                    "comentario_bitacora",

                    "numero_ot",

                    "numero_oc",

                    "costo_total_evento_usd",

                ]

            ],

            use_container_width=True,

            hide_index=True,

            column_config={

                "fecha_hora": "Fecha y Hora",

                "numero_ciclo": "Ciclo #",

                "usuario_rol": "Responsable",

                "tipo_operacion": "Acción",

                "estado_asignado": "Estado",

                "equipo_principal": "Equipo",

                "horas_componente_evento": st.column_config.NumberColumn("Horómetro Ciclo", format="%.1f hrs"),

                "horas_sumadas": st.column_config.NumberColumn("Hrs Sumadas", format="+%.1f hrs"),

                "comentario_bitacora": "Comentario / Detalle",

                "numero_ot": "N° OT",

                "numero_oc": "N° OC",

                "costo_total_evento_usd": st.column_config.NumberColumn("Costo Evento ($)", format="$%.2f"),

            },

        )

    else:

        st.caption("Sin eventos registrados.")





# ==============================================================================

# 7. PESTAÑA 2: EDITAR, FOTO FIFO Y CORRIDA DE HORAS

# ==============================================================================

def render_tab_actualizar(df_all: pd.DataFrame, df_flota: pd.DataFrame) -> None:

    if df_flota.empty:

        st.warning("No hay componentes en esta flota.")

        return



    opciones = [

        f"{r['codigo_equipo']} | {r['descripcion']} ({r['flota']} · Mod. {r['modelo']} · {r['equipo_principal']}) [{r['estado_actual']}]"

        for _, r in df_flota.iterrows()

    ]

    cod_list = df_flota["codigo_equipo"].tolist()

    idx_def = cod_list.index(st.session_state.selected_component) if st.session_state.selected_component in cod_list else 0

    sel_comp = st.selectbox("Seleccionar Componente a Editar / Actualizar:", opciones, index=idx_def, key="upd_sel_comp")

    cod_sel = sel_comp.split(" | ")[0].strip()

    st.session_state.selected_component = cod_sel

    comp_row = df_flota[df_flota["codigo_equipo"] == cod_sel].iloc[0]



    autor_str = f"{st.session_state.user_role}" + (f" - {st.session_state.planner_nombre.strip()}" if st.session_state.planner_nombre.strip() else f" ({comp_row['flota']})")



    with st.container(border=True):

        st.markdown(get_native_badges_markdown(comp_row["flota"], comp_row["modelo"], comp_row["estado_actual"], comp_row["ciclo_actual"]))

        st.write(

            f"**Código SAP:** `{comp_row['codigo_equipo']}` — **{comp_row['descripcion']}** · "

            f"Equipo: `{comp_row['equipo_principal']}` · Ubicación SAP: `{comp_row['ubicacion_tecnica']}` · "

            f"Hrs Ciclo: `{float(comp_row['horas_ciclo_actual']):,.1f}`"

        )



    role = st.session_state.user_role

    if role == "Campo-Tablet":

        sub_bitacora, = st.tabs(["📸 Acción Rápida: Foto (Ciclo 2 Fotos) y Comentario"])

        sub_estado, sub_horas = None, None

    else:

        sub_bitacora, sub_estado, sub_horas = st.tabs(

            [

                "📸 1. Acción Rápida: Foto FIFO y Comentario",

                "🔄 2. Gestión Técnica y Financiera SAP",

                "⏱️ 3. Corrida Masiva de Horas (Todas las Flotas)",

            ]

        )



    with sub_bitacora:

        c_f1, c_f2 = st.columns(2)

        with c_f1:

            with st.container(border=True):

                st.markdown("**📷 Foto #1 Guardada (Se borra físicamente al subir la 3ra)**")

                if comp_row["foto_1_path"] and os.path.exists(comp_row["foto_1_path"]):

                    st.image(comp_row["foto_1_path"], caption=f"{comp_row['foto_1_fecha'] or ''} — {comp_row['foto_1_comentario'] or ''}", use_container_width=True)

                else:

                    st.caption("Espacio #1 libre")

        with c_f2:

            with st.container(border=True):

                st.markdown("**📷 Foto #2 Guardada (Más Reciente)**")

                if comp_row["foto_2_path"] and os.path.exists(comp_row["foto_2_path"]):

                    st.image(comp_row["foto_2_path"], caption=f"{comp_row['foto_2_fecha'] or ''} — {comp_row['foto_2_comentario'] or ''}", use_container_width=True)

                else:

                    st.caption("Espacio #2 libre")



        with st.container(border=True):

            st.markdown("**✍️ Registro Manual del Planner / Campo (Comentario + Estado + Foto FIFO)**")

            origen = st.radio("Origen de la foto:", ["📸 Tomar Foto con Cámara", "📁 Subir Imagen desde Galería/PC"], horizontal=True)

            img_file = st.camera_input("Capturar fotografía:") if origen.startswith("📸") else st.file_uploader("Seleccionar imagen:", type=["jpg", "jpeg", "png", "webp"], key="up_img")



            c_est_rap, c_eq_rap = st.columns(2)

            with c_est_rap:

                est_rapido = st.selectbox(

                    "Estado actual del componente:",

                    options=ESTADOS_COMPONENTE,

                    index=ESTADOS_COMPONENTE.index(comp_row["estado_actual"]) if comp_row["estado_actual"] in ESTADOS_COMPONENTE else 0,

                    key="est_rap_sel",

                )

            with c_eq_rap:

                eq_rapido = st.text_input("Equipo asociado (ej. PALA8 / SHO008 / PATQ8):", value=str(comp_row["equipo_principal"]), key="eq_rap_txt")



            com_txt = st.text_area(

                "Comentario manual del Planner (ej. BRASWIT231 salió de la Pala 8 y está por reparar):",

                placeholder="Escribe aquí de qué equipo salió, condición técnica, motivo o requerimiento...",

                height=90,

                key="up_com",

            )



            b1, b2 = st.columns(2)

            with b1:

                if st.button("💾 Guardar Foto (Ciclo 2 Fotos) + Comentario", type="primary", use_container_width=True):

                    if img_file is None:

                        st.error("Toma o selecciona una foto primero (o usa el botón de la derecha para guardar solo el comentario).")

                    else:

                        conn = get_connection()

                        cur = conn.cursor()

                        cur.execute(

                            "UPDATE componentes SET estado_actual = ?, equipo_principal = ? WHERE codigo_equipo = ?",

                            (est_rapido, eq_rapido.strip(), cod_sel),

                        )

                        conn.commit()

                        conn.close()

                        ok, msg, _ = save_component_photo_fifo(cod_sel, img_file, com_txt, autor_str, True)

                        if ok:

                            st.success(msg)

                            st.rerun()

                        else:

                            st.error(msg)

            with b2:

                if st.button("📝 Guardar Comentario Manual y Estado (Sin Foto)", use_container_width=True):

                    if not com_txt.strip():

                        st.warning("Escribe un comentario manual antes de guardar.")

                    else:

                        conn = get_connection()

                        cur = conn.cursor()

                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        cur.execute(

                            "UPDATE componentes SET estado_actual = ?, equipo_principal = ? WHERE codigo_equipo = ?",

                            (est_rapido, eq_rapido.strip(), cod_sel),

                        )

                        cur.execute(

                            """

                            INSERT INTO historial_eventos (

                                codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                                tipo_operacion, estado_asignado, equipo_principal,

                                horas_componente_evento, comentario_bitacora

                            ) VALUES (?, ?, ?, ?, 'Registro Manual Planner', ?, ?, ?, ?)

                            """,

                            (cod_sel, comp_row["ciclo_actual"], now_str, autor_str, est_rapido, eq_rapido.strip(), comp_row["horas_ciclo_actual"], com_txt.strip()),

                        )

                        conn.commit()

                        conn.close()

                        st.success(f"Comentario y estado ({est_rapido}) guardados en el log del Ciclo #{int(comp_row['ciclo_actual'])}.")

                        st.rerun()



    if sub_estado is not None:

        with sub_estado:

            with st.container(border=True):

                st.markdown("**📍 Ubicación Técnica SAP, Jerarquía de Flota / Modelo, Ciclo y Costos**")

                u1, u2, u3, u4 = st.columns([1.8, 1.0, 1.0, 1.3])

                with u1:

                    nueva_ubic_sap = st.text_input(

                        "Ubicación Técnica SAP:",

                        value=str(comp_row["ubicacion_tecnica"] or ""),

                        help="Si incluye TALLER/REP sugiere 'Por Reparar'; si incluye ALMACEN/LISTO sugiere 'Reparado'; si incluye el equipo sugiere 'Montado'.",

                    )

                estado_sugerido_sap = infer_estado_from_ubicacion_sap(nueva_ubic_sap, comp_row["estado_actual"])

                with u2:

                    nueva_flota = st.selectbox(

                        "Flota (4 Fijas):",

                        options=FLOTAS_MACRO,

                        index=FLOTAS_MACRO.index(comp_row["flota"]) if comp_row["flota"] in FLOTAS_MACRO else 0,

                    )

                with u3:

                    nuevo_modelo = st.text_input("Modelo:", value=str(comp_row["modelo"] or "4100A"))

                with u4:

                    equipo_dest = st.text_input("Equipo Principal:", value=str(comp_row["equipo_principal"]))



                st.caption(f"Estado deducido automáticamente por Ubicación SAP (`{nueva_ubic_sap}`): **{estado_sugerido_sap}**")

                nuevo_estado = st.selectbox(

                    "Confirmar Estado a Asignar:",

                    options=ESTADOS_COMPONENTE,

                    index=ESTADOS_COMPONENTE.index(estado_sugerido_sap) if estado_sugerido_sap in ESTADOS_COMPONENTE else 0,

                )



                instalado_por, reiniciar_ciclo = "N/A", False

                motivo_cambio, modo_falla = "N/A", "N/A"

                numero_ot, numero_oc, proveedor, tipo_interv, detalle_kit = "", "", "", "Ninguna", ""

                costo_kit_txt, costo_rep_txt = "0", "0"



                if nuevo_estado == "Montado":

                    m1, m2 = st.columns(2)

                    with m1:

                        instalado_por = st.radio("Instalado por:", ["Southern", "Contratista"], horizontal=True)

                    with m2:

                        reiniciar_ciclo = st.checkbox(f"Iniciar nueva vida / ciclo ({int(comp_row['ciclo_actual']) + 1} vidas)", value=(comp_row["estado_actual"] != "Montado"))



                if nuevo_estado in ("Desmontado", "Por Reparar", "Salida por Reclamo", "De Baja"):

                    d1, d2 = st.columns(2)

                    with d1:

                        motivo_cambio = st.selectbox("Motivo del retiro / cambio:", ["PCR por Horas", "Falla por Condición", "Reclamo Garantía"])

                    with d2:

                        modo_falla = st.selectbox("Modo de falla:", MODOS_FALLA_COMUNES, index=0 if motivo_cambio == "PCR por Horas" else 1)



                if nuevo_estado in ("Reparado", "Por Reparar", "Salida por Reclamo") or st.checkbox("Registrar OT / OC / Costos de Reparación", value=False):

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

                    detalle_kit = st.text_input("Detalle de Kit / Alcance de Reparación:")



                ok_k, c_kit, err_k = validate_numeric_input(costo_kit_txt, "Costo Kit", allow_zero=True, allow_negative=False, allow_decimal=True)

                ok_r, c_rep, err_r = validate_numeric_input(costo_rep_txt, "Costo Reparación", allow_zero=True, allow_negative=False, allow_decimal=True)

                if not ok_k:

                    st.error(err_k)

                if not ok_r:

                    st.error(err_r)



                com_est = st.text_area("Comentario detallado para el Log del Ciclo:", height=80)

                confirm_baja = True

                if nuevo_estado == "De Baja":

                    st.warning("⚠️ Dar de baja retirará el componente de la operación activa.")

                    confirm_baja = st.checkbox("Confirmo dar de baja este componente.")



                if st.button("💾 Guardar Cambios en el Log del Ciclo", type="primary", use_container_width=True, disabled=(not ok_k or not ok_r or not confirm_baja)):

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

                        """

                        UPDATE componentes

                        SET flota = ?, modelo = ?, equipo_principal = ?, ubicacion_tecnica = ?,

                            estado_actual = ?, ciclo_actual = ?, horas_ciclo_actual = ?,

                            gasto_ciclo_actual_usd = ?, gasto_acumulado_total_usd = ?

                        WHERE codigo_equipo = ?

                        """,

                        (nueva_flota, normalize_model_name(nuevo_modelo), equipo_dest.strip().upper(), nueva_ubic_sap.strip().upper(), nuevo_estado, n_ciclo, round(n_hrs, 2), round(n_gasto_c, 2), round(n_gasto_t, 2), cod_sel),

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

                        (

                            cod_sel, n_ciclo, now_str, autor_str, f"Actualización ({nuevo_estado})",

                            nuevo_estado, equipo_dest.strip().upper(), round(n_hrs, 2), instalado_por,

                            motivo_cambio, modo_falla, numero_ot.strip(), numero_oc.strip(),

                            proveedor.strip(), tipo_interv, detalle_kit.strip(),

                            round(float(c_kit or 0.0), 2), round(float(c_rep or 0.0), 2), costo_tot,

                            (com_est or f"Ubicación SAP: {nueva_ubic_sap.strip().upper()} -> Estado: {nuevo_estado}").strip(),

                        ),

                    )

                    conn.commit()

                    conn.close()

                    st.success(f"Componente `{cod_sel}` actualizado en Flota **{nueva_flota}** (Modelo **{normalize_model_name(nuevo_modelo)}** · Ciclo #{n_ciclo}).")

                    st.rerun()



    if sub_horas is not None:

        with sub_horas:

            st.write("Sube un archivo Excel/CSV o edita en tabla para actualizar simultáneamente a **todos los componentes Montados de todas las flotas**.")

            st.download_button(

                "📥 Descargar Plantilla Excel de Corrida de Horas",

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

                    if st.button("🚀 Ejecutar Corrida de Horas en Montados", type="primary", use_container_width=True):

                        upd_list = []

                        for _, r in df_up.iterrows():

                            ok_n, num_h, _ = validate_numeric_input("" if pd.isna(r[v_col]) else str(r[v_col]), allow_zero=False, allow_negative=False, allow_decimal=True)

                            if pd.notna(r[t_col]) and ok_n:

                                upd_list.append({"target_type": t_type, "target_value": str(r[t_col]).strip(), "mode": m_op, "value": num_h})

                        n_upd, det = apply_hours_to_mounted_components(upd_list, autor_str, f"Archivo ({archivo_hrs.name})")

                        st.success(f"Corrida registrada: se actualizaron {n_upd} componente(s) montado(s).")

                        if det:

                            st.dataframe(pd.DataFrame(det), use_container_width=True, hide_index=True)

                else:

                    st.error("El archivo debe tener las columnas 'Equipo_Principal' (o 'Codigo_Componente') y 'Horas_a_Sumar' (o 'Horometro_Actual').")



            st.divider()

            st.markdown("**⚡ Asignación Rápida en Pantalla para Componentes Montados**")

            df_montados_act = df_all[df_all["estado_actual"] == "Montado"][

                ["codigo_equipo", "descripcion", "flota", "modelo", "equipo_principal", "ciclo_actual", "horas_ciclo_actual", "horas_acumuladas_totales"]

            ].copy()

            if not df_montados_act.empty:

                c_eq, c_hr = st.columns(2)

                with c_eq:

                    eq_quick = st.selectbox("Filtrar por Equipo Principal Montado:", ["Todos"] + sorted(df_montados_act["equipo_principal"].dropna().unique().tolist()))

                with c_hr:

                    hrs_q_txt = st.text_input("Pre-llenar horas a sumar (+ hrs):", value="0", key="hrs_q_input")

                ok_hq, h_q_val, err_hq = validate_numeric_input(hrs_q_txt, "Horas", allow_zero=True, allow_negative=False, allow_decimal=True)

                if not ok_hq:

                    st.error(err_hq)

                    h_q_val = 0.0



                df_ed_src = df_montados_act if eq_quick == "Todos" else df_montados_act[df_montados_act["equipo_principal"] == eq_quick]

                df_ed_src = df_ed_src.copy()

                df_ed_src["Horas_a_Sumar"] = float(h_q_val or 0.0)



                edited_montados = st.data_editor(

                    df_ed_src,

                    use_container_width=True,

                    hide_index=True,

                    disabled=["codigo_equipo", "descripcion", "flota", "modelo", "equipo_principal", "ciclo_actual", "horas_ciclo_actual", "horas_acumuladas_totales"],

                    column_config={

                        "codigo_equipo": "Código SAP",

                        "descripcion": "Componente",

                        "flota": "Flota",

                        "modelo": "Modelo",

                        "equipo_principal": "Equipo",

                        "ciclo_actual": "Ciclo #",

                        "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo Actual", format="%.1f"),

                        "horas_acumuladas_totales": st.column_config.NumberColumn("Hrs Totales", format="%.1f"),

                        "Horas_a_Sumar": st.column_config.NumberColumn("➕ Horas a Sumar (+)", min_value=0.0, max_value=10000.0, step=5.0, format="%.1f hrs"),

                    },

                    key="editor_montados_cmms",

                )

                if st.button("⚡ Aplicar Corrida de Horas a Montados", type="primary", use_container_width=True):

                    upd_grid = []

                    for _, r_ed in edited_montados.iterrows():

                        h_add = float(r_ed.get("Horas_a_Sumar", 0.0) or 0.0)

                        if h_add > 0:

                            upd_grid.append({"target_type": "codigo_equipo", "target_value": str(r_ed["codigo_equipo"]), "mode": "add", "value": h_add})

                    if not upd_grid:

                        st.warning("Ingresa un valor mayor a 0 en la columna '➕ Horas a Sumar (+)' en al menos un componente.")

                    else:

                        n_upd, det = apply_hours_to_mounted_components(upd_grid, autor_str, f"Tabla Rápida ({eq_quick})")

                        st.success(f"Se aplicaron las horas a {n_upd} componente(s) montados.")

                        if det:

                            st.dataframe(pd.DataFrame(det), use_container_width=True, hide_index=True)



            st.divider()

            st.markdown("**📜 Historial General de Corridas de Horas por Tiempo**")

            conn = get_connection()

            df_log_h = pd.read_sql_query("SELECT * FROM log_carga_horas ORDER BY id_carga DESC LIMIT 50", conn)

            conn.close()

            if not df_log_h.empty:

                st.dataframe(df_log_h, use_container_width=True, hide_index=True)





# ==============================================================================

# 8. PESTAÑA 3: INGRESAR COMPONENTE EN CASCADA O SUBIR EXCEL SAP

# ==============================================================================

def render_tab_agregar() -> None:

    sub_alta, sub_sap = st.tabs(["➕ Alta Guiada de Componente", "🔄 Subir / Sincronizar Excel SAP (Sheet1)"])



    with sub_alta:

        with st.container(border=True):

            st.markdown("**1. Selección en Cascada: Flota (4 Fijas) → Modelo Compatible → Equipo Principal**")

            j1, j2, j3 = st.columns(3)

            with j1:

                flota_sel = st.selectbox("🚜 Flota *:", options=FLOTAS_MACRO, index=0)

            with j2:

                modelos_sugeridos = list(CATALOGO_FLOTAS.get(flota_sel, {}).keys()) + ["Otro (escribir manual)"]

                mod_pick = st.selectbox("⚙️ Modelo Compatible *:", options=modelos_sugeridos)

                if mod_pick == "Otro (escribir manual)":

                    modelo_final = normalize_model_name(st.text_input("Escribir Modelo:", placeholder="Ej. 4100A, PV271, 797F"))

                else:

                    modelo_final = mod_pick

            with j3:

                equipos_sugeridos = CATALOGO_FLOTAS.get(flota_sel, {}).get(mod_pick, []) + ["TALLER / ALMACEN", "Otro (escribir manual)"]

                eq_pick = st.selectbox("🏷️ Equipo (Pala/Perfo/Volq) *:", options=equipos_sugeridos)

                if eq_pick == "Otro (escribir manual)":

                    equipo_final = st.text_input("Escribir Equipo:", placeholder="Ej. PALA4 (SHO004 / PATQ4)").strip().upper()

                else:

                    equipo_final = eq_pick



            st.divider()

            st.markdown("**2. Código SAP, Ubicación Técnica Autocompletada y Comentario del Planner**")

            c1, c2 = st.columns(2)

            with c1:

                cod_m = st.text_input("Código SAP del Componente *:", placeholder="Ej. BRASWIT231 o 10045990")

                desc_m = st.text_input("Descripción del Componente *:", placeholder="Ej. INTERRUPTOR / MOTOR DE GIRO")

                cod_corto_eq = equipo_final.split("(")[0].strip().replace(" ", "") if equipo_final else "EQ01"

                ubic_default = f"SPCC-TOQ-MIN-{modelo_final or '4100A'}-{cod_corto_eq}-SIST"

                ubic_m = st.text_input("Ubicación Técnica SAP *:", value=ubic_default)

                est_sugerido = infer_estado_from_ubicacion_sap(ubic_m)

                crit_m = st.selectbox("Criticidad ABC:", ["A", "B", "C"])

            with c2:

                est_m = st.selectbox(

                    "Estado Inicial (sugerido por Ubicación SAP):",

                    options=ESTADOS_COMPONENTE,

                    index=ESTADOS_COMPONENTE.index(est_sugerido) if est_sugerido in ESTADOS_COMPONENTE else 0,

                )

                fab_m = st.text_input("Fabricante:", placeholder="KOMATSU / CAT / EPIROC")

                serie_m = st.text_input("N° de Serie:", placeholder="SN-884920")

                h1, h2, h3 = st.columns(3)

                with h1:

                    hrs_txt = st.text_input("Horas Iniciales:", value="0")

                with h2:

                    costo_txt = st.text_input("Costo Inicial ($):", value="0")

                with h3:

                    vidas_txt = st.text_input("Vidas (>=1):", value="1")



            comentario_inicial = st.text_area(

                "Comentario inicial del Planner (ej. salió de la Pala 8 y está por reparar):",

                placeholder="Ej. Salió de la Pala 8 (SHO008 / PATQ8) y queda por reparar...",

                height=75,

            )



            ok_h, h_val, err_h = validate_numeric_input(hrs_txt, "Horas Iniciales", allow_zero=True, allow_negative=False, allow_decimal=True)

            ok_c, c_val, err_c = validate_numeric_input(costo_txt, "Costo Inicial", allow_zero=True, allow_negative=False, allow_decimal=True)

            ok_v, v_val, err_v = validate_numeric_input(vidas_txt, "Vidas", allow_zero=False, allow_negative=False, allow_decimal=False)

            for ok_i, err_i in ((ok_h, err_h), (ok_c, err_c), (ok_v, err_v)):

                if not ok_i:

                    st.error(err_i)



            if st.button("💾 Ingresar Componente", type="primary", use_container_width=True):

                if not cod_m.strip() or not desc_m.strip() or not modelo_final:

                    st.error("El Código SAP, la Descripción y el Modelo son obligatorios.")

                elif ok_h and ok_c and ok_v:

                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    autor_alta = f"{st.session_state.user_role}" + (f" - {st.session_state.planner_nombre.strip()}" if st.session_state.planner_nombre.strip() else f" ({flota_sel})")

                    conn = get_connection()

                    cur = conn.cursor()

                    try:

                        cur.execute(

                            """

                            INSERT INTO componentes (

                                codigo_equipo, descripcion, flota, modelo, equipo_principal,

                                ubicacion_tecnica, criticidad_abc, fabricante, serie,

                                estado_actual, ciclo_actual, horas_acumuladas_totales,

                                horas_ciclo_actual, gasto_acumulado_total_usd,

                                gasto_ciclo_actual_usd, ultima_actualizacion_horas

                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

                            """,

                            (

                                cod_m.strip().upper(), desc_m.strip().upper(), flota_sel, modelo_final,

                                equipo_final, ubic_m.strip().upper(), crit_m,

                                (fab_m.strip() or "S/I").upper(), (serie_m.strip() or "S/N").upper(),

                                est_m, int(v_val), round(h_val, 2), round(h_val, 2),

                                round(c_val, 2), round(c_val, 2), now_str,

                            ),

                        )

                        cur.execute(

                            """

                            INSERT INTO historial_eventos (

                                codigo_equipo, numero_ciclo, fecha_hora, usuario_rol,

                                tipo_operacion, estado_asignado, equipo_principal,

                                horas_componente_evento, costo_total_evento_usd, comentario_bitacora

                            ) VALUES (?, ?, ?, ?, 'Alta de Componente', ?, ?, ?, ?, ?)

                            """,

                            (

                                cod_m.strip().upper(), int(v_val), now_str, autor_alta, est_m,

                                equipo_final, round(h_val, 2), round(c_val, 2),

                                (comentario_inicial.strip() or f"Alta en Flota {flota_sel} · Modelo {modelo_final} · Ubicación SAP {ubic_m.strip().upper()}"),

                            ),

                        )

                        conn.commit()

                        st.success(f"Componente `{cod_m.strip().upper()}` registrado en Flota **{flota_sel}** (Modelo **{modelo_final}** · Estado **{est_m}**).")

                    except sqlite3.IntegrityError:

                        st.error(f"El código SAP `{cod_m.strip().upper()}` ya existe en la base de datos.")

                    finally:

                        conn.close()



    with sub_sap:

        with st.container(border=True):

            st.write(

                "Sube tu archivo Excel de SAP (`Sheet1`). El sistema extrae automáticamente el **Modelo** (4to bloque de `Ubicac.técnica`), "

                "el **Equipo Principal** (5to bloque), lo asigna a su **Flota** (`PALAS`, `PERFOS`, `VOLQUETES`, `AUXILIARES`) y deduce el **Estado**."

            )

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

# 9. PESTAÑA 4: ESTRATEGIA CMMS, BARRAS DE DESGASTE Y BACKUP (.XLSX / .DB)

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

    df_comp["pct_desgaste_ciclo"] = df_comp.apply(

        lambda r: min(100.0, round((float(r["horas_ciclo_actual"] or 0.0) / PCR_META_HORAS_POR_FLOTA.get(str(r["flota"]), 12000.0)) * 100.0, 1)),

        axis=1,

    )



    c1, c2 = st.columns(2)

    with c1:

        with st.container(border=True):

            df_fl = df_comp.groupby(["flota", "modelo"], as_index=False).agg(

                gasto=("gasto_acumulado_total_usd", "sum"),

                hrs=("horas_acumuladas_totales", "sum"),

                equipos=("codigo_equipo", "count"),

            )

            df_fl["costo_hr"] = df_fl.apply(lambda r: round(r["gasto"] / r["hrs"], 2) if r["hrs"] > 0 else 0.0, axis=1)

            fig = px.bar(

                df_fl,

                x="modelo",

                y="costo_hr",

                color="flota",

                text="costo_hr",

                title="💵 Costo por Hora ($/hr) por Flota y Modelo",

            )

            fig.update_traces(texttemplate="$%{text:.2f}/hr", textposition="outside")

            fig.update_layout(height=340)

            st.plotly_chart(fig, use_container_width=True)

    with c2:

        with st.container(border=True):

            fig_sc = px.scatter(

                df_comp,

                x="horas_acumuladas_totales",

                y="gasto_acumulado_total_usd",

                color="flota",

                size="ciclo_actual",

                hover_name="descripcion",

                hover_data=["codigo_equipo", "modelo", "equipo_principal", "estado_actual"],

                title="🎯 Horas Totales vs. Gasto Acumulado ($ USD)",

            )

            fig_sc.update_layout(height=340)

            st.plotly_chart(fig_sc, use_container_width=True)



    with st.container(border=True):

        st.markdown("**🚨 Componentes Críticos por Costo/Hora ($/hr) y % de Avance de Ciclo PCR**")

        st.dataframe(

            df_comp.sort_values("costo_por_hora_usd", ascending=False)[

                ["flota", "modelo", "codigo_equipo", "descripcion", "equipo_principal", "estado_actual", "ciclo_actual", "horas_ciclo_actual", "pct_desgaste_ciclo", "costo_por_hora_usd"]

            ],

            use_container_width=True,

            hide_index=True,

            column_config={

                "flota": "Flota",

                "modelo": "Modelo",

                "codigo_equipo": "Código SAP",

                "descripcion": "Componente",

                "equipo_principal": "Equipo",

                "estado_actual": "Estado",

                "ciclo_actual": "Vidas",

                "horas_ciclo_actual": st.column_config.NumberColumn("Hrs Ciclo", format="%.0f hrs"),

                "pct_desgaste_ciclo": st.column_config.ProgressColumn("% Desgaste Ciclo PCR", min_value=0, max_value=100, format="%.0f%%"),

                "costo_por_hora_usd": st.column_config.NumberColumn("Costo/Hr ($/hr)", format="$%.2f/hr"),

            },

        )



    st.divider()

    st.markdown("#### 🛡️ Respaldo Completo de Base de Datos (Inventario + Log por Ciclo + Corridas de Horas)")

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

# 10. PUNTO DE ENTRADA PRINCIPAL

# ==============================================================================

def main() -> None:

    init_session_state()

    init_db()



    conn = get_connection()

    df_all = pd.read_sql_query("SELECT * FROM componentes ORDER BY flota, modelo, equipo_principal, codigo_equipo", conn)

    conn.close()



    df_flota = render_top_bar(df_all)



    if st.session_state.user_role == "Campo-Tablet":

        t_vis, t_upd = st.tabs(["👁️ 1. Inventario y Reemplazos", "✏️ 2. Editar, Foto FIFO y Comentario"])

        with t_vis:

            render_tab_visualizar(df_all, df_flota)

        with t_upd:

            render_tab_actualizar(df_all, df_flota)

    else:

        t_vis, t_upd, t_add, t_est = st.tabs(

            [

                "👁️ 1. Inventario y Reemplazos",

                "✏️ 2. Editar, Foto FIFO y Horas",

                "➕ 3. Ingresar / Subir SAP",

                "📊 4. Estrategia y Backup",

            ]

        )

        with t_vis:

            render_tab_visualizar(df_all, df_flota)

        with t_upd:

            render_tab_actualizar(df_all, df_flota)

        with t_add:

            render_tab_agregar()

        with t_est:

            render_tab_estrategia(df_all)





if __name__ == "__main__":

    main()

