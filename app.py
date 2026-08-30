from __future__ import annotations

from collections import Counter, defaultdict
from contextlib import redirect_stdout
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo
import html
import io
import math
import re
import shutil
import tempfile
import urllib.error
import urllib.request
import zipfile

import openpyxl
import streamlit as st

import modelo_abasto as engine


APP_NAME = "Transfer Planner"
MAX_UPLOAD_MB = 500
DATA_TRANSFERS_SPREADSHEET_ID = "18kHevkMvf9l4s6ANg3h5KdNyj2yEPGAp5C_t8JwxFVw"

ORIGIN_WAREHOUSES = {
    444: "CITYPARK TURBO",
    831: "CITYPARK CHEDRAUI",
    811: "CEDA TURBO",
    834: "CEDA CHEDRAUI",
    425: "CEDIS LOCAL GDL",
    856: "CEDIS - LOCAL MTY",
    49: "NODO ALTAVISTA",
}

ALEPH_SHEETS = {
    "NO_DISPONIBLE",
    "STOCK",
    "INSUMOS",
    "GOLDEN_INFALTABLES",
    "TIENDA",
    "STORAGE",
}

MANUAL_BACKEND_SHEETS = {"TIENDAS_CERRADAS"}

REQUIRED_DATABASE_SHEETS = (
    "TIENDAS_CERRADAS",
    "VOLUMETRIA",
    "BLOQUEOS",
    "RUTA_COSTOS",
    "PRIORIDAD",
    "HIGH_VALUE",
    "RACKEADOS",
    "CAP_RECIBO",
    "COPERNICO",
    "NO_DISPONIBLE",
    "STOCK",
    "INSUMOS",
    "GOLDEN_INFALTABLES",
    "TIENDA",
    "STORAGE",
)

SHEET_DESCRIPTIONS = {
    "TIENDAS_CERRADAS": (
        "Lista permanente de warehouses destino bloqueados. Los requerimientos de "
        "estas tiendas se eliminan automáticamente antes de asignar stock y no "
        "pueden reactivarse desde la interfaz."
    ),
    "VOLUMETRIA": (
        "Volumen en metros cúbicos por unidad de cada SKU. Se utiliza para "
        "calcular el consumo de capacidad de recibo de las tiendas."
    ),
    "BLOQUEOS": (
        "Productos con restricciones regionales de envío, especialmente desde "
        "orígenes de CDMX hacia Guadalajara o Monterrey."
    ),
    "RUTA_COSTOS": (
        "Combinaciones de tienda destino y producto que no cuentan con una ruta "
        "de costos habilitada para la transferencia."
    ),
    "PRIORIDAD": (
        "Prioridad de atención por warehouse destino. El valor 1 representa la "
        "prioridad más alta."
    ),
    "HIGH_VALUE": (
        "Clasificación referencial de productos de alto valor utilizada en los "
        "archivos de salida."
    ),
    "RACKEADOS": (
        "Productos rackeados por warehouse. Para el origen 444, estos productos "
        "se excluyen completamente del stock utilizable."
    ),
    "CAP_RECIBO": (
        "Capacidad máxima de recibo de cada tienda expresada en metros cúbicos."
    ),
    "COPERNICO": (
        "Inventario detallado por ubicación del warehouse 444. Permite identificar "
        "y descontar el stock no utilizable."
    ),
    "NO_DISPONIBLE": (
        "Stock no disponible por warehouse y producto que debe descontarse del "
        "inventario disponible final."
    ),
    "STOCK": (
        "Stock disponible final por warehouse y producto. Es la fuente mandante "
        "para determinar cuánto puede enviarse."
    ),
    "INSUMOS": (
        "Cantidades de insumos calculadas automáticamente por Aleph para cada "
        "tienda. Solo se anexan al BulkCD_444 cuando la tienda ya recibe producto "
        "normal desde ese mismo origen."
    ),
    "GOLDEN_INFALTABLES": (
        "Productos Golden e Infaltables definidos por producto y ciudad, con "
        "prioridad y excepciones especiales de negocio."
    ),
    "TIENDA": (
        "Catálogo de warehouses con el nombre y la ciudad de cada tienda o nodo."
    ),
    "STORAGE": (
        "Condición de almacenamiento de cada producto, como room temperature, "
        "refrigerated o freezer."
    ),
}

DEMAND_RULE_LABELS = {
    "MOV_MINIMO_3": "ROQ POSITIVO",
    "HARDCODE_4_CERO_TOTAL": "FORECAST 0 · FORZADO A 4",
    "HARDCODE_3_INVENTARIO_MENOR_DEMANDA": "ROQ 0 · INVENTARIO MENOR A DEMANDA",
    "SIN_DEMANDA": "SIN RECOMENDACIÓN",
}

INSUMOS_COLUMNS = [
    "WAREHOUSE_DESTINATION",
    "WAREHOUSE_SOURCE",
    "RETAIL_ID",
    "QUANTITY",
    "PLANNED_DATE",
    "ROUTE",
    "DELIVERY_PRIORITY",
]

CITY_DISPLAY_NAMES = {
    "CDMX": "Ciudad de México",
    "GDL": "Guadalajara",
    "MTY": "Monterrey",
}

ENGLISH_MONTHS = {
    "JANUARY": 1,
    "FEBRUARY": 2,
    "MARCH": 3,
    "APRIL": 4,
    "MAY": 5,
    "JUNE": 6,
    "JULY": 7,
    "AUGUST": 8,
    "SEPTEMBER": 9,
    "OCTOBER": 10,
    "NOVEMBER": 11,
    "DECEMBER": 12,
}

TIMEZONE_OFFSETS = {
    "UTC": 0,
    "GMT": 0,
    "EDT": -4,
    "EST": -5,
    "CDT": -5,
    "CST": -6,
    "MDT": -6,
    "MST": -7,
    "PDT": -7,
    "PST": -8,
}


st.set_page_config(
    page_title=APP_NAME,
    page_icon="↗",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Archivo+Black&family=IBM+Plex+Mono:wght@400;500;600;700&display=swap');

        :root {
            --ink: #111111;
            --paper: #f2efe6;
            --acid: #d9ff3f;
            --coral: #ff5a47;
            --blue: #5e7cff;
            --white: #fffdf7;
            --muted: #6f6b63;
        }

        html, body, [class*="css"] {
            font-family: "IBM Plex Mono", monospace;
        }

        .stApp {
            background:
                linear-gradient(rgba(17,17,17,.055) 1px, transparent 1px),
                linear-gradient(90deg, rgba(17,17,17,.055) 1px, transparent 1px),
                var(--paper);
            background-size: 28px 28px;
            color: var(--ink);
        }

        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stToolbar"], #MainMenu, footer { visibility: hidden; }
        .block-container { max-width: 1440px; padding: 2rem 3rem 4rem; }

        .hero {
            border: 4px solid var(--ink);
            background: var(--coral);
            box-shadow: 10px 10px 0 var(--ink);
            padding: 28px 32px;
            margin: 6px 10px 34px 0;
            position: relative;
            overflow: hidden;
        }

        .hero::after {
            content: "RUN / ALLOCATE / EXPORT";
            position: absolute;
            right: -42px;
            top: 34px;
            transform: rotate(8deg);
            border: 3px solid var(--ink);
            background: var(--acid);
            padding: 8px 46px;
            font-weight: 800;
            letter-spacing: .08em;
        }

        .hero-kicker {
            display: inline-block;
            border: 3px solid var(--ink);
            background: var(--white);
            padding: 5px 9px;
            font-weight: 800;
            margin-bottom: 18px;
        }

        .hero h1 {
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(2.6rem, 7vw, 6.5rem);
            line-height: .88;
            letter-spacing: -.06em;
            margin: 0;
            max-width: 1050px;
            color: var(--ink);
        }

        .hero p {
            max-width: 780px;
            font-weight: 700;
            font-size: 1rem;
            margin: 22px 0 0;
        }

        .section-label {
            display: inline-block;
            background: var(--ink);
            color: var(--white);
            border: 3px solid var(--ink);
            padding: 7px 12px;
            margin: 18px 0 10px;
            font-weight: 800;
            letter-spacing: .08em;
        }

        .info-strip {
            border: 3px solid var(--ink);
            background: var(--acid);
            box-shadow: 6px 6px 0 var(--ink);
            padding: 14px 16px;
            margin: 8px 7px 24px 0;
            font-weight: 700;
        }

        .database-status {
            border: 4px solid var(--ink);
            box-shadow: 8px 8px 0 var(--ink);
            padding: 20px 22px;
            margin: 8px 9px 22px 0;
        }

        .database-status.online { background: var(--acid); }
        .database-status.review { background: var(--coral); }

        .database-title {
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(1.4rem, 3vw, 2.4rem);
            line-height: 1;
        }

        .source-card {
            height: 138px;
            border: 3px solid var(--ink);
            box-shadow: 5px 5px 0 var(--ink);
            padding: 14px 15px;
            margin: 0 5px 18px 0;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            position: relative;
            overflow: hidden;
        }

        .source-card.ok { background: #baf264; }
        .source-card.error { background: var(--coral); }

        .source-card-name {
            font-family: "Archivo Black", sans-serif;
            font-size: .95rem;
            line-height: 1.05;
            overflow-wrap: anywhere;
        }

        .source-card-meta {
            margin-top: 9px;
            color: #242424;
            font-size: .68rem;
            font-weight: 800;
        }

        .source-card-result {
            margin-top: 13px;
            border-top: 2px solid var(--ink);
            padding-top: 9px;
            font-size: .78rem;
            font-weight: 900;
            text-transform: uppercase;
        }

        .source-card-tooltip {
            position: absolute;
            inset: 0;
            z-index: 5;
            box-sizing: border-box;
            background: var(--ink);
            color: var(--white);
            padding: 12px 13px;
            opacity: 0;
            visibility: hidden;
            overflow-y: auto;
            font-size: .68rem;
            font-weight: 700;
            line-height: 1.35;
            text-transform: none;
            transition: opacity .12s ease, visibility 0s linear .12s;
        }

        .source-card-tooltip strong {
            display: block;
            color: var(--acid);
            margin-bottom: 6px;
            font-size: .67rem;
            letter-spacing: .05em;
        }

        .source-card:hover .source-card-tooltip {
            opacity: 1;
            visibility: visible;
            transition: opacity .12s ease 1s, visibility 0s linear 1s;
        }

        div[data-testid="stFileUploader"] {
            border: 4px dashed var(--ink);
            background: var(--white);
            box-shadow: 8px 8px 0 var(--ink);
            padding: 18px;
            margin: 8px 9px 24px 0;
        }

        div[data-testid="stFileUploaderDropzone"] {
            background: var(--white);
            border: 0;
            min-height: 190px;
        }

        div[data-testid="stForm"], div[data-testid="stExpander"] {
            border: 3px solid var(--ink);
            border-radius: 0;
            background: rgba(255,253,247,.92);
            box-shadow: 6px 6px 0 var(--ink);
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input,
        div[data-testid="stDateInput"] input,
        div[data-baseweb="select"] > div {
            border: 2px solid var(--ink) !important;
            border-radius: 0 !important;
            background: var(--white) !important;
        }

        .stButton > button,
        .stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            border: 3px solid var(--ink);
            border-radius: 0;
            background: var(--acid);
            color: var(--ink);
            box-shadow: 5px 5px 0 var(--ink);
            font-family: "IBM Plex Mono", monospace;
            font-weight: 900;
            text-transform: uppercase;
            min-height: 50px;
            transition: transform .08s ease, box-shadow .08s ease;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {
            color: var(--ink);
            border-color: var(--ink);
            transform: translate(3px, 3px);
            box-shadow: 2px 2px 0 var(--ink);
        }

        div[data-testid="stMetric"] {
            border: 3px solid var(--ink);
            background: var(--white);
            box-shadow: 5px 5px 0 var(--ink);
            padding: 14px 16px;
        }

        div[data-testid="stMetricLabel"] { font-weight: 800; }
        div[data-testid="stMetricValue"] { font-family: "Archivo Black", sans-serif; }

        .kpi-card {
            min-height: 132px;
            border: 3px solid var(--ink);
            background: var(--white);
            box-shadow: 5px 5px 0 var(--ink);
            padding: 13px 15px;
            margin: 0 5px 18px 0;
            box-sizing: border-box;
            position: relative;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }

        .kpi-card.acid { background: var(--acid); }
        .kpi-card.blue { background: var(--blue); color: var(--white); }
        .kpi-card.coral { background: var(--coral); }

        .kpi-card-category {
            font-size: .62rem;
            font-weight: 900;
            letter-spacing: .08em;
            opacity: .75;
        }

        .kpi-card-label {
            margin-top: 7px;
            font-size: .72rem;
            font-weight: 900;
            line-height: 1.2;
        }

        .kpi-card-value {
            margin-top: 10px;
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(1.45rem, 2.6vw, 2.25rem);
            line-height: 1;
        }

        .kpi-card-tooltip {
            position: absolute;
            inset: 0;
            z-index: 5;
            box-sizing: border-box;
            background: var(--ink);
            color: var(--white);
            padding: 13px 14px;
            opacity: 0;
            visibility: hidden;
            overflow-y: auto;
            font-size: .69rem;
            font-weight: 700;
            line-height: 1.4;
            transition: opacity .12s ease, visibility 0s linear .12s;
        }

        .kpi-card-tooltip strong {
            display: block;
            color: var(--acid);
            margin-bottom: 7px;
            font-size: .66rem;
            letter-spacing: .06em;
        }

        .kpi-card:hover .kpi-card-tooltip {
            opacity: 1;
            visibility: visible;
            transition: opacity .12s ease 1s, visibility 0s linear 1s;
        }

        [data-testid="stAlert"] {
            border: 3px solid var(--ink);
            border-radius: 0;
        }

        .result-title {
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(2rem, 5vw, 4.3rem);
            line-height: .95;
            margin: 38px 0 18px;
        }

        .report-title {
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(1.8rem, 4vw, 3.1rem);
            line-height: .95;
            margin: 52px 0 18px;
        }

        .report-note {
            border: 3px solid var(--ink);
            background: var(--blue);
            color: var(--white);
            box-shadow: 5px 5px 0 var(--ink);
            padding: 12px 14px;
            margin: 0 6px 22px 0;
            font-size: .78rem;
            font-weight: 800;
        }

        .origin-banner {
            border: 3px solid var(--ink);
            background: var(--acid);
            box-shadow: 6px 6px 0 var(--ink);
            padding: 14px 17px;
            margin: 30px 7px 20px 0;
            font-family: "Archivo Black", sans-serif;
            font-size: clamp(1.05rem, 2.2vw, 1.7rem);
            line-height: 1.1;
        }

        .file-pill {
            border: 2px solid var(--ink);
            background: var(--blue);
            color: white;
            padding: 7px 10px;
            display: inline-block;
            font-weight: 700;
        }

        @media (max-width: 800px) {
            .block-container { padding: 1rem 1rem 3rem; }
            .hero { padding: 22px 18px; }
            .hero::after { display: none; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def format_origin(warehouse_id: int) -> str:
    return f"{warehouse_id} - {ORIGIN_WAREHOUSES[warehouse_id]}"


def database_sheet_type(sheet_name: str) -> str:
    if sheet_name in ALEPH_SHEETS:
        return "ALEPH"
    if sheet_name in MANUAL_BACKEND_SHEETS:
        return "BACKEND"
    return "IMPORTRANGE"


def database_sheet_control(sheet_name: str) -> str:
    if sheet_name in ALEPH_SHEETS:
        return "C7"
    return "A1"


def render_kpi_cards(
    cards: list[dict[str, Any]],
    *,
    columns_count: int = 4,
) -> None:
    """Renderiza KPIs brutalistas con definición visible tras 1 s de hover."""
    if not cards:
        return
    for start in range(0, len(cards), columns_count):
        columns = st.columns(columns_count)
        for column, card in zip(columns, cards[start : start + columns_count]):
            tone = str(card.get("tone", ""))
            if tone not in {"acid", "blue", "coral"}:
                tone = ""
            class_name = f"kpi-card {tone}".strip()
            card_html = (
                f'<div class="{class_name}">'
                f'<div><div class="kpi-card-category">'
                f'{html.escape(str(card.get("category", "KPI")))}</div>'
                f'<div class="kpi-card-label">'
                f'{html.escape(str(card["label"]))}</div></div>'
                f'<div class="kpi-card-value">'
                f'{html.escape(str(card["value"]))}</div>'
                f'<div class="kpi-card-tooltip"><strong>QUÉ SIGNIFICA</strong>'
                f'{html.escape(str(card["description"]))}</div>'
                '</div>'
            )
            with column:
                st.markdown(card_html, unsafe_allow_html=True)


def contains_ref_error(value: Any) -> bool:
    return value is not None and "#REF!" in str(value).upper()


def parse_update_timestamp(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime.combine(value, datetime.min.time())
    else:
        raw = str(value or "").strip()
        if not raw:
            return None

        english = re.fullmatch(
            r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4}),?\s*"
            r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*"
            r"(AM|PM)?(?:\s+([A-Za-z]{2,5}))?",
            raw,
            flags=re.IGNORECASE,
        )
        if english:
            month_name, day, year, hour, minute, second, meridiem, zone = english.groups()
            month = ENGLISH_MONTHS.get(month_name.upper())
            if month is None:
                return None
            hour_number = int(hour)
            if meridiem and hour_number <= 12:
                if meridiem.upper() == "PM" and hour_number < 12:
                    hour_number += 12
                elif meridiem.upper() == "AM" and hour_number == 12:
                    hour_number = 0
            if hour_number > 23:
                return None
            zone_name = (zone or "").upper()
            tzinfo = (
                timezone(timedelta(hours=TIMEZONE_OFFSETS[zone_name]))
                if zone_name in TIMEZONE_OFFSETS
                else ZoneInfo("America/Mexico_City")
            )
            try:
                return datetime(
                    int(year),
                    month,
                    int(day),
                    hour_number,
                    int(minute),
                    int(second or 0),
                    tzinfo=tzinfo,
                )
            except ValueError:
                return None

        try:
            parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            parsed = None
            for date_format in (
                "%d/%m/%Y %H:%M:%S",
                "%m/%d/%Y %H:%M:%S",
                "%d-%m-%Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S",
            ):
                try:
                    parsed = datetime.strptime(raw, date_format)
                    break
                except ValueError:
                    continue
            if parsed is None:
                return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo("America/Mexico_City"))
    return parsed


def relative_update_text(
    value: Any,
    now: datetime | None = None,
) -> tuple[str, bool]:
    parsed = parse_update_timestamp(value)
    if parsed is None:
        return "C7 SIN FECHA VÁLIDA", False

    reference = now or datetime.now(ZoneInfo("America/Mexico_City"))
    if reference.tzinfo is None:
        reference = reference.replace(tzinfo=ZoneInfo("America/Mexico_City"))
    seconds = (reference.astimezone(timezone.utc) - parsed.astimezone(timezone.utc)).total_seconds()
    if seconds < -300:
        return "C7 TIENE FECHA FUTURA", False
    seconds = max(seconds, 0)

    if seconds < 60:
        return "Hace menos de 1 min", True
    if seconds < 3600:
        minutes = int(seconds // 60)
        return f"Hace {minutes} min", True
    if seconds < 86400:
        hours = int(seconds // 3600)
        return f"Hace {hours} h", True
    days = int(seconds // 86400)
    return f"Hace {days} día" if days == 1 else f"Hace {days} días", True


def inspect_database(workbook_bytes: bytes) -> dict[str, Any]:
    """Valida presencia de hojas, actualizaciones Aleph y errores IMPORTRANGE."""
    rows: list[dict[str, Any]] = []
    formula_workbook = openpyxl.load_workbook(
        io.BytesIO(workbook_bytes), read_only=True, data_only=False
    )
    value_workbook = openpyxl.load_workbook(
        io.BytesIO(workbook_bytes), read_only=True, data_only=True
    )
    try:
        for sheet_name in REQUIRED_DATABASE_SHEETS:
            if sheet_name not in formula_workbook.sheetnames:
                rows.append(
                    {
                        "HOJA": sheet_name,
                        "TIPO": database_sheet_type(sheet_name),
                        "CONTROL": database_sheet_control(sheet_name),
                        "DETALLE": "HOJA NO ENCONTRADA",
                        "ESTADO": "ERROR",
                    }
                )
                continue

            formula_sheet = formula_workbook[sheet_name]
            value_sheet = value_workbook[sheet_name]
            if sheet_name in ALEPH_SHEETS:
                value = value_sheet["C7"].value
                if value is None:
                    value = formula_sheet["C7"].value
                detail, healthy = relative_update_text(value)
                if contains_ref_error(value):
                    healthy = False
                    detail = "#REF! DETECTADO EN C7"
            elif sheet_name in MANUAL_BACKEND_SHEETS:
                formula_value = formula_sheet["A1"].value
                displayed_value = value_sheet["A1"].value
                header = engine.normalize_header(
                    displayed_value if displayed_value is not None else formula_value
                )
                healthy = header == "WAREHOUSE_ID"
                detail = (
                    "BLOQUEO BACKEND DISPONIBLE"
                    if healthy
                    else "A1 DEBE SER WAREHOUSE_ID"
                )
            else:
                formula_value = formula_sheet["A1"].value
                displayed_value = value_sheet["A1"].value
                healthy = not (
                    contains_ref_error(formula_value)
                    or contains_ref_error(displayed_value)
                )
                detail = "SIN #REF!" if healthy else "#REF! DETECTADO"

            rows.append(
                {
                    "HOJA": sheet_name,
                    "TIPO": database_sheet_type(sheet_name),
                    "CONTROL": database_sheet_control(sheet_name),
                    "DETALLE": detail,
                    "ESTADO": "OK" if healthy else "ERROR",
                }
            )
    finally:
        formula_workbook.close()
        value_workbook.close()

    healthy_count = sum(row["ESTADO"] == "OK" for row in rows)
    return {
        "online": healthy_count == len(rows),
        "healthy_count": healthy_count,
        "total_count": len(rows),
        "error_count": len(rows) - healthy_count,
        "rows": rows,
    }


def extract_available_cities(workbook_bytes: bytes) -> dict[str, str]:
    workbook = openpyxl.load_workbook(
        io.BytesIO(workbook_bytes), read_only=True, data_only=True
    )
    try:
        labels: dict[str, str] = {}
        for row in engine.iter_sheet_records(
            workbook,
            "TIENDA",
            ["CITY"],
            ["CITY"],
        ):
            raw_city = engine.clean_text(row["CITY"])
            normalized_city = engine.normalize_city(raw_city)
            if not normalized_city:
                continue
            labels.setdefault(
                normalized_city,
                CITY_DISPLAY_NAMES.get(normalized_city, raw_city or normalized_city),
            )
    finally:
        workbook.close()

    preferred_order = {"CDMX": 0, "GDL": 1, "MTY": 2}
    return dict(
        sorted(
            labels.items(),
            key=lambda item: (
                preferred_order.get(item[0], 99),
                item[1].upper(),
            ),
        )
    )


def save_uploaded_file(uploaded_file, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    uploaded_file.seek(0)
    with destination.open("wb") as handle:
        shutil.copyfileobj(uploaded_file, handle, length=8 * 1024 * 1024)
    uploaded_file.seek(0)


@st.cache_data(ttl=300, show_spinner=False)
def fetch_public_database() -> bytes:
    """Exporta el Google Sheet público completo como XLSX y conserva 5 min de caché."""
    export_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{DATA_TRANSFERS_SPREADSHEET_ID}/export?format=xlsx"
    )
    request = urllib.request.Request(
        export_url,
        headers={"User-Agent": "Mozilla/5.0 TransferPlanner/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=300) as response:
            workbook_bytes = response.read()
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
        raise RuntimeError(
            "No pude consultar la base de datos. Confirma que el Google Sheet "
            "siga configurado como 'Cualquier persona con el enlace: Lector'."
        ) from exc

    if not workbook_bytes:
        raise RuntimeError("Google devolvió una base de datos vacía.")
    if not zipfile.is_zipfile(io.BytesIO(workbook_bytes)):
        raise RuntimeError(
            "Google no devolvió un archivo Excel válido. Revisa el acceso público."
        )
    return workbook_bytes


def save_database(workbook_bytes: bytes, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        handle.write(workbook_bytes)


def load_closed_store_ids(database_path: Path) -> set[int]:
    """Carga el bloqueo permanente de tiendas definido en DATA_TRANSFERS."""
    closed_stores: set[int] = set()
    workbook = openpyxl.load_workbook(database_path, read_only=True, data_only=True)
    try:
        for record in engine.iter_sheet_records(
            workbook,
            "TIENDAS_CERRADAS",
            ["WAREHOUSE_ID"],
            ["WAREHOUSE_ID"],
        ):
            warehouse = engine.to_id(
                record["WAREHOUSE_ID"],
                "TIENDAS_CERRADAS.WAREHOUSE_ID",
                allow_none=True,
            )
            if warehouse is not None:
                closed_stores.add(warehouse)
    finally:
        workbook.close()
    return closed_stores


def load_insumos_rows(
    database_path: Path,
    catalogs,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Lee INSUMOS tal como lo entrega Aleph; no recalcula sus cantidades."""
    rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    workbook = openpyxl.load_workbook(database_path, read_only=True, data_only=True)
    try:
        for sheet_row, record in enumerate(
            engine.iter_sheet_records(
                workbook,
                "INSUMOS",
                INSUMOS_COLUMNS,
                INSUMOS_COLUMNS,
            ),
            start=14,
        ):
            destination = engine.to_id(
                record["WAREHOUSE_DESTINATION"],
                f"INSUMOS fila {sheet_row}.WAREHOUSE_DESTINATION",
                allow_none=True,
            )
            source = engine.to_id(
                record["WAREHOUSE_SOURCE"],
                f"INSUMOS fila {sheet_row}.WAREHOUSE_SOURCE",
                allow_none=True,
            )
            sku = engine.to_id(
                record["RETAIL_ID"],
                f"INSUMOS fila {sheet_row}.RETAIL_ID",
                allow_none=True,
            )
            quantity_value = engine.to_float(record["QUANTITY"], 0.0)

            if destination is None or source is None or sku is None:
                warnings.append(
                    f"INSUMOS fila {sheet_row}: se omitió por tener identificadores vacíos."
                )
                continue
            if quantity_value <= 0:
                continue
            if not math.isclose(
                quantity_value,
                round(quantity_value),
                abs_tol=1e-6,
            ):
                raise ValueError(
                    f"INSUMOS fila {sheet_row}: QUANTITY debe ser entera; "
                    f"se recibió {quantity_value!r}."
                )
            if source != 444:
                warnings.append(
                    f"INSUMOS fila {sheet_row}: se omitió el origen {source}; "
                    "los insumos solo pueden salir del warehouse 444."
                )
                continue

            store = catalogs.stores.get(destination, {})
            rows.append(
                {
                    "WAREHOUSE_DESTINATION": destination,
                    "WAREHOUSE_SOURCE": 444,
                    "RETAIL_ID": sku,
                    "QUANTITY": int(round(quantity_value)),
                    "PLANNED_DATE": "",
                    "ROUTE": 1,
                    "DELIVERY_PRIORITY": 1,
                    "CITY": store.get("city", ""),
                    "STORAGE": catalogs.storage.get(sku, "UNKNOWN"),
                    "VALUE": catalogs.high_value.get(sku, "REGULAR"),
                }
            )
    finally:
        workbook.close()
    return rows, warnings


def append_insumos_to_bulk_444(
    local_files: list[Path],
    result,
    insumos_rows: list[dict[str, Any]],
    origins: tuple[int, ...],
    include_insumos: bool,
) -> dict[str, Any]:
    """Anexa insumos sin modificar stock, capacidad ni el contador de tareas."""
    summary = {
        "requested": include_insumos,
        "enabled": include_insumos and 444 in origins,
        "source_rows": len(insumos_rows),
        "eligible_stores": 0,
        "lines_added": 0,
        "units_added": 0,
        "stores_added": 0,
    }
    if not summary["enabled"]:
        return summary

    regular_444_rows = [
        row
        for row in result.allocation_rows
        if row["WAREHOUSE_SOURCE"] == 444 and row["QUANTITY"] > 0
    ]
    eligible_destinations = {
        row["WAREHOUSE_DESTINATION"] for row in regular_444_rows
    }
    summary["eligible_stores"] = len(eligible_destinations)
    selected = [
        row
        for row in insumos_rows
        if row["WAREHOUSE_DESTINATION"] in eligible_destinations
    ]
    if not selected:
        return summary

    bulk_path = next(
        (path for path in local_files if path.name.lower() == "bulkcd_444.csv"),
        None,
    )
    if bulk_path is None:
        raise RuntimeError(
            "Se encontraron insumos elegibles, pero no se generó BulkCD_444.csv."
        )

    engine.write_csv(
        bulk_path,
        regular_444_rows + selected,
        engine.OUTPUT_COLUMNS,
    )
    summary.update(
        {
            "lines_added": len(selected),
            "units_added": sum(row["QUANTITY"] for row in selected),
            "stores_added": len(
                {row["WAREHOUSE_DESTINATION"] for row in selected}
            ),
        }
    )
    return summary


def render_database_health(health: dict[str, Any]) -> None:
    online = health["online"]
    css_class = "online" if online else "review"
    status = "ONLINE" if online else "REVISAR"
    st.markdown(
        f"""
        <div class="database-status {css_class}">
            <div class="database-title">BASE DE DATOS — {status}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    rows = health["rows"]
    for start in range(0, len(rows), 4):
        columns = st.columns(4)
        for column, row in zip(columns, rows[start : start + 4]):
            card_class = "ok" if row["ESTADO"] == "OK" else "error"
            if row["TIPO"] == "ALEPH":
                card_meta = "ÚLTIMA ACTUALIZACIÓN"
            elif row["TIPO"] == "BACKEND":
                card_meta = "BLOQUEO PERMANENTE"
            else:
                card_meta = "CONEXIÓN IMPORTRANGE"
            show_result = row["TIPO"] == "ALEPH" or row["ESTADO"] == "ERROR"
            result_html = (
                f'<div class="source-card-result">{html.escape(row["DETALLE"])}</div>'
                if show_result
                else ""
            )
            description = SHEET_DESCRIPTIONS.get(
                row["HOJA"], "Fuente de información utilizada por el modelo de abasto."
            )
            card_html = (
                f'<div class="source-card {card_class}">'
                f'<div><div class="source-card-name">{html.escape(row["HOJA"])}</div>'
                f'<div class="source-card-meta">{card_meta}</div></div>'
                f'{result_html}'
                f'<div class="source-card-tooltip"><strong>QUÉ CONTIENE</strong>'
                f'{html.escape(description)}</div>'
                '</div>'
            )
            with column:
                st.markdown(card_html, unsafe_allow_html=True)

    if not online:
        st.error(
            f"Se detectaron {health['error_count']} hojas con problemas. "
            "Revísalas antes de ejecutar la planeación."
        )


def create_zip(paths: list[Path], destination: Path) -> Path:
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            archive.write(path, arcname=path.name)
    return destination


def clear_previous_workspace() -> None:
    previous = st.session_state.pop("last_workspace", "")
    if not previous:
        return
    path = Path(previous)
    temp_root = Path(tempfile.gettempdir()).resolve()
    resolved = path.resolve()
    if temp_root in resolved.parents and resolved.name.startswith("transfer_planner_"):
        shutil.rmtree(resolved, ignore_errors=True)


def percentage(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        return 0.0
    return round((numerator / denominator) * 100, 1)


def build_planning_analytics(
    result,
    configured_origins: tuple[int, ...] = (),
) -> dict[str, Any]:
    base_rows = result.base_rows
    allocation_rows = result.allocation_rows

    def original_roq_units(row: dict[str, Any]) -> int:
        return max(int(math.ceil(max(float(row["MOV_ORIGINAL"]), 0.0))), 0)

    eligible_rows = [row for row in base_rows if row["CANTIDAD_OBJETIVO"] > 0]
    assigned_rows = [row for row in eligible_rows if row["CANTIDAD_ASIGNADA"] > 0]
    fully_covered_rows = [
        row
        for row in eligible_rows
        if row["CANTIDAD_ASIGNADA"] >= row["CANTIDAD_OBJETIVO"]
    ]
    partially_covered_rows = [
        row
        for row in eligible_rows
        if 0 < row["CANTIDAD_ASIGNADA"] < row["CANTIDAD_OBJETIVO"]
    ]
    not_assigned_rows = [
        row for row in eligible_rows if row["CANTIDAD_ASIGNADA"] <= 0
    ]

    city_accumulators: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "stores": set(),
            "products": set(),
            "units": 0,
            "tasks": 0,
            "m3": 0.0,
        }
    )
    store_accumulators: dict[tuple[str, int, str], dict[str, Any]] = defaultdict(
        lambda: {"products": set(), "units": 0, "tasks": 0, "m3": 0.0}
    )

    for row in assigned_rows:
        city = row.get("CITY") or "SIN CIUDAD"
        destination = row["WAREHOUSE_DESTINATION"]
        warehouse_name = row.get("WAREHOUSE_NAME") or "SIN NOMBRE"
        sku = row["RETAIL_ID"]
        units = int(row["CANTIDAD_ASIGNADA"])
        tasks = int(row["TAREAS_GENERADAS"])
        assigned_m3 = float(row["M3_ASIGNADO"])

        city_data = city_accumulators[city]
        city_data["stores"].add(destination)
        city_data["products"].add(sku)
        city_data["units"] += units
        city_data["tasks"] += tasks
        city_data["m3"] += assigned_m3

        store_data = store_accumulators[(city, destination, warehouse_name)]
        store_data["products"].add(sku)
        store_data["units"] += units
        store_data["tasks"] += tasks
        store_data["m3"] += assigned_m3

    city_rows = [
        {
            "CIUDAD": city,
            "TIENDAS_ATENDIDAS": len(data["stores"]),
            "PRODUCTOS_DISTINTOS": len(data["products"]),
            "UNIDADES": data["units"],
            "M3": round(data["m3"], 3),
            "TAREAS": data["tasks"],
        }
        for city, data in sorted(city_accumulators.items())
    ]
    store_rows = [
        {
            "CIUDAD": city,
            "WAREHOUSE_ID": destination,
            "TIENDA": warehouse_name,
            "PRODUCTOS_DISTINTOS": len(data["products"]),
            "UNIDADES": data["units"],
            "M3": round(data["m3"], 3),
            "TAREAS": data["tasks"],
        }
        for (city, destination, warehouse_name), data in store_accumulators.items()
    ]
    store_rows.sort(key=lambda row: (row["CIUDAD"], -row["UNIDADES"], row["WAREHOUSE_ID"]))

    stockout_rows = [
        row for row in eligible_rows if bool(row.get("ES_STOCKOUT", False))
    ]
    stockout_served = [row for row in stockout_rows if row["CANTIDAD_ASIGNADA"] > 0]
    stockout_fully_covered = [
        row
        for row in stockout_rows
        if row["CANTIDAD_ASIGNADA"] >= row["CANTIDAD_OBJETIVO"]
    ]
    forecast_zero_forced = [
        row
        for row in eligible_rows
        if row["REGLA_DEMANDA"] == "HARDCODE_4_CERO_TOTAL"
    ]
    forecast_zero_served = [
        row for row in forecast_zero_forced if row["CANTIDAD_ASIGNADA"] > 0
    ]
    deficit_forced = [
        row
        for row in eligible_rows
        if row["REGLA_DEMANDA"] == "HARDCODE_3_INVENTARIO_MENOR_DEMANDA"
    ]
    minimum_three_applied = [
        row
        for row in eligible_rows
        if 0 < row["MOV_ORIGINAL"] < 3 and row["CANTIDAD_OBJETIVO"] == 3
    ]
    golden_rows = [
        row for row in eligible_rows if bool(row.get("ES_GOLDEN_INFALTABLE", False))
    ]

    rule_accumulators: dict[str, dict[str, float]] = defaultdict(
        lambda: {
            "cases": 0,
            "served": 0,
            "full": 0,
            "original_roq_units": 0,
            "hardcode_target_units": 0,
            "hardcode_assigned_units": 0,
            "target_units": 0,
            "assigned_units": 0,
        }
    )
    for row in base_rows:
        rule = row["REGLA_DEMANDA"]
        data = rule_accumulators[rule]
        target = int(row["CANTIDAD_OBJETIVO"])
        assigned = int(row["CANTIDAD_ASIGNADA"])
        original_roq = original_roq_units(row)
        hardcode_target = max(target - original_roq, 0)
        hardcode_assigned = max(assigned - original_roq, 0)
        data["cases"] += 1
        data["served"] += int(assigned > 0)
        data["full"] += int(target > 0 and assigned >= target)
        data["original_roq_units"] += original_roq
        data["hardcode_target_units"] += hardcode_target
        data["hardcode_assigned_units"] += hardcode_assigned
        data["target_units"] += target
        data["assigned_units"] += assigned

    rule_order = {rule: index for index, rule in enumerate(DEMAND_RULE_LABELS)}
    rule_rows = []
    for rule, data in sorted(
        rule_accumulators.items(), key=lambda item: rule_order.get(item[0], 999)
    ):
        rule_rows.append(
            {
                "REGLA": DEMAND_RULE_LABELS.get(rule, rule),
                "CASOS": int(data["cases"]),
                "CASOS_CON_ENVIO": int(data["served"]),
                "COBERTURA_COMPLETA": int(data["full"]),
                "UNIDADES_ROQ_ORIGINAL": int(data["original_roq_units"]),
                "INCREMENTO_HARDCODE_OBJETIVO": int(
                    data["hardcode_target_units"]
                ),
                "INCREMENTO_HARDCODE_ENVIADO": int(
                    data["hardcode_assigned_units"]
                ),
                "UNIDADES_OBJETIVO": int(data["target_units"]),
                "UNIDADES_ASIGNADAS": int(data["assigned_units"]),
                "COMPLIANCE_UNIDADES_%": percentage(
                    data["assigned_units"], data["target_units"]
                ),
            }
        )

    stockout_city_accumulators: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "cases": 0,
            "served": 0,
            "full": 0,
            "products": set(),
            "units": 0,
        }
    )
    for row in stockout_rows:
        city = row.get("CITY") or "SIN CIUDAD"
        data = stockout_city_accumulators[city]
        assigned = int(row["CANTIDAD_ASIGNADA"])
        data["cases"] += 1
        if assigned > 0:
            data["served"] += 1
            data["products"].add(row["RETAIL_ID"])
            data["units"] += assigned
        if assigned >= row["CANTIDAD_OBJETIVO"]:
            data["full"] += 1

    stockout_city_rows = [
        {
            "CIUDAD": city,
            "CASOS_STOCKOUT": data["cases"],
            "CASOS_CON_ENVIO": data["served"],
            "COBERTURA_COMPLETA": data["full"],
            "PRODUCTOS_DISTINTOS_ATENDIDOS": len(data["products"]),
            "UNIDADES_ASIGNADAS": data["units"],
            "ATENCION_%": percentage(data["served"], data["cases"]),
        }
        for city, data in sorted(stockout_city_accumulators.items())
    ]

    m3_by_destination_sku = {
        (row["WAREHOUSE_DESTINATION"], row["RETAIL_ID"]): row["M3_POR_UNIDAD"]
        for row in base_rows
    }
    base_by_destination_sku = {
        (row["WAREHOUSE_DESTINATION"], row["RETAIL_ID"]): row
        for row in base_rows
    }
    source_accumulators: dict[int, dict[str, Any]] = defaultdict(
        lambda: {
            "tasks": 0,
            "units": 0,
            "products": set(),
            "stores": set(),
            "m3": 0.0,
        }
    )
    source_city_accumulators: dict[tuple[int, str], dict[str, Any]] = defaultdict(
        lambda: {
            "stores": set(),
            "products": set(),
            "tasks": 0,
            "units": 0,
            "m3": 0.0,
        }
    )
    source_store_accumulators: dict[
        tuple[int, str, int, str], dict[str, Any]
    ] = defaultdict(
        lambda: {"products": set(), "tasks": 0, "units": 0, "m3": 0.0}
    )
    source_storage_accumulators: dict[
        tuple[int, str], dict[str, Any]
    ] = defaultdict(
        lambda: {
            "products": set(),
            "stores": set(),
            "tasks": 0,
            "units": 0,
            "m3": 0.0,
        }
    )

    source_order = list(dict.fromkeys(configured_origins))
    for source in source_order:
        source_accumulators[source]

    for row in allocation_rows:
        source = row["WAREHOUSE_SOURCE"]
        if source not in source_order:
            source_order.append(source)
        data = source_accumulators[source]
        quantity = int(row["QUANTITY"])
        destination = row["WAREHOUSE_DESTINATION"]
        sku = row["RETAIL_ID"]
        base_row = base_by_destination_sku.get((destination, sku), {})
        city = row.get("CITY") or base_row.get("CITY") or "SIN CIUDAD"
        warehouse_name = (
            base_row.get("WAREHOUSE_NAME") or f"WAREHOUSE {destination}"
        )
        storage = row.get("STORAGE") or "UNKNOWN"
        line_m3 = quantity * float(
            m3_by_destination_sku.get((destination, sku), 0.0)
        )
        data["tasks"] += 1
        data["units"] += quantity
        data["products"].add(sku)
        data["stores"].add(destination)
        data["m3"] += line_m3

        city_data = source_city_accumulators[(source, city)]
        city_data["stores"].add(destination)
        city_data["products"].add(sku)
        city_data["tasks"] += 1
        city_data["units"] += quantity
        city_data["m3"] += line_m3

        store_data = source_store_accumulators[
            (source, city, destination, warehouse_name)
        ]
        store_data["products"].add(sku)
        store_data["tasks"] += 1
        store_data["units"] += quantity
        store_data["m3"] += line_m3

        storage_data = source_storage_accumulators[(source, storage)]
        storage_data["products"].add(sku)
        storage_data["stores"].add(destination)
        storage_data["tasks"] += 1
        storage_data["units"] += quantity
        storage_data["m3"] += line_m3

    total_source_tasks = sum(
        source_accumulators[source]["tasks"] for source in source_order
    )
    total_source_units = sum(
        source_accumulators[source]["units"] for source in source_order
    )
    source_rows: list[dict[str, Any]] = []
    for source in source_order:
        data = source_accumulators[source]
        source_rows.append(
            {
                "WAREHOUSE_SOURCE": source,
                "NOMBRE": ORIGIN_WAREHOUSES.get(source, "ORIGEN CONFIGURADO"),
                "TAREAS": data["tasks"],
                "UNIDADES": data["units"],
                "PRODUCTOS_DISTINTOS": len(data["products"]),
                "TIENDAS_ATENDIDAS": len(data["stores"]),
                "M3": round(data["m3"], 3),
                "UNIDADES_POR_TAREA": round(
                    data["units"] / data["tasks"], 2
                )
                if data["tasks"]
                else 0.0,
                "PARTICIPACION_TAREAS_%": percentage(
                    data["tasks"], total_source_tasks
                ),
                "PARTICIPACION_UNIDADES_%": percentage(
                    data["units"], total_source_units
                ),
            }
        )

    source_details: list[dict[str, Any]] = []
    for source, source_summary in zip(source_order, source_rows):
        city_detail_rows = [
            {
                "CIUDAD": city,
                "TIENDAS_ATENDIDAS": len(data["stores"]),
                "PRODUCTOS_DISTINTOS": len(data["products"]),
                "TAREAS": data["tasks"],
                "UNIDADES": data["units"],
                "M3": round(data["m3"], 3),
            }
            for (row_source, city), data in source_city_accumulators.items()
            if row_source == source
        ]
        city_detail_rows.sort(key=lambda row: (-row["UNIDADES"], row["CIUDAD"]))

        store_detail_rows = [
            {
                "CIUDAD": city,
                "WAREHOUSE_ID": destination,
                "TIENDA": warehouse_name,
                "PRODUCTOS_DISTINTOS": len(data["products"]),
                "TAREAS": data["tasks"],
                "UNIDADES": data["units"],
                "M3": round(data["m3"], 3),
            }
            for (
                row_source,
                city,
                destination,
                warehouse_name,
            ), data in source_store_accumulators.items()
            if row_source == source
        ]
        store_detail_rows.sort(
            key=lambda row: (-row["UNIDADES"], row["WAREHOUSE_ID"])
        )

        storage_detail_rows = [
            {
                "STORAGE": storage,
                "PRODUCTOS_DISTINTOS": len(data["products"]),
                "TIENDAS_ATENDIDAS": len(data["stores"]),
                "TAREAS": data["tasks"],
                "UNIDADES": data["units"],
                "M3": round(data["m3"], 3),
            }
            for (
                row_source,
                storage,
            ), data in source_storage_accumulators.items()
            if row_source == source
        ]
        storage_detail_rows.sort(
            key=lambda row: (-row["UNIDADES"], row["STORAGE"])
        )

        source_details.append(
            {
                "warehouse_source": source,
                "name": ORIGIN_WAREHOUSES.get(source, "ORIGEN CONFIGURADO"),
                "summary": source_summary,
                "city_rows": city_detail_rows,
                "store_rows": store_detail_rows,
                "storage_rows": storage_detail_rows,
            }
        )

    target_units = sum(int(row["CANTIDAD_OBJETIVO"]) for row in eligible_rows)
    assigned_units = sum(int(row["CANTIDAD_ASIGNADA"]) for row in eligible_rows)
    original_roq_total = sum(original_roq_units(row) for row in eligible_rows)
    original_roq_fulfilled = sum(
        min(int(row["CANTIDAD_ASIGNADA"]), original_roq_units(row))
        for row in eligible_rows
    )
    hardcode_target_units = sum(
        max(int(row["CANTIDAD_OBJETIVO"]) - original_roq_units(row), 0)
        for row in eligible_rows
    )
    hardcode_assigned_units = sum(
        max(int(row["CANTIDAD_ASIGNADA"]) - original_roq_units(row), 0)
        for row in eligible_rows
    )
    hardcode_cases = sum(
        int(row["CANTIDAD_OBJETIVO"]) > original_roq_units(row)
        for row in eligible_rows
    )
    return {
        "city_rows": city_rows,
        "store_rows": store_rows,
        "rule_rows": rule_rows,
        "stockout_city_rows": stockout_city_rows,
        "source_rows": source_rows,
        "source_details": source_details,
        "summary": {
            "eligible_cases": len(eligible_rows),
            "fully_covered_cases": len(fully_covered_rows),
            "partial_cases": len(partially_covered_rows),
            "not_assigned_cases": len(not_assigned_rows),
            "target_units": target_units,
            "assigned_units": assigned_units,
            "case_compliance_pct": percentage(
                len(fully_covered_rows), len(eligible_rows)
            ),
            "case_service_pct": percentage(len(assigned_rows), len(eligible_rows)),
            "unit_compliance_pct": percentage(assigned_units, target_units),
            "original_roq_units": original_roq_total,
            "original_roq_fulfilled_units": original_roq_fulfilled,
            "original_roq_compliance_pct": percentage(
                original_roq_fulfilled, original_roq_total
            ),
            "hardcode_cases": hardcode_cases,
            "hardcode_target_units": hardcode_target_units,
            "hardcode_assigned_units": hardcode_assigned_units,
            "m3_assigned": round(
                sum(float(row["M3_ASIGNADO"]) for row in assigned_rows), 3
            ),
            "cities_served": len(city_rows),
            "stores_served": len(store_rows),
            "products_served": len({row["RETAIL_ID"] for row in assigned_rows}),
            "stockout_cases": len(stockout_rows),
            "stockout_cases_served": len(stockout_served),
            "stockout_cases_full": len(stockout_fully_covered),
            "stockout_products_served": len(
                {row["RETAIL_ID"] for row in stockout_served}
            ),
            "stockout_stores_served": len(
                {row["WAREHOUSE_DESTINATION"] for row in stockout_served}
            ),
            "stockout_attention_pct": percentage(
                len(stockout_served), len(stockout_rows)
            ),
            "forecast_zero_forced_cases": len(forecast_zero_forced),
            "forecast_zero_served_cases": len(forecast_zero_served),
            "forecast_zero_target_units": sum(
                int(row["CANTIDAD_OBJETIVO"]) for row in forecast_zero_forced
            ),
            "forecast_zero_assigned_units": sum(
                int(row["CANTIDAD_ASIGNADA"]) for row in forecast_zero_forced
            ),
            "deficit_forced_cases": len(deficit_forced),
            "deficit_forced_served_cases": sum(
                row["CANTIDAD_ASIGNADA"] > 0 for row in deficit_forced
            ),
            "minimum_three_cases": len(minimum_three_applied),
            "golden_cases": len(golden_rows),
            "golden_served_cases": sum(
                row["CANTIDAD_ASIGNADA"] > 0 for row in golden_rows
            ),
        },
    }


def apply_reporting_labels(result) -> None:
    """Cambia únicamente etiquetas de salida; las reglas internas no se alteran."""
    for row in result.base_rows:
        cut_type = str(row.get("TIPO_DE_CORTE", ""))
        assigned = int(row.get("CANTIDAD_ASIGNADA", 0) or 0)
        target = int(row.get("CANTIDAD_OBJETIVO", 0) or 0)
        stock_excluded_by_rack_444 = max(
            float(row.get("STOCK_BASE_444", 0) or 0)
            - float(row.get("NO_DISPONIBLE_444", 0) or 0)
            - float(row.get("COPERNICO_NO_USABLE_444", 0) or 0),
            0.0,
        )
        row["STOCK_EXCLUIDO_RACKEADO_444"] = stock_excluded_by_rack_444
        if (
            bool(row.get("RACKEADO_444", False))
            and stock_excluded_by_rack_444 > 0
            and assigned < target
            and "CORTE POR STOCK" in cut_type
        ):
            if assigned > 0:
                row["TIPO_DE_CORTE"] = (
                    "OK PARCIAL - CORTE POR PRODUCTO RACKEADO 444"
                )
                row["DETALLE_MOTIVO"] = (
                    f"Asignadas {assigned} de {target}; el SKU está excluido al "
                    "100% del stock utilizable del warehouse 444 por RACKEADOS."
                )
            else:
                row["TIPO_DE_CORTE"] = "CORTE POR PRODUCTO RACKEADO 444"
                row["DETALLE_MOTIVO"] = (
                    "El SKU está excluido al 100% del stock utilizable del "
                    "warehouse 444 por RACKEADOS."
                )

        if row.get("TIPO_DE_CORTE") == "SIN DEMANDA":
            row["TIPO_DE_CORTE"] = "SIN RECOMENDACIÓN"

        detail = row.get("DETALLE_MOTIVO")
        if isinstance(detail, str):
            row["DETALLE_MOTIVO"] = detail.replace("MOV", "ROQ").replace(
                "Sin demanda", "Sin recomendación"
            )

        demand_rule = row.get("REGLA_DEMANDA")
        if isinstance(demand_rule, str):
            row["REGLA_DEMANDA"] = demand_rule.replace(
                "SIN_DEMANDA", "SIN_RECOMENDACION"
            ).replace("MOV", "ROQ")

        if "MOV_ORIGINAL" in row:
            renamed_row: dict[str, Any] = {}
            for key, value in row.items():
                renamed_row["ROQ_ORIGINAL" if key == "MOV_ORIGINAL" else key] = value
            row.clear()
            row.update(renamed_row)


def split_plan_rows_by_closed_store(
    plan_rows: list[dict[str, Any]],
    closed_store_ids: set[int],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    active_rows: list[dict[str, Any]] = []
    blocked_rows: list[dict[str, Any]] = []
    for row in plan_rows:
        if row["WAREHOUSE_DESTINATION"] in closed_store_ids:
            blocked_rows.append(row)
        else:
            active_rows.append(row)
    return active_rows, blocked_rows


def closed_store_summary(
    blocked_rows: list[dict[str, Any]],
    closed_store_ids: set[int],
    config: engine.Config,
) -> dict[str, Any]:
    impacted_stores = {
        row["WAREHOUSE_DESTINATION"] for row in blocked_rows
    }
    products = {row["RETAIL_ID"] for row in blocked_rows}
    target_units = sum(
        engine.calculate_target_quantity(row, config)[0] for row in blocked_rows
    )
    return {
        "configured_stores": len(closed_store_ids),
        "requirements": len(blocked_rows),
        "stores": len(impacted_stores),
        "store_ids": sorted(impacted_stores),
        "products": len(products),
        "target_units": target_units,
    }


def split_plan_rows_by_blocked_city(
    plan_rows: list[dict[str, Any]],
    catalogs,
    blocked_cities: tuple[str, ...],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    blocked_city_set = set(blocked_cities)
    active_rows: list[dict[str, Any]] = []
    blocked_rows: list[dict[str, Any]] = []
    for row in plan_rows:
        store = catalogs.stores.get(row["WAREHOUSE_DESTINATION"], {})
        if store.get("city_norm", "") in blocked_city_set:
            blocked_rows.append(row)
        else:
            active_rows.append(row)
    return active_rows, blocked_rows


def city_block_summary(
    blocked_rows: list[dict[str, Any]],
    catalogs,
    config: engine.Config,
    blocked_cities: tuple[str, ...],
) -> dict[str, Any]:
    rows_by_city: Counter[str] = Counter()
    city_names: dict[str, str] = {}
    for store in catalogs.stores.values():
        city_norm = store.get("city_norm", "")
        if city_norm in blocked_cities and city_norm not in city_names:
            city_names[city_norm] = CITY_DISPLAY_NAMES.get(
                city_norm,
                store.get("city") or city_norm,
            )
    target_units = 0
    stores: set[int] = set()
    products: set[int] = set()
    for row in blocked_rows:
        destination = row["WAREHOUSE_DESTINATION"]
        store = catalogs.stores.get(destination, {})
        city_norm = store.get("city_norm", "")
        rows_by_city[city_norm] += 1
        stores.add(destination)
        products.add(row["RETAIL_ID"])
        target, _ = engine.calculate_target_quantity(row, config)
        target_units += target

    return {
        "cities": [
            {
                "code": city,
                "name": city_names.get(city, CITY_DISPLAY_NAMES.get(city, city)),
                "requirements": rows_by_city.get(city, 0),
            }
            for city in blocked_cities
        ],
        "requirements": len(blocked_rows),
        "stores": len(stores),
        "products": len(products),
        "target_units": target_units,
    }


def execute_planning(
    uploaded_plan,
    database_bytes: bytes,
    origins: tuple[int, ...],
    max_tasks: int,
    run_date,
    blocked_cities: tuple[str, ...] = (),
    include_insumos: bool = True,
) -> dict[str, Any]:
    clear_previous_workspace()
    workspace = Path(tempfile.mkdtemp(prefix="transfer_planner_"))
    st.session_state["last_workspace"] = str(workspace)
    plan_path = workspace / "input" / Path(uploaded_plan.name).name
    data_path = workspace / "input" / "DATA_TRANSFERS.xlsx"

    save_uploaded_file(uploaded_plan, plan_path)
    save_database(database_bytes, data_path)

    config = engine.Config(
        origin_warehouses=origins,
        max_tasks=max_tasks,
        run_date_override=run_date.strftime("%d-%m-%Y"),
        default_store_capacity_m3=engine.CONFIG.default_store_capacity_m3,
        default_m3_per_unit=engine.CONFIG.default_m3_per_unit,
        minimum_positive_quantity=engine.CONFIG.minimum_positive_quantity,
        local_work_dir=str(workspace / "engine"),
        replace_same_day_outputs=True,
        generate_empty_source_files=False,
    )

    captured = io.StringIO()
    with redirect_stdout(captured):
        catalogs = engine.load_catalogs(data_path, config)
        closed_store_ids = load_closed_store_ids(data_path)
        insumos_rows: list[dict[str, Any]] = []
        if include_insumos and 444 in origins:
            insumos_rows, insumos_warnings = load_insumos_rows(
                data_path,
                catalogs,
            )
            catalogs.warnings.extend(insumos_warnings)
        plan_read = engine.read_plan_csv(plan_path, config)
        rows_after_closed_stores, closed_plan_rows = (
            split_plan_rows_by_closed_store(
                plan_read.rows,
                closed_store_ids,
            )
        )
        active_plan_rows, blocked_plan_rows = split_plan_rows_by_blocked_city(
            rows_after_closed_stores,
            catalogs,
            blocked_cities,
        )

        result = engine.plan_transfers(active_plan_rows, catalogs, config)
        result.warnings.extend(plan_read.warnings)
        closed_summary = closed_store_summary(
            closed_plan_rows,
            closed_store_ids,
            config,
        )
        if closed_plan_rows:
            closed_ids_text = ", ".join(map(str, closed_summary["store_ids"]))
            result.warnings.append(
                "Bloqueo permanente TIENDAS_CERRADAS: se excluyeron "
                f"{len(closed_plan_rows):,} requerimientos de las tiendas "
                f"{closed_ids_text} antes de asignar stock."
            )
        block_summary = city_block_summary(
            blocked_plan_rows,
            catalogs,
            config,
            blocked_cities,
        )
        if blocked_plan_rows:
            blocked_names = ", ".join(
                item["name"] for item in block_summary["cities"]
            )
            result.warnings.append(
                f"Bloqueo manual de ciudad: {blocked_names}. Se excluyeron "
                f"{len(blocked_plan_rows):,} requerimientos antes de asignar stock."
            )

        analytics = build_planning_analytics(result, origins)
        apply_reporting_labels(result)
        output_dir = (
            Path(config.local_work_dir)
            / "outputs"
            / run_date.strftime("%d-%m-%Y")
        )
        local_files = engine.create_output_files(
            result,
            config,
            run_date,
            plan_path.name,
            output_dir,
        )
        local_files = [Path(path) for path in local_files]
        insumos_summary = append_insumos_to_bulk_444(
            local_files,
            result,
            insumos_rows,
            origins,
            include_insumos,
        )
        print(f"Requerimientos únicos del input: {len(plan_read.rows):,}")
        print(f"Requerimientos activos: {len(active_plan_rows):,}")
        print(
            "Requerimientos excluidos por tiendas cerradas: "
            f"{len(closed_plan_rows):,}"
        )
        print(f"Requerimientos excluidos por ciudad: {len(blocked_plan_rows):,}")
        print(f"Tareas generadas: {result.tasks_used:,}")
        if insumos_summary["enabled"]:
            print(
                "Insumos agregados a BulkCD_444: "
                f"{insumos_summary['lines_added']:,} líneas / "
                f"{insumos_summary['units_added']:,} unidades / "
                f"{insumos_summary['stores_added']:,} tiendas"
            )

    zip_path = create_zip(
        local_files,
        workspace / f"Planeacion_{run_date:%d-%m-%Y}.zip",
    )
    status_counts = Counter(row["TIPO_DE_CORTE"] for row in result.base_rows)
    if closed_summary["requirements"]:
        status_counts["TIENDA CERRADA - BLOQUEO BACKEND"] += closed_summary[
            "requirements"
        ]
    if block_summary["requirements"]:
        status_counts["CIUDAD BLOQUEADA MANUALMENTE"] += block_summary[
            "requirements"
        ]
    units = sum(row["QUANTITY"] for row in result.allocation_rows)
    requirements = len(result.base_rows)
    return {
        "workspace": str(workspace),
        "files": [str(path) for path in local_files],
        "zip": str(zip_path),
        "tasks": result.tasks_used,
        "units": units,
        "requirements": requirements,
        "input_requirements": len(plan_read.rows),
        "status_counts": dict(status_counts),
        "warnings": list(result.warnings),
        "logs": captured.getvalue(),
        "origins": list(origins),
        "analytics": analytics,
        "city_block": block_summary,
        "closed_stores": closed_summary,
        "insumos": insumos_summary,
    }


def report_table(
    rows: list[dict[str, Any]],
    *,
    column_config: dict[str, Any] | None = None,
    max_height: int = 560,
) -> None:
    if not rows:
        st.info("No existen registros para mostrar en esta sección.")
        return
    height = min(max(150, 36 * (len(rows) + 1)), max_height)
    st.dataframe(
        rows,
        use_container_width=True,
        hide_index=True,
        height=height,
        column_config=column_config or {},
    )


def render_source_analysis(analytics: dict[str, Any]) -> None:
    source_details = analytics.get("source_details", [])
    if not source_details:
        return

    st.markdown(
        '<div class="report-title">ANÁLISIS POR ORIGEN.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="report-note">
            Cada bloque corresponde a un warehouse origen seleccionado. TAREAS son
            líneas operativas del Bulk; UNIDADES son la suma de QUANTITY; PRODUCTOS son
            SKUs distintos. Este análisis considera únicamente abasto normal y excluye
            insumos para mantener comparables los orígenes. Mantén el cursor un segundo
            sobre cualquier tarjeta para consultar su definición.
        </div>
        """,
        unsafe_allow_html=True,
    )

    for detail in source_details:
        source = detail["warehouse_source"]
        name = detail["name"]
        summary = detail["summary"]
        st.markdown(
            f'<div class="origin-banner">{source} — {html.escape(str(name))}</div>',
            unsafe_allow_html=True,
        )
        render_kpi_cards(
            [
                {
                    "category": "TAREAS · ORIGEN",
                    "label": "LÍNEAS DE ABASTO",
                    "value": f"{summary['TAREAS']:,}",
                    "description": (
                        "Número de filas de transferencia generadas desde este origen. "
                        "Una misma tienda–SKU puede crear una tarea en más de un origen. "
                        "No incluye líneas de insumos."
                    ),
                    "tone": "blue",
                },
                {
                    "category": "UNIDADES · ORIGEN",
                    "label": "PRODUCTO ASIGNADO",
                    "value": f"{summary['UNIDADES']:,}",
                    "description": (
                        "Suma de QUANTITY de producto normal que saldrá desde este "
                        "warehouse origen. No incluye insumos."
                    ),
                    "tone": "acid",
                },
                {
                    "category": "SKUs · ORIGEN",
                    "label": "PRODUCTOS DISTINTOS",
                    "value": f"{summary['PRODUCTOS_DISTINTOS']:,}",
                    "description": (
                        "Cantidad de RETAIL_ID diferentes asignados desde este origen. "
                        "No es un conteo de tareas ni de unidades."
                    ),
                },
                {
                    "category": "TIENDAS · ORIGEN",
                    "label": "DESTINOS ATENDIDOS",
                    "value": f"{summary['TIENDAS_ATENDIDAS']:,}",
                    "description": (
                        "Número de warehouses destino diferentes que reciben producto "
                        "normal desde este origen."
                    ),
                },
                {
                    "category": "M³ · ORIGEN",
                    "label": "VOLUMEN PLANEADO",
                    "value": f"{summary['M3']:,.3f}",
                    "description": (
                        "Volumen de producto normal asignado desde este origen: unidades "
                        "por metros cúbicos por unidad."
                    ),
                },
                {
                    "category": "UNIDADES / TAREA",
                    "label": "PROMEDIO POR LÍNEA",
                    "value": f"{summary['UNIDADES_POR_TAREA']:,.2f}",
                    "description": (
                        "Unidades asignadas desde este origen divididas entre sus tareas. "
                        "Sirve para entender el tamaño promedio de cada línea operativa."
                    ),
                },
                {
                    "category": "PARTICIPACIÓN · TAREAS",
                    "label": "PESO OPERATIVO",
                    "value": f"{summary['PARTICIPACION_TAREAS_%']:,.1f}%",
                    "description": (
                        "Porcentaje de todas las tareas de abasto que salen desde este "
                        "origen. No utiliza unidades como denominador."
                    ),
                    "tone": "blue",
                },
                {
                    "category": "PARTICIPACIÓN · UNIDADES",
                    "label": "PESO EN UNIDADES",
                    "value": f"{summary['PARTICIPACION_UNIDADES_%']:,.1f}%",
                    "description": (
                        "Porcentaje de todas las unidades de abasto asignadas a este "
                        "origen. No utiliza tareas como denominador."
                    ),
                    "tone": "acid",
                },
            ]
        )

        left, right = st.columns(2)
        with left:
            st.markdown(
                '<span class="section-label">POR CIUDAD</span>',
                unsafe_allow_html=True,
            )
            report_table(
                detail["city_rows"],
                column_config={
                    "M3": st.column_config.NumberColumn("M³", format="%.3f"),
                },
                max_height=340,
            )
        with right:
            st.markdown(
                '<span class="section-label">POR STORAGE</span>',
                unsafe_allow_html=True,
            )
            report_table(
                detail["storage_rows"],
                column_config={
                    "M3": st.column_config.NumberColumn("M³", format="%.3f"),
                },
                max_height=340,
            )

        st.markdown(
            '<span class="section-label">TIENDAS ATENDIDAS DESDE ESTE ORIGEN</span>',
            unsafe_allow_html=True,
        )
        report_table(
            detail["store_rows"],
            column_config={
                "WAREHOUSE_ID": st.column_config.NumberColumn(
                    "WAREHOUSE ID", format="%d"
                ),
                "M3": st.column_config.NumberColumn("M³", format="%.3f"),
            },
            max_height=520,
        )


def render_planning_analytics(analytics: dict[str, Any]) -> None:
    summary = analytics["summary"]

    st.markdown(
        '<div class="report-title">CIUDADES + TIENDAS.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="report-note">
            PRODUCTOS = SKUs distintos con al menos una unidad asignada · M³ = volumen
            realmente planeado · TAREAS = líneas generadas considerando cada origen.
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_kpi_cards(
        [
            {
                "category": "COBERTURA · GEOGRAFÍA",
                "label": "CIUDADES ATENDIDAS",
                "value": f"{summary['cities_served']:,}",
                "description": (
                    "Número de ciudades distintas con al menos una unidad de producto "
                    "asignada. No incluye insumos."
                ),
            },
            {
                "category": "COBERTURA · TIENDAS",
                "label": "TIENDAS ATENDIDAS",
                "value": f"{summary['stores_served']:,}",
                "description": (
                    "Warehouses destino distintos que reciben al menos una unidad de "
                    "producto normal."
                ),
                "tone": "acid",
            },
            {
                "category": "COBERTURA · SKUs",
                "label": "PRODUCTOS DISTINTOS",
                "value": f"{summary['products_served']:,}",
                "description": (
                    "RETAIL_ID distintos con al menos una unidad asignada en toda la "
                    "planeación. No equivale al número de tareas."
                ),
            },
            {
                "category": "VOLUMEN · PRODUCTO",
                "label": "M³ PLANEADOS",
                "value": f"{summary['m3_assigned']:,.3f}",
                "description": (
                    "Suma del volumen de producto normal asignado: QUANTITY por metros "
                    "cúbicos por unidad. No incluye insumos."
                ),
                "tone": "blue",
            },
        ]
    )

    st.markdown('<span class="section-label">RESUMEN POR CIUDAD</span>', unsafe_allow_html=True)
    report_table(
        analytics["city_rows"],
        column_config={
            "M3": st.column_config.NumberColumn("M³", format="%.3f"),
        },
        max_height=360,
    )

    st.markdown('<span class="section-label">DETALLE POR TIENDA</span>', unsafe_allow_html=True)
    report_table(
        analytics["store_rows"],
        column_config={
            "WAREHOUSE_ID": st.column_config.NumberColumn(
                "WAREHOUSE ID", format="%d"
            ),
            "M3": st.column_config.NumberColumn("M³", format="%.3f"),
        },
        max_height=600,
    )

    st.markdown(
        '<div class="report-title">REPORTE DE REABASTO.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="report-note">
            Un caso representa una combinación tienda–SKU. COMPLIANCE DE CASOS mide
            cuántos casos se cubrieron al 100%; COMPLIANCE DE UNIDADES compara lo
            asignado contra el objetivo final del modelo. El incremento por hardcode se
            presenta separado del ROQ original para no inflar artificialmente el
            cumplimiento. Mantén el cursor un segundo sobre cualquier tarjeta para ver
            exactamente qué mide y cuál es su denominador.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<span class="section-label">CASOS TIENDA–SKU</span>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "CASOS · BASE",
                "label": "CON RECOMENDACIÓN",
                "value": f"{summary['eligible_cases']:,}",
                "description": (
                    "Combinaciones tienda–SKU con CANTIDAD_OBJETIVO mayor a cero "
                    "después de aplicar ROQ y hardcodes."
                ),
            },
            {
                "category": "CASOS · COMPLETOS",
                "label": "CUBIERTOS AL 100%",
                "value": f"{summary['fully_covered_cases']:,}",
                "description": (
                    "Casos cuya cantidad asignada alcanzó completamente la cantidad "
                    "objetivo. Un caso no es lo mismo que una tarea."
                ),
                "tone": "acid",
            },
            {
                "category": "CASOS · PARCIALES",
                "label": "COBERTURA PARCIAL",
                "value": f"{summary['partial_cases']:,}",
                "description": (
                    "Casos que reciben al menos una unidad, pero no cubren la cantidad "
                    "objetivo completa."
                ),
                "tone": "coral",
            },
            {
                "category": "CASOS · SIN SALIDA",
                "label": "SIN ENVÍO",
                "value": f"{summary['not_assigned_cases']:,}",
                "description": (
                    "Casos con recomendación positiva a los que no se asignó ninguna "
                    "unidad por las reglas de corte."
                ),
                "tone": "coral",
            },
        ]
    )

    st.markdown('<span class="section-label">COMPLIANCE</span>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "PORCENTAJE · CASOS",
                "label": "COMPLIANCE DE CASOS",
                "value": f"{summary['case_compliance_pct']:,.1f}%",
                "description": (
                    "Casos tienda–SKU cubiertos al 100% dividido entre casos con "
                    "recomendación. Mide cumplimiento completo, no tareas generadas."
                ),
                "tone": "acid",
            },
            {
                "category": "PORCENTAJE · CASOS",
                "label": "ATENCIÓN DE CASOS",
                "value": f"{summary['case_service_pct']:,.1f}%",
                "description": (
                    "Casos tienda–SKU con al menos una unidad asignada dividido entre "
                    "casos con recomendación. Incluye coberturas parciales."
                ),
            },
            {
                "category": "PORCENTAJE · UNIDADES",
                "label": "COMPLIANCE DE UNIDADES",
                "value": f"{summary['unit_compliance_pct']:,.1f}%",
                "description": (
                    "Unidades asignadas dividido entre el objetivo final del modelo, "
                    "incluidos los incrementos de hardcode."
                ),
                "tone": "blue",
            },
            {
                "category": "PORCENTAJE · UNIDADES",
                "label": "COMPLIANCE DEL ROQ ORIGINAL",
                "value": f"{summary['original_roq_compliance_pct']:,.1f}%",
                "description": (
                    "Unidades del ROQ original cubiertas dividido entre las unidades "
                    "del ROQ antes de aplicar mínimos y hardcodes."
                ),
                "tone": "blue",
            },
        ]
    )

    st.markdown('<span class="section-label">UNIDADES</span>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "UNIDADES · ORIGINALES",
                "label": "ROQ ORIGINAL",
                "value": f"{summary['original_roq_units']:,}",
                "description": (
                    "Suma del ROQ original positivo antes de aplicar el mínimo de tres "
                    "o cualquier hardcode."
                ),
            },
            {
                "category": "UNIDADES · OBJETIVO",
                "label": "OBJETIVO FINAL",
                "value": f"{summary['target_units']:,}",
                "description": (
                    "Unidades solicitadas finalmente por el modelo después de ROQ, "
                    "mínimos y hardcodes."
                ),
                "tone": "acid",
            },
            {
                "category": "UNIDADES · ASIGNADAS",
                "label": "PRODUCTO PLANEADO",
                "value": f"{summary['assigned_units']:,}",
                "description": (
                    "Unidades de producto normal que efectivamente fueron asignadas. "
                    "No incluye insumos."
                ),
                "tone": "blue",
            },
            {
                "category": "UNIDADES · INCREMENTALES",
                "label": "HARDCODE ENVIADO",
                "value": f"{summary['hardcode_assigned_units']:,}",
                "description": (
                    "Unidades realmente asignadas por encima del ROQ original debido "
                    "a las reglas de hardcode."
                ),
                "tone": "coral",
            },
        ]
    )

    st.markdown('<span class="section-label">REGLAS ESPECIALES</span>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "CASOS · HARDCODE",
                "label": "CASOS CON HARDCODE",
                "value": f"{summary['hardcode_cases']:,}",
                "description": (
                    "Casos donde la cantidad objetivo quedó por encima del ROQ "
                    "original por una regla de negocio."
                ),
            },
            {
                "category": "UNIDADES · HARDCODE",
                "label": "OBJETIVO AÑADIDO",
                "value": f"{summary['hardcode_target_units']:,}",
                "description": (
                    "Diferencia total entre el objetivo final y el ROQ original. "
                    "Representa lo solicitado adicionalmente por hardcodes."
                ),
                "tone": "coral",
            },
            {
                "category": "CASOS · FORECAST 0",
                "label": "FORZADOS A 4",
                "value": f"{summary['forecast_zero_forced_cases']:,}",
                "description": (
                    "Casos con forecast, inventario de apertura y ROQ en cero cuyo "
                    "objetivo fue forzado a cuatro unidades."
                ),
            },
            {
                "category": "CASOS · DÉFICIT",
                "label": "FORZADOS A 3",
                "value": f"{summary['deficit_forced_cases']:,}",
                "description": (
                    "Casos con inventario de apertura menor a la demanda y ROQ cero "
                    "cuyo objetivo fue forzado a tres unidades."
                ),
            },
        ]
    )

    st.markdown('<span class="section-label">STOCKOUTS + PRIORIDADES</span>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "CASOS · STOCKOUT",
                "label": "STOCKOUTS CON ENVÍO",
                "value": f"{summary['stockout_cases_served']:,}",
                "description": (
                    "Casos tienda–SKU que iniciaron en stockout y reciben al menos "
                    "una unidad."
                ),
                "tone": "acid",
            },
            {
                "category": "SKUs · STOCKOUT",
                "label": "PRODUCTOS DISTINTOS ATENDIDOS",
                "value": f"{summary['stockout_products_served']:,}",
                "description": (
                    "RETAIL_ID distintos en stockout atendidos en al menos una tienda. "
                    "No es un conteo de tareas."
                ),
            },
            {
                "category": "CASOS · STOCKOUT",
                "label": "CUBIERTOS AL 100%",
                "value": f"{summary['stockout_cases_full']:,}",
                "description": (
                    "Casos tienda–SKU en stockout cuya cantidad objetivo fue cubierta "
                    "completamente."
                ),
                "tone": "acid",
            },
            {
                "category": "CASOS · GOLDEN",
                "label": "GOLDEN CON ENVÍO",
                "value": f"{summary['golden_served_cases']:,}",
                "description": (
                    "Casos Golden Infaltables que reciben al menos una unidad, sin "
                    "importar cuántas tareas se generaron."
                ),
                "tone": "blue",
            },
        ]
    )

    st.markdown('<span class="section-label">REGLAS DE DEMANDA</span>', unsafe_allow_html=True)
    report_table(
        analytics["rule_rows"],
        column_config={
            "COMPLIANCE_UNIDADES_%": st.column_config.NumberColumn(
                "COMPLIANCE UNIDADES %", format="%.1f%%"
            ),
        },
        max_height=360,
    )

    st.markdown('<span class="section-label">STOCKOUTS POR CIUDAD</span>', unsafe_allow_html=True)
    report_table(
        analytics["stockout_city_rows"],
        column_config={
            "ATENCION_%": st.column_config.NumberColumn(
                "ATENCIÓN %", format="%.1f%%"
            ),
        },
        max_height=400,
    )

    st.markdown('<span class="section-label">ASIGNACIÓN POR ORIGEN</span>', unsafe_allow_html=True)
    report_table(
        analytics["source_rows"],
        column_config={
            "WAREHOUSE_SOURCE": st.column_config.NumberColumn(
                "WAREHOUSE SOURCE", format="%d"
            ),
            "M3": st.column_config.NumberColumn("M³", format="%.3f"),
            "UNIDADES_POR_TAREA": st.column_config.NumberColumn(
                "UNIDADES / TAREA", format="%.2f"
            ),
            "PARTICIPACION_TAREAS_%": st.column_config.NumberColumn(
                "% DE TAREAS", format="%.1f%%"
            ),
            "PARTICIPACION_UNIDADES_%": st.column_config.NumberColumn(
                "% DE UNIDADES", format="%.1f%%"
            ),
        },
        max_height=360,
    )

    render_source_analysis(analytics)


def render_results(run: dict[str, Any]) -> None:
    st.markdown('<div class="result-title">PLANEACIÓN LISTA.</div>', unsafe_allow_html=True)
    render_kpi_cards(
        [
            {
                "category": "CASOS · INPUT",
                "label": "REQUERIMIENTOS ÚNICOS RECIBIDOS",
                "value": f"{run.get('input_requirements', run['requirements']):,}",
                "description": (
                    "Combinaciones tienda–SKU únicas leídas del CSV después de "
                    "consolidar duplicados. Incluye tiendas cerradas y ciudades "
                    "bloqueadas antes de aplicar exclusiones."
                ),
                "tone": "acid",
            },
            {
                "category": "CASOS · MODELO",
                "label": "REQUERIMIENTOS EVALUADOS",
                "value": f"{run['requirements']:,}",
                "description": (
                    "Casos tienda–SKU que llegaron al motor después de quitar "
                    "TIENDAS_CERRADAS y los bloqueos opcionales de ciudad."
                ),
            },
            {
                "category": "TAREAS · OPERACIÓN",
                "label": "TAREAS DE ABASTO",
                "value": f"{run['tasks']:,}",
                "description": (
                    "Líneas operativas de transferencia generadas por el modelo. "
                    "Un caso puede crear dos tareas si se divide entre dos orígenes. "
                    "Las líneas de insumos no cuentan aquí."
                ),
                "tone": "blue",
            },
            {
                "category": "UNIDADES · ABASTO",
                "label": "UNIDADES DE PRODUCTO",
                "value": f"{run['units']:,}",
                "description": (
                    "Suma de QUANTITY de las transferencias normales de producto. "
                    "No incluye insumos."
                ),
            },
            {
                "category": "LÍNEAS · INSUMOS",
                "label": "LÍNEAS DE INSUMOS",
                "value": f"{run.get('insumos', {}).get('lines_added', 0):,}",
                "description": (
                    "Filas de INSUMOS anexadas al BulkCD_444 para tiendas que ya "
                    "reciben producto normal desde 444. No consumen tareas."
                ),
            },
            {
                "category": "UNIDADES · INSUMOS",
                "label": "UNIDADES DE INSUMOS",
                "value": f"{run.get('insumos', {}).get('units_added', 0):,}",
                "description": (
                    "Suma de QUANTITY de las líneas de insumos agregadas. Se muestra "
                    "separada para no mezclarla con unidades de producto."
                ),
                "tone": "coral",
            },
        ],
        columns_count=3,
    )

    insumos = run.get("insumos", {})
    if insumos.get("lines_added", 0) > 0:
        st.success(
            f"Insumos anexados a BulkCD_444: {insumos['lines_added']:,} líneas, "
            f"{insumos['units_added']:,} unidades y "
            f"{insumos['stores_added']:,} tiendas. No consumen tareas del modelo."
        )

    closed_stores = run.get("closed_stores", {})
    if closed_stores.get("requirements", 0) > 0:
        store_ids = ", ".join(map(str, closed_stores.get("store_ids", [])))
        st.warning(
            "Bloqueo backend TIENDAS_CERRADAS aplicado: se excluyeron "
            f"{closed_stores['requirements']:,} requerimientos de "
            f"{closed_stores['stores']:,} tiendas ({store_ids}) antes de asignar stock."
        )

    city_block = run.get("city_block", {})
    if city_block.get("requirements", 0) > 0:
        city_names = ", ".join(
            item["name"] for item in city_block.get("cities", [])
        )
        st.warning(
            f"Bloqueo de ciudad aplicado: {city_names}. Se excluyeron "
            f"{city_block['requirements']:,} requerimientos, "
            f"{city_block['stores']:,} tiendas y "
            f"{city_block['products']:,} productos antes de asignar stock."
        )

    st.markdown('<span class="section-label">DESCARGAR TODO</span>', unsafe_allow_html=True)
    zip_path = Path(run["zip"])
    st.download_button(
        "Descargar planeación completa (.zip)",
        data=zip_path.read_bytes(),
        file_name=zip_path.name,
        mime="application/zip",
        use_container_width=True,
    )

    st.markdown('<span class="section-label">ARCHIVOS INDIVIDUALES</span>', unsafe_allow_html=True)
    columns = st.columns(min(max(len(run["files"]), 1), 3))
    for index, raw_path in enumerate(run["files"]):
        path = Path(raw_path)
        mime = (
            "text/csv"
            if path.suffix.lower() == ".csv"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with columns[index % len(columns)]:
            st.markdown(f'<span class="file-pill">{path.name}</span>', unsafe_allow_html=True)
            st.download_button(
                f"Descargar {path.name}",
                data=path.read_bytes(),
                file_name=path.name,
                mime=mime,
                key=f"download_{index}_{path.name}",
                use_container_width=True,
            )

    st.markdown('<span class="section-label">BREAKDOWN</span>', unsafe_allow_html=True)
    breakdown = [
        {"TIPO_DE_CORTE": status, "FILAS": count}
        for status, count in sorted(run["status_counts"].items())
    ]
    st.dataframe(breakdown, use_container_width=True, hide_index=True)

    if run.get("analytics"):
        render_planning_analytics(run["analytics"])

    if run["warnings"]:
        with st.expander(f"Advertencias de calidad ({len(run['warnings'])})"):
            for warning in run["warnings"]:
                st.warning(warning)
    with st.expander("Log técnico de la ejecución"):
        st.code(run["logs"] or "Ejecución completada sin mensajes adicionales.")


def main() -> None:
    inject_styles()

    st.markdown(
        """
        <section class="hero">
            <span class="hero-kicker">ABASTO / MX / WEB</span>
            <h1>TRANSFER<br>PLANNER.</h1>
            <p>Sube el requerimiento diario, define los orígenes y ejecuta el mismo motor sin editar código. La base se consulta y valida automáticamente.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<span class="section-label">01 — BASE DE DATOS</span>', unsafe_allow_html=True)
    database_bytes: bytes | None = None
    database_health: dict[str, Any] | None = None
    city_labels: dict[str, str] = {}
    try:
        with st.spinner("Validando fuentes de información…"):
            database_bytes = fetch_public_database()
            database_health = inspect_database(database_bytes)
            city_labels = extract_available_cities(database_bytes)
        render_database_health(database_health)
    except Exception as exc:
        st.markdown(
            """
            <div class="database-status review">
                <div class="database-title">BASE DE DATOS — OFFLINE</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.error(f"No fue posible validar la base de datos: {exc}")

    if st.button("VOLVER A VALIDAR LA BASE", use_container_width=False):
        fetch_public_database.clear()
        st.rerun()

    st.markdown('<span class="section-label">02 — ARCHIVO DE PLANEACIÓN</span>', unsafe_allow_html=True)
    uploaded_plan = st.file_uploader(
        "Plan diario (.csv)",
        type=["csv"],
        max_upload_size=MAX_UPLOAD_MB,
        help="El archivo debe conservar las columnas originales del plan.",
    )
    if uploaded_plan is not None:
        st.success(
            f"{uploaded_plan.name} · {uploaded_plan.size / (1024 ** 2):,.1f} MB"
        )

    st.markdown('<span class="section-label">03 — VARIABLES</span>', unsafe_allow_html=True)
    run_date = datetime.now(ZoneInfo("America/Mexico_City")).date()
    default_origins = [
        origin
        for origin in engine.CONFIG.origin_warehouses
        if origin in ORIGIN_WAREHOUSES
    ] or [811, 834]

    with st.form("planning_config"):
        left, right = st.columns([2, 1])
        with left:
            selected_origins = st.multiselect(
                "Warehouses origen — selecciona en orden de prioridad",
                options=list(ORIGIN_WAREHOUSES),
                default=default_origins,
                format_func=format_origin,
                help=(
                    "El motor consumirá stock en el orden seleccionado. Para cambiar "
                    "la prioridad, elimina las opciones y vuelve a elegirlas."
                ),
            )
        with right:
            max_tasks = st.number_input(
                "Máximo de tareas",
                min_value=0,
                max_value=1_000_000,
                value=int(engine.CONFIG.max_tasks),
                step=500,
            )

        selected_blocked_cities = st.multiselect(
            "Bloquear ciudades completas — opcional",
            options=list(city_labels),
            default=[],
            format_func=lambda city: city_labels.get(city, city),
            help=(
                "No se generarán envíos hacia las ciudades seleccionadas. Sus "
                "requerimientos se eliminan antes de asignar, por lo que ese stock "
                "queda disponible para otras tiendas."
            ),
        )
        if selected_blocked_cities:
            selected_names = ", ".join(
                city_labels.get(city, city) for city in selected_blocked_cities
            )
            st.warning(f"Se bloqueará completamente: {selected_names}")

        include_insumos = st.toggle(
            "Agregar insumos al BulkCD_444",
            value=True,
            help=(
                "Cuando está activo, anexa los insumos de Aleph únicamente a las "
                "tiendas que ya reciben producto normal desde el warehouse 444. "
                "Los insumos no consumen tareas, stock ni capacidad."
            ),
        )

        submitted = st.form_submit_button(
            "EJECUTAR PLANEACIÓN →", use_container_width=True
        )

    if submitted:
        if uploaded_plan is None:
            st.error("Primero sube el CSV del plan diario.")
            st.stop()
        if not selected_origins:
            st.error("Selecciona al menos un warehouse origen.")
            st.stop()
        try:
            origins = tuple(selected_origins)
            with st.status("Ejecutando motor de planeación…", expanded=True) as status:
                st.write("Guardando el CSV cargado de forma temporal…")
                st.write("Validando la versión actual de la base de datos…")
                fetch_public_database.clear()
                database_bytes = fetch_public_database()
                database_health = inspect_database(database_bytes)
                if not database_health["online"]:
                    raise RuntimeError(
                        "La base de datos tiene una o más fuentes con error. "
                        "Abre el panel de estado y corrige las hojas indicadas."
                    )
                if selected_blocked_cities:
                    blocked_names = ", ".join(
                        city_labels.get(city, city)
                        for city in selected_blocked_cities
                    )
                    st.write(f"Excluyendo ciudades bloqueadas: {blocked_names}…")
                if include_insumos and 444 in origins:
                    st.write("Preparando insumos para las rutas activas del 444…")
                elif include_insumos:
                    st.write(
                        "Insumos activos, pero sin efecto porque el origen 444 no "
                        "está seleccionado."
                    )
                else:
                    st.write("Envío de insumos desactivado para esta corrida.")
                st.write("Calculando demanda, stock, capacidad y tareas…")
                run = execute_planning(
                    uploaded_plan=uploaded_plan,
                    database_bytes=database_bytes,
                    origins=origins,
                    max_tasks=int(max_tasks),
                    run_date=run_date,
                    blocked_cities=tuple(selected_blocked_cities),
                    include_insumos=include_insumos,
                )
                status.update(label="Planeación finalizada", state="complete", expanded=False)
            st.session_state["last_run"] = run
        except Exception as exc:
            st.session_state.pop("last_run", None)
            st.error(f"No se pudo completar la planeación: {exc}")
            st.info(
                "Revisa el CSV, el panel de estado y que los warehouses origen existan en TIENDA."
            )

    last_run = st.session_state.get("last_run")
    if last_run and Path(last_run["zip"]).exists():
        render_results(last_run)


if __name__ == "__main__":
    main()
