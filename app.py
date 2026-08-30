from __future__ import annotations

from collections import Counter
from contextlib import redirect_stdout
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo
import html
import io
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
    "GOLDEN_INFALTABLES",
    "TIENDA",
    "STORAGE",
}

REQUIRED_DATABASE_SHEETS = (
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
    "GOLDEN_INFALTABLES",
    "TIENDA",
    "STORAGE",
)

SHEET_DESCRIPTIONS = {
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
                        "TIPO": "ALEPH" if sheet_name in ALEPH_SHEETS else "IMPORTRANGE",
                        "CONTROL": "C7" if sheet_name in ALEPH_SHEETS else "A1",
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
                    "TIPO": "ALEPH" if sheet_name in ALEPH_SHEETS else "IMPORTRANGE",
                    "CONTROL": "C7" if sheet_name in ALEPH_SHEETS else "A1",
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
            card_meta = (
                "ÚLTIMA ACTUALIZACIÓN"
                if row["TIPO"] == "ALEPH"
                else "CONEXIÓN IMPORTRANGE"
            )
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


def execute_planning(
    uploaded_plan,
    database_bytes: bytes,
    origins: tuple[int, ...],
    max_tasks: int,
    run_date,
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
        response = engine.run_pipeline(
            config,
            local_data_transfers_path=data_path,
            local_plan_path=plan_path,
            upload_to_drive=False,
        )

    result = response["result"]
    local_files = [Path(path) for path in response["local_files"]]
    zip_path = create_zip(
        local_files,
        workspace / f"Planeacion_{run_date:%d-%m-%Y}.zip",
    )
    status_counts = dict(Counter(row["TIPO_DE_CORTE"] for row in result.base_rows))
    units = sum(row["QUANTITY"] for row in result.allocation_rows)
    requirements = len(result.base_rows)
    return {
        "workspace": str(workspace),
        "files": [str(path) for path in local_files],
        "zip": str(zip_path),
        "tasks": result.tasks_used,
        "units": units,
        "requirements": requirements,
        "status_counts": status_counts,
        "warnings": list(result.warnings),
        "logs": captured.getvalue(),
        "origins": list(origins),
    }


def render_results(run: dict[str, Any]) -> None:
    st.markdown('<div class="result-title">PLANEACIÓN LISTA.</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    col1.metric("REQUERIMIENTOS", f"{run['requirements']:,}")
    col2.metric("TAREAS", f"{run['tasks']:,}")
    col3.metric("UNIDADES", f"{run['units']:,}")

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
    try:
        with st.spinner("Validando fuentes de información…"):
            database_bytes = fetch_public_database()
            database_health = inspect_database(database_bytes)
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
                st.write("Calculando demanda, stock, capacidad y tareas…")
                run = execute_planning(
                    uploaded_plan=uploaded_plan,
                    database_bytes=database_bytes,
                    origins=origins,
                    max_tasks=int(max_tasks),
                    run_date=run_date,
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
