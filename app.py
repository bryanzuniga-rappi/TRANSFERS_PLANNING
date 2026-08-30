from __future__ import annotations

from collections import Counter
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo
import io
import re
import shutil
import tempfile
import urllib.error
import urllib.request
import zipfile

import streamlit as st

import modelo_abasto as engine


APP_NAME = "Transfer Planner"
MAX_UPLOAD_MB = 500
DATA_TRANSFERS_SPREADSHEET_ID = "18kHevkMvf9l4s6ANg3h5KdNyj2yEPGAp5C_t8JwxFVw"


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


def parse_origins(value: str) -> tuple[int, ...]:
    tokens = [token for token in re.split(r"[\s,;|]+", value.strip()) if token]
    if not tokens:
        raise ValueError("Ingresa al menos un warehouse origen")
    try:
        origins = tuple(int(token) for token in tokens)
    except ValueError as exc:
        raise ValueError("Los orígenes deben ser IDs enteros separados por comas") from exc
    if any(origin <= 0 for origin in origins):
        raise ValueError("Los IDs de origen deben ser positivos")
    if len(origins) != len(set(origins)):
        raise ValueError("La lista de orígenes contiene duplicados")
    return origins


def save_uploaded_file(uploaded_file, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    uploaded_file.seek(0)
    with destination.open("wb") as handle:
        shutil.copyfileobj(uploaded_file, handle, length=8 * 1024 * 1024)
    uploaded_file.seek(0)


def download_public_data_transfers(destination: Path) -> None:
    """Exporta el Google Sheet público completo como XLSX."""
    destination.parent.mkdir(parents=True, exist_ok=True)
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
            with destination.open("wb") as handle:
                shutil.copyfileobj(response, handle, length=8 * 1024 * 1024)
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
        raise RuntimeError(
            "No pude descargar DATA_TRANSFERS. Confirma que el Google Sheet "
            "siga configurado como 'Cualquier persona con el enlace: Lector'."
        ) from exc

    if not destination.exists() or destination.stat().st_size == 0:
        raise RuntimeError("Google devolvió un DATA_TRANSFERS vacío.")
    if not zipfile.is_zipfile(destination):
        destination.unlink(missing_ok=True)
        raise RuntimeError(
            "Google no devolvió un archivo Excel. Revisa que DATA_TRANSFERS "
            "sea público y que el ID configurado sea correcto."
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
    origins: tuple[int, ...],
    max_tasks: int,
    run_date,
    default_capacity: float,
    default_m3: float,
    minimum_quantity: int,
) -> dict[str, Any]:
    clear_previous_workspace()
    workspace = Path(tempfile.mkdtemp(prefix="transfer_planner_"))
    st.session_state["last_workspace"] = str(workspace)
    plan_path = workspace / "input" / Path(uploaded_plan.name).name
    data_path = workspace / "input" / "DATA_TRANSFERS.xlsx"

    save_uploaded_file(uploaded_plan, plan_path)
    download_public_data_transfers(data_path)

    config = engine.Config(
        origin_warehouses=origins,
        max_tasks=max_tasks,
        run_date_override=run_date.strftime("%d-%m-%Y"),
        default_store_capacity_m3=default_capacity,
        default_m3_per_unit=default_m3,
        minimum_positive_quantity=minimum_quantity,
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
            <p>Sube el requerimiento diario, define los orígenes y ejecuta el mismo motor sin editar código. DATA_TRANSFERS se consulta automáticamente.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<span class="section-label">01 — ARCHIVOS</span>', unsafe_allow_html=True)
    st.markdown(
        '<div class="info-strip">DATA_TRANSFERS conectado automáticamente · Solo carga el CSV diario · Los resultados se procesan temporalmente.</div>',
        unsafe_allow_html=True,
    )
    uploaded_plan = st.file_uploader(
        "Plan diario (.csv)",
        type=["csv"],
        max_upload_size=MAX_UPLOAD_MB,
        help="El archivo debe conservar las columnas originales del plan.",
    )
    st.success("DATA_TRANSFERS · CONECTADO")
    if uploaded_plan is not None:
        st.success(
            f"{uploaded_plan.name} · {uploaded_plan.size / (1024 ** 2):,.1f} MB"
        )

    st.markdown('<span class="section-label">02 — VARIABLES</span>', unsafe_allow_html=True)
    today_mx = datetime.now(ZoneInfo("America/Mexico_City")).date()

    with st.form("planning_config"):
        left, middle, right = st.columns([1.4, 1, 1])
        with left:
            origins_text = st.text_input(
                "Warehouses origen — en orden de prioridad",
                value=", ".join(map(str, engine.CONFIG.origin_warehouses)),
                help="El motor consumirá stock de izquierda a derecha.",
            )
        with middle:
            max_tasks = st.number_input(
                "Máximo de tareas",
                min_value=0,
                max_value=1_000_000,
                value=int(engine.CONFIG.max_tasks),
                step=500,
            )
        with right:
            run_date = st.date_input("Fecha de ejecución", value=today_mx)

        with st.expander("Configuración avanzada"):
            adv1, adv2, adv3 = st.columns(3)
            with adv1:
                default_capacity = st.number_input(
                    "Capacidad default por tienda (m³)",
                    min_value=0.0,
                    value=float(engine.CONFIG.default_store_capacity_m3),
                    step=1.0,
                )
            with adv2:
                default_m3 = st.number_input(
                    "Volumen default por unidad (m³)",
                    min_value=0.0,
                    value=float(engine.CONFIG.default_m3_per_unit),
                    step=0.001,
                    format="%.6f",
                )
            with adv3:
                minimum_quantity = st.number_input(
                    "Cantidad mínima positiva",
                    min_value=1,
                    value=int(engine.CONFIG.minimum_positive_quantity),
                    step=1,
                )

        submitted = st.form_submit_button(
            "EJECUTAR PLANEACIÓN →", use_container_width=True
        )

    if submitted:
        if uploaded_plan is None:
            st.error("Primero sube el CSV del plan diario.")
            st.stop()
        try:
            origins = parse_origins(origins_text)
            with st.status("Ejecutando motor de planeación…", expanded=True) as status:
                st.write("Guardando el CSV cargado de forma temporal…")
                st.write("Descargando la versión actual de DATA_TRANSFERS…")
                st.write("Calculando demanda, stock, capacidad y tareas…")
                run = execute_planning(
                    uploaded_plan=uploaded_plan,
                    origins=origins,
                    max_tasks=int(max_tasks),
                    run_date=run_date,
                    default_capacity=float(default_capacity),
                    default_m3=float(default_m3),
                    minimum_quantity=int(minimum_quantity),
                )
                status.update(label="Planeación finalizada", state="complete", expanded=False)
            st.session_state["last_run"] = run
        except Exception as exc:
            st.session_state.pop("last_run", None)
            st.error(f"No se pudo completar la planeación: {exc}")
            st.info(
                "Revisa el CSV, el acceso público de DATA_TRANSFERS y que los warehouses origen existan en TIENDA."
            )

    last_run = st.session_state.get("last_run")
    if last_run and Path(last_run["zip"]).exists():
        render_results(last_run)


if __name__ == "__main__":
    main()
