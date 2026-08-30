from datetime import date
from pathlib import Path
import io
import tempfile

import app


def test_parse_origins_preserves_priority_order():
    assert app.parse_origins("811, 834 444") == (811, 834, 444)


def test_parse_origins_rejects_duplicates():
    try:
        app.parse_origins("811,811")
    except ValueError as exc:
        assert "duplicados" in str(exc)
    else:
        raise AssertionError("Debió rechazar orígenes duplicados")


def test_uploaded_file_is_copied_in_chunks():
    payload = io.BytesIO(b"warehouse,sku\n105,10087\n")
    payload.name = "plan.csv"
    with tempfile.TemporaryDirectory() as directory:
        destination = Path(directory) / "plan.csv"
        app.save_uploaded_file(payload, destination)
        assert destination.read_bytes() == b"warehouse,sku\n105,10087\n"


def test_config_date_shape():
    assert date(2026, 8, 30).strftime("%d-%m-%Y") == "30-08-2026"
