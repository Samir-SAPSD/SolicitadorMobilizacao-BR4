"""Compatibility layer for services - re-exports from specialized modules."""

import json
import os

from backend.config import (
    ALLOWED_EXTENSIONS,
    REPORTS_FOLDER,
    TEMPLATE_FILENAME,
    TEMPLATE_STATUS_FILE,
    UPLOAD_FOLDER,
)
from backend.correction_helper import (
    allowed_file,
    extract_lista_suspensa_columns,
    normalize_key,
)
from backend.encoding_repair import (
    decode_powershell_output,
    has_mojibake_markers,
    repair_mojibake,
    repair_mojibake_once,
    score_decoded_text,
)
from backend.excel_processor import (
    count_excel_rows as _count_excel_rows,
    create_group_workbook as _create_group_workbook,
    read_grouped_excel as _read_grouped_excel,
    row_has_any_data as _row_has_any_data,
    value_to_exact_text as _value_to_exact_text,
)
from backend.powershell_executor import (
    build_powershell_command,
    build_powershell_populate_command,
    build_powershell_template_update_command,
    popen_hidden_kwargs as _popen_hidden_kwargs,
)
from backend.report_generator import generate_report_excel


def get_template_status():
    if not os.path.exists(TEMPLATE_STATUS_FILE):
        return {"last_updated": None}

    try:
        with open(TEMPLATE_STATUS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return {"last_updated": data.get("last_updated")}
    except Exception:
        return {"last_updated": None}


def save_template_status(iso_datetime):
    payload = {"last_updated": iso_datetime}
    with open(TEMPLATE_STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
