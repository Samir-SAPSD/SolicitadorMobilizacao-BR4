import os
import threading
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"
FRONTEND_TEMPLATES_DIR = FRONTEND_DIR / "templates"
FRONTEND_STATIC_DIR = FRONTEND_DIR / "static"

BACKEND_DIR = BASE_DIR / "backend"
DATA_DIR = BACKEND_DIR / "data"
SCRIPTS_DIR = BACKEND_DIR / "scripts"

UPLOAD_FOLDER = os.path.join(str(DATA_DIR), "templates")
REPORTS_FOLDER = os.path.join(str(DATA_DIR), "uploads")
TEMPLATE_FILENAME = "ModeloSolicitacaoMob.xlsx"
TEMPLATE_STATUS_FILE = os.path.join(UPLOAD_FOLDER, "template_update_status.json")
ALLOWED_EXTENSIONS = {"xlsx"}

for folder in (
    UPLOAD_FOLDER,
    REPORTS_FOLDER,
    str(FRONTEND_TEMPLATES_DIR),
    str(FRONTEND_STATIC_DIR),
    str(SCRIPTS_DIR),
):
    os.makedirs(folder, exist_ok=True)

ACTIVE_JOBS = 0
JOBS_LOCK = threading.Lock()


def start_job() -> None:
    global ACTIVE_JOBS
    with JOBS_LOCK:
        ACTIVE_JOBS += 1


def end_job() -> None:
    global ACTIVE_JOBS
    with JOBS_LOCK:
        if ACTIVE_JOBS > 0:
            ACTIVE_JOBS -= 1


def get_active_jobs() -> int:
    with JOBS_LOCK:
        return ACTIVE_JOBS
