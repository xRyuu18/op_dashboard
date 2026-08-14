from __future__ import annotations

import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .config import Config, load_config, save_config
from .exporter import AmbiguousProjectName, ProjectNotFound, run_export
from .importer import run_import
from .openproject_client import OpenProjectError

BASE_DIR = Path(__file__).resolve().parent
EXPORTS_DIR = Path(tempfile.gettempdir()) / "op_dashboard_exports"
EXPORTS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="OpenProject Import/Export Dashboard")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def _cfg_ready() -> bool:
    cfg = load_config()
    return bool(cfg.base_url and cfg.api_key)


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "cfg_ready": _cfg_ready()})


@app.get("/settings", response_class=HTMLResponse)
def settings_form(request: Request):
    cfg = load_config()
    return templates.TemplateResponse("settings.html", {"request": request, "cfg": cfg})


@app.post("/settings")
def settings_save(base_url: str = Form(...), api_key: str = Form(...)):
    save_config(Config(base_url=base_url, api_key=api_key))
    return RedirectResponse("/", status_code=303)


@app.get("/import", response_class=HTMLResponse)
def import_form(request: Request):
    if not _cfg_ready():
        return RedirectResponse("/settings", status_code=303)
    return templates.TemplateResponse("import.html", {"request": request})


@app.post("/import", response_class=HTMLResponse)
def import_submit(request: Request, project_name: str = Form(...), task_file: UploadFile = File(...)):
    cfg = load_config()
    suffix = Path(task_file.filename or "upload.xlsx").suffix or ".xlsx"
    tmp_path = Path(tempfile.gettempdir()) / f"op_import_{uuid.uuid4().hex}{suffix}"
    with tmp_path.open("wb") as f:
        shutil.copyfileobj(task_file.file, f)

    error_message = None
    result = None
    try:
        result = run_import(cfg.base_url, cfg.api_key, project_name, str(tmp_path))
    except (OpenProjectError, RuntimeError) as e:
        error_message = str(e)
    finally:
        tmp_path.unlink(missing_ok=True)

    return templates.TemplateResponse(
        "import_result.html",
        {"request": request, "result": result, "error_message": error_message},
    )


@app.get("/export", response_class=HTMLResponse)
def export_form(request: Request):
    if not _cfg_ready():
        return RedirectResponse("/settings", status_code=303)
    return templates.TemplateResponse("export.html", {"request": request})


@app.post("/export", response_class=HTMLResponse)
def export_submit(request: Request, project_name: str = Form(...)):
    cfg = load_config()
    out_name = f"Export_{project_name}_{uuid.uuid4().hex[:8]}.xlsx"
    out_path = EXPORTS_DIR / out_name

    error_message = None
    result = None
    try:
        result = run_export(cfg.base_url, cfg.api_key, project_name, str(out_path))
    except ProjectNotFound as e:
        error_message = str(e)
    except AmbiguousProjectName as e:
        error_message = str(e)
    except (OpenProjectError, RuntimeError) as e:
        error_message = str(e)

    download_name = out_name if (result and result.file_path) else None

    return templates.TemplateResponse(
        "export_result.html",
        {"request": request, "result": result, "error_message": error_message, "download_name": download_name},
    )


@app.get("/export/download/{filename}")
def export_download(filename: str):
    file_path = EXPORTS_DIR / filename
    if not file_path.exists():
        return HTMLResponse("File tidak ditemukan (mungkin sudah dibersihkan).", status_code=404)
    return FileResponse(
        str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
