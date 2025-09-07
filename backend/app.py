# Flask/FastAPI
# Запуск: python -m uvicorn backend.app:app --reload --port 8000
# Flask/FastAPI
# Запуск: python -m uvicorn backend.app:app --reload --port 8000
from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import FileResponse
from pathlib import Path
import uuid
from backend.analysis import run_analysis

from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi import Request


app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
RESULTS_DIR = STATIC_DIR / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)
ERROR_LOG = BASE_DIR / "error.log"


app.mount("/static", StaticFiles(directory="backend/static"), name="static")

@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = STATIC_DIR / "index.html"
    # index_path = Path("backend/static/index.html")
    return index_path.read_text(encoding="utf-8")


@app.post("/analyze")
def analyze(background_tasks: BackgroundTasks):
    """Запускает анализ в фоне и возвращает task_id.
    ВАЖНО: run_analysis ДОЛЖНА писать в out_file, который мы передаём сюда!"""
    task_id = str(uuid.uuid4())
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"
    background_tasks.add_task(run_analysis, out_file)
    return {"task_id": task_id}

@app.get("/status/{task_id}")
def status(task_id: str):
    """Совместимо с твоим фронтом: возвращает {"ready": bool}.
    Если файла нет, дополнительно шлёт хвост error.log (если он есть)."""
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"
    if out_file.exists():
        return {"ready": True, "download_url": f"/download/{task_id}"}
    # return {"ready": out_file.exists()}
    # файл ещё не готов — попробуем отдать хвост лога (на случай падения задачи)
    error_tail = None
    if ERROR_LOG.exists():
        try:
            lines = ERROR_LOG.read_text(encoding="utf-8").splitlines()
            error_tail = "\n".join(lines[-20:]) if lines else None
        except Exception:
            error_tail = None

    return {"ready": False, "error": error_tail}

@app.get("/download/{task_id}")
def download(task_id: str):
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"
    if out_file.exists():
     return FileResponse(out_file, filename=out_file.name)
    return {"error": "Файл ещё не готов или не найден"}


