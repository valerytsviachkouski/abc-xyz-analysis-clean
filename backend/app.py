# Flask/FastAPI
# Запуск: python -m uvicorn backend.app:app --reload --port 8000
# Flask/FastAPI
# Запуск с Gitpod: python -m uvicorn backend.app:app --reload --host 0.0.0.0 --port 8000

# Запуск PyCharm Терминал uvicorn backend.app:app --reload

from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import FileResponse
from pathlib import Path
import uuid
from backend.analysis import run_analysis

from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi.responses import FileResponse
from fastapi import UploadFile, File
from fastapi import Request


app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
RESULTS_DIR = STATIC_DIR / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)
ERROR_LOG = BASE_DIR / "error.log"

UPLOAD_DIR = BASE_DIR / "data"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


app.mount("/static", StaticFiles(directory="backend/static"), name="static")

@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = STATIC_DIR / "index.html"
    # index_path = Path("backend/static/index.html")
    return index_path.read_text(encoding="utf-8")

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    file_path = UPLOAD_DIR / file.filename
    contents = await file.read()
    with open(file_path, "wb") as f:
        f.write(contents)
    return {"message": f"Файл {file.filename} загружен", "filename": file.filename}    


@app.post("/analyze")
async def analyze(request: Request, background_tasks: BackgroundTasks):
    body = await request.json()
    filename = body.get("filename")
    input_path = BASE_DIR / "data" / filename

    task_id = str(uuid.uuid4())
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"

    # ✅ Передаём оба аргумента
    background_tasks.add_task(run_analysis, out_file, input_path, task_id)

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


# скачиваем диаграмму на вэб-странице
@app.get("/chart/{task_id}")
def get_chart(task_id: str):
    chart_path = RESULTS_DIR / f"ABC_XYZ_pie.png"
    if chart_path.exists():
        return FileResponse(chart_path, media_type="image/png")
    return {"error": "Диаграмма не найдена"}

