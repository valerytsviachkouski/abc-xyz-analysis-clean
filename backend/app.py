# Flask/FastAPI# Запуск PyCharm Терминал uvicorn backend.app:app --reload
# Запуск: python -m uvicorn backend.app:app --reload --port 8000
# Flask/FastAPI
# Запуск с Gitpod: python -m 2025-09-23T17:25:55.883732211Z ==> Running 'uvicorn backend.app:app --host 0.0.0.0 --port $PORT'


from fastapi import FastAPI, BackgroundTasks, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import uuid
import time
import aiofiles
import asyncio
import gc

from backend.analysis import run_analysis

app = FastAPI()

# Пути
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "data"                  # входные файлы
RESULTS_DIR = BASE_DIR / "static" / "results"   # выходные файлы

UPLOAD_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# Раздаём /static
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")

# ---------copilot---------------------------------------------------------
@app.middleware("http")
async def log_requests(request, call_next):
    print(f"📥 {request.method} {request.url}")
    response = await call_next(request)
    print(f"📤 {response.status_code}")
    return response
# --------------------copilot--------------------------------------------

# Главная страница
@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = BASE_DIR / "static" / "index.html"
    return index_path.read_text(encoding="utf-8")

# Очистка старых файлов
def cleanup_old_files():
    now = time.time()
    for file in RESULTS_DIR.glob("analysis_*.xlsx"):
        try:
            if now - file.stat().st_mtime > 86400:  # старше 1 суток
                file.unlink()
                print(f"🧹 Удалён файл: {file}")
        except Exception as e:
            print(f"Ошибка при удалении {file.name}: {e}")

# Анализ файла
@app.post("/analyze")
async def analyze_file(background_tasks: BackgroundTasks,
                       file: UploadFile = File(...)) -> JSONResponse:
    task_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"input_{task_id}.xlsx"
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"

    # -------------------copilot---------------------------------------
    async with aiofiles.open(input_path, "wb") as f:
        await f.write(await file.read())

    loop = asyncio.get_event_loop()
    loop.run_in_executor(None, run_analysis, out_file, input_path, task_id)
    # ----------------copilot----------------------------------------------

    # # Сохраняем загруженный файл
    # with open(input_path, "wb") as f:
    #     f.write(await file.read())
    #
    # # Запускаем анализ в фоне
    # background_tasks.add_task(run_analysis, out_file, input_path, task_id)

    # Запускаем очистку старых файлов в фоне
    background_tasks.add_task(cleanup_old_files)

    return {"task_id": task_id}


# Проверка готовности
@app.get("/status/{task_id}")
def status(task_id: str):
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"
    return {"ready": out_file.exists()}


# # Скачивание результата
@app.get("/download/{task_id}")
def download_file(task_id: str):
    out_file = RESULTS_DIR / f"analysis_{task_id}.xlsx"
    if out_file.exists():
        return FileResponse(out_file, filename=out_file.name)
    return {"error": "Файл ещё не готов или не найден"}

