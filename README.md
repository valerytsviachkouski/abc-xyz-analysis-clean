# ABC-XYZ Анализ Excel-файлов

Веб-сервис для автоматизированного анализа Excel-таблиц по методике ABC-XYZ. Загруженный файл обрабатывается на сервере, результаты сохраняются в виде отчёта с группировкой, фильтрацией и вычислениями.

---

## 🚀 Быстрый старт

### 🔧 Локальный запуск

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/valerytsviachkouski/abc-xyz-analysis-clean.git
   cd abc-xyz-analysis-clean
   
2.  Создайте виртуальное окружение и установите зависимости:
python -m venv .venv
source .venv/bin/activate  # или .venv\\Scripts\\activate для Windows
pip install -r requirements.txt

Запустите сервер:
uvicorn backend.app:app --reload

Деплой на Render
Build Command: pip install -r requirements.txt
Start Command: uvicorn backend.app:app --host 0.0.0.0 --port $PORT
Python Version: 3.11
Репозиторий: abc-xyz-analysis-clean

API Эндпоинты
Метод	URL	Описание
GET	/	HTML-страница с формой загрузки
POST	/analyze	Загружает Excel-файл и запускает анализ
GET	/status/{task_id}	Проверяет готовность результата
GET	/download/{task_id}	Скачивает готовый Excel-отчёт

Структура проекта
backend/
├── app.py              # FastAPI сервер
├── analysis.py         # Логика обработки Excel-файлов
├── static/
│   ├── index.html      # Веб-интерфейс
│   └── results/        # Папка для выходных файлов
├── data/
│   └── ABC_группы_январь_август.xlsx  # Исходные данные
├── config.json         # Настройки анализа

Зависимости
FastAPI
Uvicorn
Pandas
Openpyxl
Matplotlib
Aiofiles
Python-Multipart
Jinja2
Устанавливаются через requirements.txt.


