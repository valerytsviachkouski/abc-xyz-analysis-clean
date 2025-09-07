#  код (обёрнут в функцию run_analysis)
import os
import traceback
import pandas as pd
import matplotlib

# отключаем GUI-бэкенд для Matplotlib (важно для серверов без экрана)
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from pathlib import Path
import json
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


os.environ['TCL_LIBRARY'] = r'C:\Program Files\Python313\tcl\tcl8.6'


def log_message(msg: str):
    """
    Записывает сообщение в error.log (используется и для ошибок, и для прогресса).
    """
    log_file = Path(__file__).resolve().parent / "error.log"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[INFO] {msg}\n")


def run_analysis(out_file: Path):
    try:
        log_message("=== Запуск анализа ===")

        # === 1. Загружаем конфиг ===
        BASE_DIR = Path(__file__).resolve().parent.parent
        CONFIG_PATH = BASE_DIR / "config.json"

        with open(CONFIG_PATH, encoding="utf-8") as f:
            config = json.load(f)

        input_file = Path(config["input_file"])
        abc_file = Path(config["abc_file"])
        out_dir = Path(config["output_dir"])
        xyz_thresholds = config["xyz_thresholds"]
        period_days = config["period_days"]

        log_message("Конфиг загружен")

        # Пути для промежуточных файлов
        out_path_w = out_dir / "Исходная таблица_w.xlsx"
        out_path_ship = out_dir / "Исходная таблица_отгрузка.xlsx"
        out_path_stock = out_dir / "Исходная таблица_остаток.xlsx"

        # === 2. Трансформация исходной таблицы ===
        df = pd.read_excel(input_file, header=None)
        log_message("Исходная таблица загружена")

        df.iloc[3, 1:] = df.iloc[3, 1:].replace("отгрузка продукта", "отгрузка")
        for i in range(4, len(df), 3):
            if i < len(df):
                df.iloc[i + 2, 0] = df.iloc[i, 0]
        df = df.drop(df.index[2::3]).reset_index(drop=True)

        df.to_excel(out_path_w, header=False, index=False)
        log_message("Таблица трансформирована и сохранена (W)")

        # === 3. Создание таблиц "отгрузка" и "остаток" ===
        df_ship = df.drop(df.index[1::2]).reset_index(drop=True)
        df_stock = df.drop(df.index[2::2]).reset_index(drop=True)

        def add_total_column(df_in, label: str):
            df_out = df_in.copy()
            df_num = df_out.iloc[:, 1:].replace("-", 0)
            df_num = pd.to_numeric(df_num.stack(), errors="coerce").unstack(fill_value=0)
            df_out["Всего"] = df_num.sum(axis=1).astype("object")

            if df_out.shape[0] > 0:
                df_out.at[0, "Всего"] = "Всего"
            if df_out.shape[0] > 1:
                df_out.at[1, "Всего"] = label
            return df_out

        df_ship = add_total_column(df_ship, "отгрузка")
        df_stock = add_total_column(df_stock, "остаток")

        df_stock_num = df_stock.iloc[:, 1:-1].replace("-", 0)
        df_stock_num = pd.to_numeric(df_stock_num.stack(), errors="coerce").unstack(fill_value=0)
        avg_values = df_stock_num.mean(axis=1).tolist()
        df_stock.insert(df_stock.shape[1], "Средний", avg_values)

        if df_stock.shape[0] > 0:
            df_stock.at[0, "Средний"] = "Средний"
        if df_stock.shape[0] > 1:
            df_stock.at[1, "Средний"] = "остаток"

        df_ship.to_excel(out_path_ship, header=False, index=False)
        df_stock.to_excel(out_path_stock, header=False, index=False)
        log_message("Файлы 'отгрузка' и 'остаток' сохранены")

        # === 4. ABC-XYZ-анализ ===
        ship = pd.read_excel(out_path_ship, header=0)
        stock = pd.read_excel(out_path_stock, header=0)
        abc = pd.read_excel(abc_file)

        ship = ship[pd.to_numeric(ship["Всего"], errors="coerce").notna()]
        stock = stock[pd.to_numeric(stock["Средний"], errors="coerce").notna()]

        ship["Всего"] = pd.to_numeric(ship["Всего"], errors="coerce")
        stock["Средний"] = pd.to_numeric(stock["Средний"], errors="coerce")

        df = pd.merge(
            ship[["Наименование", "Всего"]],
            stock[["Наименование", "Средний"]],
            on="Наименование",
            how="inner",
            suffixes=("_отгрузка", "_остаток")
        )

        df = pd.merge(df, abc[["Наименование", "Группа ABC"]], on="Наименование", how="left")

        df["Оборачиваемость_дни"] = df.apply(
            lambda x: (x["Средний"] * period_days) / x["Всего"]
            if pd.notna(x["Всего"]) and x["Всего"] > 0 else 9999,
            axis=1
        )

        def assign_xyz(turnover: float) -> str:
            if turnover <= xyz_thresholds["X"]:
                return "X"
            elif turnover <= xyz_thresholds["Y"]:
                return "Y"
            elif turnover <= xyz_thresholds["Z"]:
                return "Z"
            else:
                return "Неликвид"

        df["Группа XYZ"] = df["Оборачиваемость_дни"].apply(assign_xyz)
        df["ABC_XYZ"] = df["Группа ABC"] + "-" + df["Группа XYZ"]

        log_message("ABC-XYZ анализ рассчитан")

        # === 5. Сводная матрица и диаграмма ===
        pivot = pd.crosstab(df["Группа ABC"], df["Группа XYZ"])
        counts = df["ABC_XYZ"].value_counts().sort_index()

        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="ABC_XYZ_данные", index=False)
            pivot.to_excel(writer, sheet_name="Сводная матрица")

        log_message("Excel с данными и сводной матрицей сохранён")

        fig, ax = plt.subplots(figsize=(6, 6))
        ax.pie(
            counts,
            labels=counts.index,
            autopct="%1.1f%%",
            startangle=90,
            textprops={'fontsize': 9},
            radius=0.7
        )

        xyz_info = (
            f"X ≤ {xyz_thresholds['X']} дн., "
            f"Y ≤ {xyz_thresholds['Y']} дн., "
            f"Z ≤ {xyz_thresholds['Z']} дн."
        )

        ax.set_title(f"ABC-XYZ анализ (доли SKU)\nОборачиваемость: {xyz_info}", fontsize=11)
        ax.axis("equal")
        plt.tight_layout()

        chart_path = out_dir / "ABC_XYZ_pie.png"
        plt.savefig(chart_path, dpi=300)
        plt.close()
        log_message("Диаграмма сохранена как PNG")

        wb = load_workbook(out_file)
        ws_chart = wb.create_sheet("Диаграмма")
        img = Image(str(chart_path))
        img.width, img.height = 480, 480
        ws_chart.add_image(img, "B2")
        wb.save(out_file)

        log_message("Диаграмма добавлена в Excel")
        log_message("=== Анализ успешно завершён ===")

    except Exception as e:
        error_log = Path(__file__).resolve().parent / "error.log"
        with open(error_log, "a", encoding="utf-8") as f:
            f.write(f"\n[ERROR] Ошибка при выполнении анализа: {e}\n")
            f.write(traceback.format_exc() + "\n")




