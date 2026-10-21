import os
import warnings
import time
import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm
import app_settings_config as cfg
import analytics_engine as eng
import excel_formatter as fmt

warnings.filterwarnings("ignore")
os.environ["PYTHONWARNINGS"] = "ignore"
pd.set_option('future.no_silent_downcasting', True)


def main():
    # Запуск прогресс-бара с параметром leave=False (стирает шкалу после 100%)
    with tqdm(total=100, desc="Data processing", leave=False, bar_format="{desc}: |{bar}| {percentage:3.0f}%") as pbar:

        # Шаг 1: Математические расчеты (0% -> 25%)
        raw_metrics_df = eng.process_data(cfg.FILE_STOCK, cfg.FILE_SALES)
        time.sleep(0.1)
        pbar.update(25)

        # Шаг 2: Сегментация пулов (25% -> 45%)
        fin_A, fin_B, fin_C, hipo_df = eng.segment_into_groups(raw_metrics_df)
        time.sleep(0.1)
        pbar.update(20)

        # Шаг 3: Формирование структур листов (45% -> 65%)
        sheet_A, sheet_B, sheet_C = fmt.format_data_structures(
            fin_A, fin_B, fin_C, hipo_df)
        time.sleep(0.1)
        pbar.update(20)

        # Шаг 4: Запись в Excel-файл (65% -> 85%)
        try:
            with pd.ExcelWriter(cfg.FILE_OUTPUT, engine="openpyxl") as writer:
                sheet_A.to_excel(writer, sheet_name="Группа А", index=False)
                sheet_B.to_excel(writer, sheet_name="Группа B", index=False)
                sheet_C.to_excel(writer, sheet_name="Группа C", index=False)
            time.sleep(0.1)
            pbar.update(20)

            # Шаг 5: Наложение дизайна и шрифтов (85% -> 100%)
            wb = load_workbook(cfg.FILE_OUTPUT)
            fmt.apply_excel_styles(wb)
            wb.save(cfg.FILE_OUTPUT)
            time.sleep(0.1)
            pbar.update(15)

        except Exception as e:
            print(f"\n❌ Критическая ошибка при сборке: {e}")
            return

    # Вывод финального чистого статуса строго ПОСЛЕ закрытия и удаления прогресс-бара
    print(f"Отчет создан: '{cfg.FILE_OUTPUT}'")


if __name__ == "__main__":
    main()
