"""Генерирует samples/input_structural.xlsx из samples/input_big.xlsx.

Структурный образец — это производственный образец (input_big.xlsx), обрезанный
до первых ROW_COUNT строк. Он сохраняет структуру и форматы ячеек (заголовки,
столбцы, формулы SUBTOTAL в строке 7, форматы данных), но содержит мало строк,
поэтому валидация входных файлов против него быстрая.

Объём данных (число строк) на валидацию не влияет: run_production.py
нормализует обе стороны к меньшему объёму перед сравнением. Поэтому образец
не обязан совпадать с входным файлом по числу строк — только по структуре и
форматам.

Пересоздайте этот файл, если изменилась структура производственных книг:

    python samples/make_structural_sample.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import transform  # noqa: E402
SOURCE = HERE / "input_big.xlsx"
OUTPUT = HERE / "input_structural.xlsx"
ROW_COUNT = 100


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"Не найден исходный образец: {SOURCE}")

    wb = load_workbook(SOURCE)
    if transform.SHEET_NAME not in wb.sheetnames:
        raise SystemExit(f"Лист '{transform.SHEET_NAME}' не найден в {SOURCE}.")
    ws = wb[transform.SHEET_NAME]

    current_last = transform.data_row(ws)
    if current_last > ROW_COUNT:
        ws.delete_rows(ROW_COUNT + 1, current_last - ROW_COUNT)

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(7, col)
        value = cell.value
        if isinstance(value, str) and value.startswith("="):
            match = transform._ROW7_SUBTOTAL_FULL_RE.match(value)
            if match:
                cell.value = (
                    f"=SUBTOTAL({match.group(1)},{match.group(2)}10:"
                    f"{match.group(3)}{ROW_COUNT + 1})"
                )

    wb.save(OUTPUT)
    print(f"Создан: {OUTPUT} ({ROW_COUNT} строк, из {SOURCE})")


if __name__ == "__main__":
    main()
