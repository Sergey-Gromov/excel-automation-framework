"""Генерирует тестовые файлы с примерами всех отслеживаемых типов различий."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

OUT = Path(__file__).parent


def make_header(ws, row, cols):
    for col, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal="center")


DATA = [(1, "Apple", 10, 1.50), (2, "Banana", 20, 0.75), (3, "Cherry", 15, 3.00)]


# ── ОБРАЗЕЦ (sample.xlsx) ────────────────────────────────────────────────────
wb_s = Workbook()

ws_orders = wb_s.active
ws_orders.title = "Orders"

# Преамбула: 2 строки перед заголовками
ws_orders.cell(row=1, column=1, value="Отчёт по продажам")
ws_orders.cell(row=2, column=1, value="Дата: 01.07.2025")

# Строка заголовков — row 3, столбцы A–F (6 заголовков, первый в колонке A)
make_header(ws_orders, 3, ["ID", "Product", "Qty", "Price", "Total", "Tax"])

for r, (rid, prod, qty, price) in enumerate(DATA, 4):
    ws_orders.cell(row=r, column=1, value=rid)
    ws_orders.cell(row=r, column=2, value=prod)
    ws_orders.cell(row=r, column=3, value=qty)
    ws_orders.cell(row=r, column=4, value=price).number_format = "#,##0.00"
    c5 = ws_orders.cell(row=r, column=5, value=f"=C{r}*D{r}")
    c5.number_format = "#,##0.00"
    c6 = ws_orders.cell(row=r, column=6, value=f"=E{r}*0.2")
    c6.number_format = "0.00%"

# Объединение над финансовыми столбцами
ws_orders.cell(row=2, column=5, value="Финансовые показатели").font = Font(bold=True)
ws_orders.merge_cells("E2:F2")

ws_cfg = wb_s.create_sheet("Config")
ws_cfg["A1"] = "Version"
ws_cfg["B1"] = "1.0"

ws_sum = wb_s.create_sheet("Summary")
ws_sum["A1"] = "Итого"
ws_sum["B1"] = "=SUM(Orders!E4:E6)"
ws_sum["B1"].number_format = "#,##0.00"

wb_s.save(OUT / "sample.xlsx")
print("Создан: sample.xlsx")


# ── ВХОДНОЙ С ОТЛИЧИЯМИ (input_modified.xlsx) ────────────────────────────────
wb_i = Workbook()

ws_i = wb_i.active
ws_i.title = "Orders"

# [header_row_offset] преамбула из 3 строк вместо 2
ws_i.cell(row=1, column=1, value="Отчёт по продажам")
ws_i.cell(row=2, column=1, value="Период: Q3 2025")
ws_i.cell(row=3, column=1, value="Дата: 01.07.2025")

# [header_leading_columns] первый заголовок сдвинут в колонку B (колонка A пустая)
# [header_count] 5 заголовков вместо 6 (убран Tax)
# [header_value] "Qty" → "Quantity", "Total" → "Sum"
make_header(ws_i, 4, [None, "ID", "Product", "Quantity", "Price", "Sum"])

for r, (rid, prod, qty, price) in enumerate(DATA, 5):
    ws_i.cell(row=r, column=2, value=rid)
    ws_i.cell(row=r, column=3, value=prod)
    ws_i.cell(row=r, column=4, value=qty)
    # [number_format] цена без денежного формата
    ws_i.cell(row=r, column=5, value=price)
    # [formula] формула изменена
    c6 = ws_i.cell(row=r, column=6, value=f"=D{r}*E{r}*1")
    # [number_format] формат Total изменён
    c6.number_format = "0.00"

# [merged_cells] объединения E2:F2 нет
# [sheet_order_or_names] Config удалён, Summary → Totals
ws_totals = wb_i.create_sheet("Totals")
ws_totals["A1"] = "Итого"
ws_totals["B1"] = "=SUM(Orders!F5:F7)*1.1"
ws_totals["B1"].number_format = "#,##0.00"

wb_i.save(OUT / "input_modified.xlsx")
print("Создан: input_modified.xlsx")

print()
print("Внесённые отличия (все отслеживаемые типы):")
print("  sheet_order_or_names     : 'Config' удалён; 'Summary' → 'Totals'")
print("  dimensions               : 6 столбцов → 7 (сдвиг + пустой A)")
print("  header_row_offset        : заголовки на строке 3 → строке 4")
print("  header_leading_columns   : первый заголовок в A → в B")
print("  header_count             : 6 заголовков → 5 (без Tax)")
print("  header_value             : 'Qty'→'Quantity', 'Total'→'Sum'")
print("  formula                  : =C*D → =D*E*1")
print("  number_format            : #,##0.00 → General (Price); #,##0.00 → 0.00 (Total)")
print("  merged_cells             : объединение E2:F2 отсутствует")
