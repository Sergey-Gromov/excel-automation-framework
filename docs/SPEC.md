# Excel Transformation Specification

## Goal

Transform the source workbook by moving two year-adjustment columns next to the
accrued-total column, shifting the intervening columns right by two positions,
normalizing four date columns to `dd.mm.yyyy` display, adding two row-7
`SUBTOTAL` formulas, and replacing the date in three row-9 headers with the
date from `sh513.xlsx!B1`. All unrelated workbook content must be preserved.

---

## Source workbook

- Input file: `samples/input.xlsx`
- Reference result: `samples/expected_output.xlsx`
- Actual generated result: `samples/actual_output.xlsx`
- Date source: `sh513.xlsx`
- Date source cell: `Ликвидность!B1`
- Replacement date format: `DD.MM.YYYY`
- Sheet: `Ликвидность`
- Header row: `9`
- Expected dimensions: `100` rows and `183` columns (`GA`)

The reference result contains known accidental differences and is therefore not
used as a full-cell equality target. See "Known accidental reference differences".

---

## Column identification

Columns are identified by row-9 header text after stripping surrounding spaces.
The transformation rejects input if any required header is missing or is not in
the expected sample position.

Required headers and sample positions:

| Header | Sample column |
|---|---:|
| `Дата заключения` | `I` / 9 |
| `Дата расторжения договора` | `J` / 10 |
| `ДЗ начальная дата образования` | `S` / 19 |
| `ДЗ конечная дата образования` | first occurrence, `T` / 20 |
| `Начислено нарастающим итогом с начала периода с учетом всех перерасчетов` | `BV` / 74 |
| `Оплачено нарастающим итогом с начала периода` | `BW` / 75 |
| `Сумма годовой корректировки` | `FF` / 162 |
| `Перерасчет суммы годовой корректировки` | `FG` / 163 |

The workbook contains a second `ДЗ конечная дата образования` header at `Y` / 25.
Only the first occurrence at `T` / 20 is normalized.

Dated headers are identified by a stable text prefix plus a `DD.MM.YYYY` date
in row `9`. The transformation rejects input if a dated header is missing,
duplicated, or moved from its sample position.

| Header prefix | Source column | Final column |
|---|---:|---:|
| `БУ ДЗ всего на` | `CS` / 97 | `CU` / 99 |
| `Сумма просроченной ДЗ с учетом сроков оплаты счетов по состоянию на` | `CT` / 98 | `CV` / 100 |
| `БУ Авансы на` | `CW` / 101 | `CY` / 103 |

---

## Required changes

### Change 1: Move year-adjustment columns

**Sheet:** `Ликвидность`

**Operation:**

- Move source column `FF` / 162 (`Сумма годовой корректировки`) to final column
  `BW` / 75.
- Move source column `FG` / 163 (`Перерасчет суммы годовой корректировки`) to
  final column `BX` / 76.
- Shift source columns `BW:FE` / 75:161 right by two columns to final
  `BY:FG` / 77:163.
- Leave source columns `1:74` and `164:183` in their original positions.

**Final column mapping:**

| Final columns | Source columns |
|---|---|
| `1:74` | `1:74` |
| `75` / `BW` | `162` / `FF` |
| `76` / `BX` | `163` / `FG` |
| `77:163` / `BY:FG` | `75:161` / `BW:FE` |
| `164:183` | `164:183` |

The output must remain `183` columns wide. It must not introduce the two blank
columns present in the accidental reference workbook.

### Change 2: Preserve and translate formulas in moved columns

**Sheet:** `Ликвидность`

**Formula semantics:**

- Formulas remain formulas.
- Formula references are remapped to the final positions of the referenced
  logical columns.
- String constants inside formulas are not modified.
- Existing row-7 `SUBTOTAL` formulas keep the `10:101` row range and are updated
  to their final column letters.
- Row-9 month formulas such as `="Июль 2023"` are moved unchanged as formula
  strings.

Examples:

| Source formula | Final cell | Final formula |
|---|---:|---|
| `FF7`: `=SUBTOTAL(109,FF10:FF101)` | `BW7` | `=SUBTOTAL(109,BW10:BW101)` |
| `FG7`: `=SUBTOTAL(109,FG10:FG101)` | `BX7` | `=SUBTOTAL(109,BX10:BX101)` |
| `BW7`: `=SUBTOTAL(109,BW10:BW101)` | `BY7` | `=SUBTOTAL(109,BY10:BY101)` |
| `FE7`: `=SUBTOTAL(109,FE10:FE101)` | `FG7` | `=SUBTOTAL(109,FG10:FG101)` |

### Change 3: Add row-7 totals

**Sheet:** `Ликвидность`

**Cells:**

| Cell | Formula | Number format |
|---|---|---|
| `BB7` | `=SUBTOTAL(109,BB10:BB101)` | `#,##0.00` |
| `CV7` | `=SUBTOTAL(109,CV10:CV101)` | `#,##0.00` |

These are final workbook positions, not source header names.

### Change 4: Normalize date display formats

**Sheet:** `Ликвидность`

**Columns:** `I`, `J`, `S`, and first `T`

**Rows:** `10:100`

**Behavior:**

- Keep stored values as native Excel dates or `None`.
- Do not convert dates to text.
- Set number format to `dd.mm.yyyy`.

### Change 5: Replace dates in three row-9 headers

**Sheet:** `Ликвидность`

**Date source:** `sh513.xlsx`, sheet `Ликвидность`, cell `B1`, formatted as
`DD.MM.YYYY`.

**Cells:**

| Source cell | Final cell | Header prefix |
|---|---|---|
| `CS9` | `CU9` | `БУ ДЗ всего на` |
| `CT9` | `CV9` | `Сумма просроченной ДЗ с учетом сроков оплаты счетов по состоянию на` |
| `CW9` | `CY9` | `БУ Авансы на` |

**Behavior:**

- Identify each header by its stable prefix, not by the full date-containing
  text.
- Replace only the existing `DD.MM.YYYY` portion with the date from
  `sh513.xlsx!B1`.
- Preserve all other header text, including trailing spaces.
- Preserve the header cell number format and style.
- Apply the replacement before the column-move snapshot so the moved headers
  carry the new date to their final positions.
- Fail with an actionable error if `sh513.xlsx!B1` is missing, is not a date,
  or cannot be parsed.

---

## Preserved workbook content

The transformation must preserve, within the moved/shifted columns:

- cell values
- formulas
- number formats, except explicitly changed date and new-total formats
- fonts, fills, borders, alignment, and protection
- hyperlinks and comments, if present
- column widths, hidden state, best-fit state, and outline levels
- row heights
- merged ranges, remapped through the final column mapping

Existing merged ranges and their expected remapped positions:

| Source merged range | Final merged range |
|---|---|
| `DF8:EM8` | `DH8:EO8` |
| `EN8:EP8` | `EP8:ER8` |

The transformation must not change:

- sheet names or sheet order
- unrelated cell values
- unrelated formulas
- unrelated styles
- unrelated column widths or row heights
- hidden rows or columns
- comments or hyperlinks
- named ranges
- Excel tables
- filters
- frozen panes
- workbook-level metadata except ordinary save behavior

---

## Allowed changes

The authoritative allowed-change list is stored in:

```text
docs/ALLOWED_CHANGES.yml
```

Allowed ranges:

- `Ликвидность!BW1:FG100` for the moved/shifted column block
- `Ликвидность!I10:I100` for date display format
- `Ликвидность!J10:J100` for date display format
- `Ликвидность!S10:S100` for date display format
- `Ликвидность!T10:T100` for date display format
- `Ликвидность!BB7:BB7` for the new row-7 formula
- `Ликвидность!CV7:CV7` for the new row-7 formula
- `Ликвидность!CU9:CU9` for the replaced `БУ ДЗ всего на` header date
- `Ликвидность!CV9:CV9` for the replaced просроченная ДЗ header date
- `Ликвидность!CY9:CY9` for the replaced `БУ Авансы на` header date

Merged-range changes are required by the column move but are not expressible in
the current `allowed_ranges` schema. They are verified directly by tests.

---

## Known accidental reference differences

`samples/expected_output.xlsx` was reviewed and contains differences that are
not part of the requested transformation:

- It has `185` columns, including blank columns `164` and `165`.
- Most row-7 formulas are replaced by static values that do not broadly match
  the input column sums or cached formula values.
- Row-9 month formulas are converted to plain text.
- The `CU9`, `CV9`, and `CY9` header dates in the reference workbook match the
  now-intentional Change 5 behavior. Any other header-date differences remain
  accidental.
- Date columns `I`, `J`, `S`, and `T` are stored as slash-separated text
  `DD/MM/YYYY`, which conflicts with the approved native-date `dd.mm.yyyy`
  behavior.

Therefore the completion target is the intended transformation contract plus
focused tests, not full equality with `samples/expected_output.xlsx`.

---

## Definition of done

The task is complete only when:

- `transform.py` applies only the changes described above.
- `samples/actual_output.xlsx` is generated from `samples/input.xlsx`.
- The output remains `183` columns wide.
- Moved, shifted, formula, date-format, merged-range, and style-preservation
  tests pass.
- Date-header replacement tests cover normal rows, blank values, formula
  placement, string dates, invalid dates, and missing or duplicate headers.
- Direct input-vs-actual tests verify that no unrelated cells, styles, merged
  ranges, or dimensions changed.
- `pytest` passes, including batch regression tests.
- Production execution and quarantine behavior are documented.
- `run_batch.py` processes a directory of workbooks without changing the
  single-workbook transformation rules.

Routine production command:

```bash
python run_production.py path/to/incoming.xlsx path/to/output.xlsx
```

Before validation, `run_production.py` normalizes only the three dated row-9
headers to the sample dates. This allows ordinary monthly date changes to pass
the gate while still detecting structural drift. The transformation itself
uses `sh513.xlsx!B1` as the replacement date.

Exit code `0` means the incoming file matched the sample structure and the
transformation ran. Exit code `1` means structural drift was detected and the
file was quarantined under `needs_review/` without being transformed.

## Batch production command

Routine batch command:

```bash
python run_batch.py path/to/input_dir path/to/output_dir
```

Batch mode must:

- process only `.xlsx` files directly inside the input directory
- ignore `~$` temporary files, non-`.xlsx` files, and subdirectories
- process files in deterministic alphabetical order
- preserve input filenames exactly in the output directory
- never modify or overwrite source files
- require the input and output directories to be different after path
  resolution
- validate each workbook independently using the same production validation
  flow as `run_production.py`
- write per-file reports under `reports/<file stem>/`
- quarantine drifted files under `needs_review/<file stem>/...`
- continue processing after rejection or failure
- catch unexpected per-file errors and continue with remaining files
- preserve existing output files unless `--overwrite` is supplied
- replace existing output files atomically when `--overwrite` is supplied

Batch exit code `0` means every discovered workbook transformed successfully,
or no eligible workbook was found. Batch exit code `1` means at least one
workbook was rejected or failed.
