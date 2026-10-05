# Excel Transformation Specification

## Goal

Transform the source workbook by moving two year-adjustment columns next to the
accrued-total column, deleting the `Исчезнувшие` column (source `EA` / 131) and
shifting the columns after it one position left, shifting the intervening
columns right by two positions, normalizing four date columns to `dd.mm.yyyy`
display, adding two row-7 `SUBTOTAL` formulas, and replacing the date in three
row-9 headers with the date from `sh513.xlsx!B1`. All unrelated workbook
content must be preserved.

---

## Source workbook

- Input file: `samples/input.xlsx`
- Reference result: `samples/expected_output.xlsx`
- Actual generated result: `samples/actual_output.xlsx`
- Production sample: `samples/input_big.xlsx`
- Default validation sample: `samples/input_structural.xlsx` (the production
  sample truncated to 100 rows; same structure/formats, fewer rows, so the
  validation gate is fast). Regenerate with `python samples/make_structural_sample.py`.
- Production reference result: `samples/expected_output_big.xlsx`
- Date source: `sh513.xlsx`
- Date source cell: `Ликвидность!B1`
- Replacement date format: `DD.MM.YYYY`
- Sheet: `Ликвидность`
- Header row: `9`
- Expected input dimensions: at least `10` data rows and `183` columns (`GA`)
- Expected output width: `182` columns (`FZ`) — the input is `183` columns and
  the `Исчезнувшие` column (source `EA` / 131) is deleted (see Change 2)
- Data rows: `10:last_data_row`, where `last_data_row` is one less than the
  largest row referenced by the row-7 `SUBTOTAL` formulas. The small sample has
  `last_data_row = 100`; the production sample has `last_data_row = 6969`.

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
| `Исчезнувшие` | `EA` / 131 |
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
- Delete source column `EA` / 131 (`Исчезнувшие`) — see Change 2.
- Net effect on the remaining columns:
  - Source `1:74` stays in place.
  - Source `BW:DZ` / 75:130 shifts right by two to final `BY:EB` / 77:132.
  - Source `EB:FE` / 132:161 shifts right by one to final `EC:FF` / 133:162
    (the two-position move minus the one-position deletion).
  - Source `FH:GA` / 164:183 shifts left by one to final `FG:FZ` / 163:182.

**Final column mapping:**

| Final columns | Source columns |
|---|---|
| `1:74` | `1:74` |
| `75` / `BW` | `162` / `FF` |
| `76` / `BX` | `163` / `FG` |
| `77:132` / `BY:EB` | `75:130` / `BW:DZ` |
| `133:162` / `EC:FF` | `132:161` / `EB:FE` |
| `163:182` / `FG:FZ` | `164:183` / `FH:GA` |

The output is `182` columns wide (the input is `183`; the `Исчезнувшие` column
is deleted by Change 2). It must not introduce the two blank columns present in
the accidental reference workbook.

### Change 2: Delete the `Исчезнувшие` column

**Sheet:** `Ликвидность`

**Operation:**

- Delete source column `EA` / 131 (`Исчезнувшие`).
- Every column to the right of `EA` shifts one position left (see the net
  effect in Change 1).
- The output is `182` columns wide instead of `183`.

**Formula semantics:**

- The row-7 `SUBTOTAL` formula in the deleted column (`EA7`) is removed with
  the column.
- No other formula in the workbook references the deleted column, so the
  deletion itself requires no reference translation. Formulas in the shifted
  columns are translated by the same logical-column remapping as Change 3.

**Merged ranges:**

- A merged range that spans the deleted column shrinks by one column (the
  standard Excel behavior for deleting a column inside a merge). See the
  remapped merged ranges in "Preserved workbook content".

**Validation:**

- The transformation rejects input if the `Исчезнувшие` header is missing from
  `EA` / 131 or has drifted.

### Change 3: Preserve and translate formulas in moved columns

**Sheet:** `Ликвидность`

**Formula semantics:**

- Formulas remain formulas.
- Formula references are remapped to the final positions of the referenced
  logical columns.
- String constants inside formulas are not modified.
- Existing row-7 `SUBTOTAL` formulas keep the `10:{last_data_row + 1}` row
  range and are updated to their final column letters. For the 100-row sample
  this range is `10:101`.
- Row-9 month formulas such as `="Июль 2023"` are moved unchanged as formula
  strings.

Examples:

| Source formula | Final cell | Final formula |
|---|---:|---|
| `FF7`: `=SUBTOTAL(109,FF10:FF101)` | `BW7` | `=SUBTOTAL(109,BW10:BW101)` |
| `FG7`: `=SUBTOTAL(109,FG10:FG101)` | `BX7` | `=SUBTOTAL(109,BX10:BX101)` |
| `BW7`: `=SUBTOTAL(109,BW10:BW101)` | `BY7` | `=SUBTOTAL(109,BY10:BY101)` |
| `FE7`: `=SUBTOTAL(109,FE10:FE101)` | `FF7` | `=SUBTOTAL(109,FF10:FF101)` |

### Change 4: Add row-7 totals

**Sheet:** `Ликвидность`

**Cells:**

| Cell | Formula | Number format |
|---|---|---|
| `BB7` | `=SUBTOTAL(109,BB10:BB{last_data_row + 1})` | `#,##0.00` |
| `CV7` | `=SUBTOTAL(109,CV10:CV{last_data_row + 1})` | `#,##0.00` |

For the 100-row sample these formulas are `=SUBTOTAL(109,BB10:BB101)` and
`=SUBTOTAL(109,CV10:CV101)`. These are final workbook positions, not source
header names.

### Change 5: Normalize date display formats

**Sheet:** `Ликвидность`

**Columns:** `I`, `J`, `S`, and first `T`

**Rows:** `10:{last_data_row}`

**Behavior:**

- Keep stored values as native Excel dates or `None`.
- Do not convert dates to text.
- Set number format to `dd.mm.yyyy`.

### Change 6: Replace dates in three row-9 headers

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
| `DF8:EM8` | `DH8:EN8` |
| `EN8:EP8` | `EO8:EQ8` |

Both ranges span the deleted `Исчезнувшие` column (`EA` / 131), so each shrinks
by one column in addition to the move remapping.

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

- `Ликвидность!BW1:FZ1048576` for the moved/shifted column block across
  supported row counts (the block now ends at `FZ` / 182 because the
  `Исчезнувшие` column is deleted)
- `Ликвидность!I10:I1048576` for date display format
- `Ликвидность!J10:J1048576` for date display format
- `Ликвидность!S10:S1048576` for date display format
- `Ликвидность!T10:T1048576` for date display format
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
- It predates the deletion of the `Исчезнувшие` column (Change 2): the
  reference still contains that column at `EA` / 131, so its column layout
  cannot match the `182`-column output. The reference was intentionally kept
  as-is; the deletion is covered by the contract and focused tests instead.
- Most row-7 formulas are replaced by static values that do not broadly match
  the input column sums or cached formula values.
- Row-9 month formulas are converted to plain text.
- The `CU9`, `CV9`, and `CY9` header dates in the reference workbook match the
  now-intentional Change 6 behavior. Any other header-date differences remain
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
- The output is `182` columns wide: the `Исчезнувшие` column (source `EA` /
  131) is deleted and every column to its right shifts one position left.
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

The default sample is `samples/input_structural.xlsx` — the production sample
truncated to 100 rows. It has the same structure and formats as the 6969-row
production workbooks, so the plain command is both correct and fast for them.
Pass `--sample samples/input_big.xlsx` only if you need every data row checked
against the full production sample (slower).

Before validation, `run_production.py` normalizes the input and the sample so
that ordinary variation does not trigger quarantine: the three dated row-9
headers are rewritten to the sample dates, and both sides are reduced to the
smaller data volume (row-7 `SUBTOTAL` ranges rewritten, extra data rows
dropped). This lets files of different sizes pass the gate while still
detecting structural drift. The sample's size therefore no longer needs to
match the incoming file, but its structure and formats should. The
transformation itself uses `sh513.xlsx!B1` as the replacement date.

Exit code `0` means the incoming file matched the sample structure and the
transformation ran. Exit code `1` means structural drift was detected and the
file was quarantined under `needs_review/` without being transformed.

## Batch production command

Routine batch command:

```bash
python run_batch.py path/to/input_dir path/to/output_dir
```

The default sample is `samples/input_structural.xlsx`, so the plain command is
correct and fast for the 6969-row production workbooks. Pass
`--sample samples/input_big.xlsx` only if you need every data row checked
against the full production sample (slower).

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

---

## Observability (timing, progress, logging)

The transformation is deterministic and produces the same workbook regardless of
these options. Observability only adds side-channel output; it never changes the
transformed cells, formulas, formats, or styles.

### Elapsed time

`transform.py` prints the total wall-clock execution time on the last line of
stdout, in whole minutes and seconds, e.g.:

```text
Elapsed time: 1m 05s
```

The value is measured around the `transform()` call in `main()`, so it covers
loading, validation, the column move, and saving.

### Progress bar

The two dominant loops (the cell snapshot and the column move) update a
single-line percentage progress bar on stderr, e.g.:

```text
[######################----]  80.0% (14688/18300)
```

The bar is enabled by default when stderr is a terminal (interactive runs) and
disabled when stderr is redirected (scripts, CI, tests). It can be forced on or
off with `--progress` / `--no-progress` on `transform.py`, `run_production.py`,
and `run_batch.py`.

### Logging

`transform.py`, `run_production.py`, and `run_batch.py` emit structured log
records to stderr at `INFO` level for the major steps (load, validate, read
date, replace headers, snapshot, move columns, apply formats, add formulas,
save). `--verbose` raises the level to `DEBUG`. Logging is configured by
`transform.setup_logging()` and is idempotent, so a test harness that already
configured logging is left untouched.

### Flags

All three entry points accept:

- `--progress` — force the progress bar on
- `--no-progress` — disable the progress bar
- `--verbose` — enable debug logging

Direct calls to `transform()` (as the tests do) default to `progress=False` and
do not configure logging, so test output stays clean.
