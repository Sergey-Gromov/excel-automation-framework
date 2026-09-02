# Excel Automation Framework

This is a **task repository**: it implements one specific Excel transformation.
The reusable diff/validation engine lives in a separate package,
[`excel-framework-core`](https://github.com/Sergey-Gromov/excel-framework-core),
installed as a dependency (see `requirements.txt`) — no particular folder
layout required, it's pulled straight from GitHub. See
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for how the two repos fit
together, and where an LLM is (and isn't) involved.

The goal is simple:

> Apply required Excel changes and automatically prove that nothing else changed.

## Project layout

```text
docs/
  ARCHITECTURE.md      How this repo and excel-framework-core fit together
  SPEC.md              Main specification and invariants
  TASKS.md             Current task for the coding agent
  TEST_PLAN.md         Test strategy
  ALLOWED_CHANGES.yml  Allowed workbook changes
  CHANGELOG.md
  DESIGN.md

prompts/
  01_specification.md
  02_test_plan.md
  03_implementation.md
  04_review.md
  05_regression.md
  06_release.md

transform.py            Entry point for your transformation (uses openpyxl)
run_production.py       Production wrapper: validates input structure before
                        transforming, quarantines drifted files instead
run_batch.py            Batch wrapper: validates and transforms every top-level
                        .xlsx file in an input directory

tests/
  test_expected_output.py
  test_no_unexpected_changes.py
  test_run_production.py
  test_run_batch.py
  conftest.py

samples/
  input.xlsx
  expected_output.xlsx
  actual_output.xlsx

reports/
  diff_report.html
  diff_report.json
```

The diff engine itself (`compare_workbooks`, `config`, `report`,
`validate_input`) is not part of this repository — it comes from the
`excel_framework` package, installed straight from its GitHub repo.

## Install

```bash
python -m venv .venv
source .venv/bin/activate      # Linux/macOS
# .venv\Scripts\activate       # Windows PowerShell

pip install -r requirements.txt   # installs excel-framework-core@v0.1.0 from GitHub, plus openpyxl/pytest
```

Requires git access to the private `excel-framework-core` repo — if you're
logged in via `gh auth setup-git` (see CI section below for token details),
this just works with no extra configuration.

## Validate input file against sample

Before running a transformation, check whether the incoming file matches the expected structure:

```bash
python -m excel_framework.validate_input samples/input.xlsx path/to/received_file.xlsx
```

This compares the received file against the sample and writes two reports:

```text
reports/input_vs_sample.html   # human-readable, colour-coded
reports/input_vs_sample.json   # machine-readable
```

The console also prints a summary of every difference (kind, sheet, cell address).

Options:

```text
--report-dir DIR   Output folder for reports (default: reports/)
--no-styles        Skip cell style comparison (faster for large files)
```

Exit code is `0` if no differences found, `1` otherwise — safe to use in CI.

## Running in production: handling input format drift

`transform.py` is deterministic — it doesn't adapt if the incoming file's
structure changes (header row moves, columns get added/removed/reordered,
formats change). Left to itself it will either crash, or worse, run
"successfully" while quietly writing into the wrong cells. Don't call
`transform.py` directly on unattended/production input — use
`run_production.py`, which gates the transformation behind the same
structural check as above:

```bash
python run_production.py path/to/incoming_file.xlsx path/to/output.xlsx
```

- **Structure matches the sample** → runs `transform.py`, writes the output,
  exits `0`.
- **Structure has drifted** → does **not** run the transformation. Instead
  it copies the incoming file plus the diff reports into
  `needs_review/<timestamp>_<filename>/` and exits `1`, so a human (or CI
  step / cron `MAILTO`) can pick it up.

This is a fail-fast/quarantine gate, not auto-repair — the framework's whole
premise is no unreviewed changes, so a drifted input is a stop-and-look
signal, not something the script should guess its way through. If the
format change is legitimate and recurring, that's a cue to go back to the
`prompts/01_specification.md` … `06_release.md` pipeline and update
`transform.py` (or make it read headers by name instead of fixed position,
the way `compare_workbooks` already does) to handle the new shape — see
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for when an LLM is (and isn't)
involved.

Options: `--sample PATH` (default `samples/input.xlsx`), `--report-dir DIR`
(default `reports/`), `--needs-review-dir DIR` (default `needs_review/`).

## Running a batch directory

`run_batch.py` processes every `.xlsx` workbook directly inside an input
directory:

```bash
python run_batch.py path/to/input_dir path/to/output_dir
```

- Processes only top-level `.xlsx` files in deterministic alphabetical order.
- Ignores `~$` temporary files, non-`.xlsx` files, and subdirectories.
- Writes each result to `output_dir/<same filename>.xlsx`.
- Validates each file against the sample before transforming it.
- Writes per-file reports under `reports/<file stem>/`.
- Quarantines drifted files under `needs_review/<file stem>/...` and continues.
- Does not overwrite existing output files unless `--overwrite` is supplied.
- With `--overwrite`, replaces existing outputs through an atomic temporary
  file in the output directory.

Exit code `0` means every discovered workbook transformed successfully, or no
eligible workbook was found. Exit code `1` means at least one workbook was
rejected or failed.

Options: `--sample PATH`, `--report-dir DIR`, `--needs-review-dir DIR`, and
`--overwrite`.

---

## Normal workflow

1. Put your sample source file into `samples/input.xlsx`.
2. Manually create the correct target result: `samples/expected_output.xlsx`.
3. Fill `docs/SPEC.md`.
4. Fill `docs/ALLOWED_CHANGES.yml`.
5. Implement transformation in `transform.py`.
6. Run:

```bash
python transform.py samples/input.xlsx samples/actual_output.xlsx
pytest
```

## Important rule

A script is not considered complete just because it runs without errors.

It is complete only when:

```bash
pytest
```

passes and the diff report contains no unexpected changes.

## Do you need Claude/an LLM to run this?

No — only to *build* it. Once `transform.py` is implemented and `pytest`
passes, running the transformation is a plain, deterministic Python script
with no AI dependency or network call:

```bash
python -m excel_framework.validate_input samples/input.xlsx new_file.xlsx   # optional pre-check
python transform.py new_file.xlsx output.xlsx
pytest
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the full dev-time vs
runtime breakdown.

## Report

After tests run, reports are written to:

```text
reports/diff_report.json
reports/diff_report.html
```

## Starting a new task

Each Excel transformation gets its own task repository, reusing the same
`excel-framework-core` engine. To start a new one:

1. Copy this whole folder to a new location — anywhere; `requirements.txt`
   installs the engine from GitHub, so there's no folder-layout requirement.
2. Clear out task-specific state: delete `samples/*.xlsx`, `reports/*`, and
   reset `docs/SPEC.md` and `docs/ALLOWED_CHANGES.yml` back to their blank
   templates (an empty `ALLOWED_CHANGES.yml` means nothing is allowed to
   change yet — start strict).
3. `pip install -r requirements.txt` in a fresh virtualenv.
4. Follow the `prompts/01_specification.md` … `06_release.md` pipeline again
   for the new task.

Improvements to the diff engine itself (new checks, bug fixes) belong in
`excel-framework-core`, not in a task repository. To publish them, tag a new
release there and bump the `@vX.Y.Z` reference in `requirements.txt` (and in
`.github/workflows/tests.yml`) of every task repo that should pick it up —
see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the release steps.

### Actively co-developing the engine and a task together

If you're editing `excel-framework-core` and want changes reflected
immediately (no tag/reinstall per edit), swap the git dependency for a local
editable install — this does require the two repos to sit next to each
other:

```bash
pip uninstall excel-framework-core
pip install -e ../excel-framework-core
```

Revert to the `git+https://…@vX.Y.Z` line in `requirements.txt` once you're
done and have cut a new tag.

## CI

`.github/workflows/tests.yml` runs `pytest` on every push/PR. It installs
`excel-framework-core` the same way `requirements.txt` does locally — a
release tag from its own GitHub repo — but since the runner has no
`gh auth setup-git` credential helper, it must embed a token explicitly. See
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md#local-dev-vs-ci-installing-the-same-way-different-auth)
for details.

### `CORE_REPO_TOKEN` secret

`excel-framework-core` is a private repository, so the default
`GITHUB_TOKEN` (scoped only to this repo) can't clone it. `CORE_REPO_TOKEN`
is a **fine-grained personal access token**, scoped to:

- Resource owner: `Sergey-Gromov`
- Repository access: only `excel-framework-core`
- Permissions: Contents → Read-only

Stored as a repo secret at Settings → Secrets and variables → Actions in
this repository. To set or rotate it, run (in a real terminal, not piped
through anything that could log the value):

```bash
gh secret set CORE_REPO_TOKEN --repo Sergey-Gromov/Excel-Automation-Framework
```

### Bumping the engine version in CI

When `excel-framework-core` gets a new release:

1. Tag it: `git tag vX.Y.Z && git push origin vX.Y.Z` (in `excel-framework-core`).
2. Update the tag in `.github/workflows/tests.yml` (`@v0.1.0` → `@vX.Y.Z`).
