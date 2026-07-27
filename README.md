# Excel Automation Framework

This is a **task repository**: it implements one specific Excel transformation.
The reusable diff/validation engine lives in a separate package,
[`excel-framework-core`](../excel-framework-core), installed as a dependency
(see `requirements.txt`). See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
for how the two repos fit together, local-vs-CI installs, and where an LLM
is (and isn't) involved.

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

transform.py           Entry point for your transformation (uses openpyxl)

tests/
  test_expected_output.py
  test_no_unexpected_changes.py
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
`excel_framework` package installed from `../excel-framework-core`.

## Install

```bash
python -m venv .venv
source .venv/bin/activate      # Linux/macOS
# .venv\Scripts\activate       # Windows PowerShell

pip install -r requirements.txt   # installs excel-framework-core from ../excel-framework-core, plus openpyxl/pytest
```

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

1. Copy this whole folder to a new location (ideally still a sibling of
   `excel-framework-core`, so the default relative path in `requirements.txt`
   keeps working — otherwise adjust the `-e ../excel-framework-core` line).
2. Clear out task-specific state: delete `samples/*.xlsx`, `reports/*`, and
   reset `docs/SPEC.md` and `docs/ALLOWED_CHANGES.yml` back to their blank
   templates (an empty `ALLOWED_CHANGES.yml` means nothing is allowed to
   change yet — start strict).
3. `pip install -r requirements.txt` in a fresh virtualenv.
4. Follow the `prompts/01_specification.md` … `06_release.md` pipeline again
   for the new task.

Improvements to the diff engine itself (new checks, bug fixes) belong in
`excel-framework-core`, not in a task repository — they immediately benefit
every task once re-installed (`pip install -e ../excel-framework-core`
picks up changes live, no reinstall needed while editable).

## CI

`.github/workflows/tests.yml` runs `pytest` on every push/PR. It installs
`excel-framework-core` from its own GitHub repo pinned to a release tag
rather than via `requirements.txt` — see
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md#local-dev-vs-ci-two-install-paths)
for why local dev and CI need different install paths.

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
