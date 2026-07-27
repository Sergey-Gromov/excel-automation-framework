# Architecture

This document explains how the pieces fit together across the whole
ecosystem, not just this one repository. For day-to-day commands, see the
main [README.md](../README.md).

## Two repositories, one engine

```
excel-framework-core/            excel_framework package: compare_workbooks,
                                  config, report, validate_input.
                                  No transformation logic, no task state.
        ^
        │  installed as a dependency
        │
Excel-Automation-Framework/      ONE task repo = ONE transformation.
  (this repo)                    Owns: transform.py, docs/SPEC.md,
                                  docs/ALLOWED_CHANGES.yml, samples/, tests/.
```

Every new Excel transformation gets its **own** task repo (see "Starting a
new task" in the README) that installs the same `excel-framework-core`
engine rather than copying its code. Bug fixes or new comparison checks are
made once in `excel-framework-core` and picked up by every task.

## Local dev vs CI: installing the same way, different auth

`requirements.txt` installs the engine straight from its GitHub repo, pinned
to a tag — this is location-independent, no sibling-folder requirement:

```
git+https://github.com/Sergey-Gromov/excel-framework-core.git@v0.1.0
```

`excel-framework-core` is private, so both local dev and CI need
authentication to clone it — they just get it differently:

| | How auth is provided |
|---|---|
| Local dev | Your own `gh auth setup-git`-configured git credential helper — no token in `requirements.txt` at all, git resolves it transparently. |
| CI (GitHub Actions) | The runner has no persistent credential helper, so the workflow embeds a token explicitly: `git+https://x-access-token:${{ secrets.CORE_REPO_TOKEN }}@github.com/...`. `CORE_REPO_TOKEN` is a fine-grained PAT scoped only to `excel-framework-core`, `Contents: Read-only`, stored as a repo secret. Full setup/rotation steps are in the README's CI section. |

### Releasing a new engine version

1. In `excel-framework-core`: commit, then `git tag vX.Y.Z && git push origin vX.Y.Z`.
2. In every task repo that should pick it up: bump `@v0.1.0` → `@vX.Y.Z` in
   both `requirements.txt` and `.github/workflows/tests.yml`, then
   `pip install -r requirements.txt` again.

### Actively co-developing the engine

Editing `excel-framework-core` and want changes visible immediately, without
cutting a tag per edit? Swap the git dependency for a local editable install
— this is the one case that does require the two repos to be sibling
folders:

```bash
pip uninstall excel-framework-core
pip install -e ../excel-framework-core
```

## Where an LLM fits in — and where it doesn't

An LLM (Claude, via the `prompts/01_specification.md` … `06_release.md`
pipeline) is a **development-time** tool only:

- writing/updating `transform.py` against `docs/SPEC.md`
- reviewing the diff for accidental data/formula/style loss
- deciding what belongs in `docs/ALLOWED_CHANGES.yml`

Once `transform.py` is implemented and `pytest` passes, running the
transformation is a plain, deterministic Python script — `openpyxl` in,
`openpyxl` out, compared cell-by-cell. Production use needs no LLM, no API
key, and no network call:

```bash
python transform.py new_file.xlsx output.xlsx
pytest
```

You'd go back to the LLM-assisted pipeline only when requirements change
(new column, new formula) or when the input becomes unstructured enough to
need interpretation rather than fixed rules — at which point this
deterministic-script architecture is the wrong tool regardless of LLM
involvement.
