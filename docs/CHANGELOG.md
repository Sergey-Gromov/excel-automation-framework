# Changelog

## Unreleased

- Add `run_batch.py` for directory batch processing.
- Batch mode processes only top-level `.xlsx` files, preserves filenames,
  validates each file, quarantines drift, continues on failure, and supports
  atomic `--overwrite`.
- Add batch regression tests and documentation.

