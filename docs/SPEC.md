# Excel Transformation Specification

## Goal

Describe what the script must do to the Excel workbook.

Example:

> Transform the source workbook by renaming selected headers, inserting calculated columns, and reordering selected columns while preserving all unrelated workbook content.

---

## Source workbook

- Input file: `samples/input.xlsx`
- Expected result: `samples/expected_output.xlsx`
- Actual generated result: `samples/actual_output.xlsx`

---

## Required changes

### Change 1

**Description:**  
...

**Sheet:**  
...

**Range / columns:**  
...

**Expected result:**  
...

---

### Change 2

**Description:**  
...

**Sheet:**  
...

**Range / columns:**  
...

**Expected result:**  
...

---

## Invariants

The script must not change anything except explicitly allowed changes.

Forbidden unintended changes include:

- Removing or renaming sheets
- Changing data outside allowed ranges
- Removing existing formulas
- Changing formulas outside allowed ranges
- Changing styles outside allowed ranges
- Changing column widths
- Changing row heights
- Changing merged cells
- Changing hidden rows or columns
- Changing comments
- Changing hyperlinks
- Changing named ranges
- Changing Excel tables
- Changing filters
- Changing frozen panes
- Changing workbook-level metadata unless required

---

## Allowed changes

The authoritative allowed-change list is stored in:

```text
docs/ALLOWED_CHANGES.yml
```

---

## Definition of done

The task is complete only when:

- All required changes are implemented
- `samples/actual_output.xlsx` matches `samples/expected_output.xlsx`
- No unexpected differences are reported
- `pytest` passes
- Diff report is generated
