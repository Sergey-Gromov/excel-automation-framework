# Test Plan

## 1. Workbook structure

Check:

- sheet names
- sheet order
- dimensions
- freeze panes
- merged cells

## 2. Values

Check:

- unchanged cell values
- expected new/modified values

## 3. Formulas

Check:

- existing formulas are preserved
- new formulas are present only where allowed

## 4. Formatting

Check:

- number formats
- fonts
- fills
- borders
- alignment
- protection
- column widths
- row heights

## 5. Regression protection

Check:

- every difference between `expected_output.xlsx` and `actual_output.xlsx`
- every difference between `input.xlsx` and `actual_output.xlsx`
- only differences listed in `ALLOWED_CHANGES.yml` are accepted

## 6. Reports

Generate:

- JSON diff report
- HTML diff report
