# Design

## Principle

The framework treats an Excel workbook as a structured document, not just a table.

Therefore the comparison engine checks:

- values
- formulas
- styles
- dimensions
- merged cells
- hidden rows and columns
- comments
- hyperlinks
- named ranges

## Main safety idea

The transformation may change only explicitly allowed locations.

Everything else is considered an invariant.
