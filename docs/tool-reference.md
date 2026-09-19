# Excel tool reference

## Sommelier RPC 0.2 / native adapters

The TaskPane and native adapters use the [RPC reference](../skills/sommelier/references/protocol.md).
MCP names replace dots with underscores: `excel.range.read` becomes `excel_range_read`. Discover
all schemas through MCP `tools/list` and read `excel_guide` for examples, including approved chart
creation and listing. The legacy tools below are a separate API; they are not all available in Sommelier.

## Legacy driver tools

All ranges are explicit rectangular A1 addresses; whole rows, whole columns, and worksheet-qualified
range arguments are rejected. Worksheet is a separate field. Results include summary, target,
cell count, truncation, warnings, capability limitations, UI message, and machine details.

## Read tools

- `excel.get_selection`: bounded values, formulas, value types, dimensions, and address.
- `excel.get_workbook_structure`: worksheet used ranges, tables, and named-item metadata; no cells.
- `excel.read_range`: bounded values, formulas, and types for worksheet plus A1 range.
- `excel.read_table`: table headers and a row subset controlled by offset and limit.

## Write and structural tools

Tool metadata uses `classification: "write"` for cell mutations and `classification: "structural"`
for workbook structure changes. Both require approval by default.

- `excel.write_range`: rectangular values with exact shape validation.
- `excel.set_formulas`: rectangular strings beginning with `=`; no semantic formula validation.
- `excel.create_table`: explicit existing range and Excel-compatible unique name.
- `excel.format_range`: bold, italic, hex fill/font color, number format, left/center/right alignment,
  and autofit only.
- `excel.add_comment`: one cell, capability-gated; unsupported hosts return a typed error.

No tool accepts JavaScript, Office Scripts, Office.js snippets, arbitrary style objects, workbook
binaries, or unbounded workbook content.
