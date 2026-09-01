# Excel MCP Server

<img src="https://github.com/pcvelz/excel-mcp-server/blob/main/docs/img/icon-800.png?raw=true" width="128">

[![NPM Version](https://img.shields.io/npm/v/excel-mcp-server-pcvelz)](https://www.npmjs.com/package/excel-mcp-server-pcvelz)

A Model Context Protocol (MCP) server that reads and writes MS Excel data.

## Why Fork?

This fork ([pcvelz/excel-mcp-server](https://github.com/pcvelz/excel-mcp-server)) adds features needed for real-world bookkeeping workflows that are not yet available upstream:

- **Alignment support in `excel_format_range`** — set horizontal/vertical alignment (e.g., left-align numbers)
- **Alignment in `showStyle` output** — read back cell alignment when using `showStyle: true`
- **Cell type attributes** — `type` attribute (string, number, formula, date) in `showStyle` output
- **Raw value attributes** — `raw` attribute showing unformatted values alongside displayed values
- **ISO date auto-conversion** — write ISO date strings (e.g., `"2026-02-03"`) and they're automatically converted to Excel date serial numbers
- **Sheet management** — `excel_rename_sheet`, `excel_delete_sheet` and `excel_move_sheet` for renaming, removing and reordering tabs
- **New workbooks from scratch** — `excel_write_to_sheet` with `newSheet: true` creates the file when the path does not exist yet
- **Row management** - `excel_delete_rows` and `excel_insert_rows`, with formulas adjusted the way Excel adjusts them (`#REF!` for references into deleted rows)
- **Merges, column widths, conditional formatting and data validation in `showStyle` output** — reported alongside the styled cell table

See the [upstream comparison](https://github.com/negokaz/excel-mcp-server/compare/main...pcvelz:excel-mcp-server:main) for full diff.

## Features

- Read/Write text values
- Read/Write formulas
- Create new sheets and new workbooks
- Rename, delete and reorder sheets

**🪟Windows only:**
- Live editing
- Capture screen image from a sheet

For more details, see the [tools](#tools) section.

## Requirements

- Node.js 20.x or later

## Supported file formats

- xlsx (Excel book)
- xlsm (Excel macro-enabled book)
- xltx (Excel template)
- xltm (Excel macro-enabled template)

## Installation

### Installing via NPM

excel-mcp-server is automatically installed by adding the following configuration to the MCP servers configuration.

For Windows:
```json
{
    "mcpServers": {
        "excel": {
            "command": "cmd",
            "args": ["/c", "npx", "--yes", "excel-mcp-server-pcvelz"],
            "env": {
                "EXCEL_MCP_PAGING_CELLS_LIMIT": "4000"
            }
        }
    }
}
```

For other platforms:
```json
{
    "mcpServers": {
        "excel": {
            "command": "npx",
            "args": ["--yes", "excel-mcp-server-pcvelz"],
            "env": {
                "EXCEL_MCP_PAGING_CELLS_LIMIT": "4000"
            }
        }
    }
}
```

### Installing via Smithery

To install Excel MCP Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/excel-mcp-server-pcvelz):

```bash
npx -y @smithery/cli install excel-mcp-server-pcvelz --client claude
```

<h2 id="tools">Tools</h2>

### `excel_describe_sheets`

List all sheet information of specified Excel file.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file

### `excel_read_sheet`

Read values from Excel sheet with pagination.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `range`
    - Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]
- `showFormula`
    - Show formula instead of value [default: false]
- `showStyle`
    - Show style information for cells [default: false]
    - When enabled, output includes:
        - `style-ref`: References to style definitions (border, font, fill, alignment, numFmt)
        - `type`: Cell type (number, string, date, bool, formula, error)
        - `raw`: Raw/unformatted value (shown when different from displayed value, e.g., `raw="45691"` for a date displayed as "3-Feb")
        - `merged cells` and `column widths` in the metadata list, since these live on the sheet rather than on individual cells

### `excel_screen_capture`

**[Windows only]** Take a screenshot of the Excel sheet with pagination.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `range`
    - Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]

### `excel_write_to_sheet`

Write values to the Excel sheet.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `newSheet` (optional, default: `false`)
    - Create a new sheet if true, otherwise write to the existing sheet
    - If the file itself does not exist yet, a new workbook is created containing just that sheet
- `range`
    - Range of cells to read in the Excel sheet (e.g., "A1:C10").
- `values`
    - Values to write to the Excel sheet. If the value is a formula, it should start with "="
    - ISO date strings (e.g., `"2026-02-03"`, `"2026-02-03T10:30:00"`, `"2026-02-03T10:30:00Z"`) are automatically converted to Excel date serial numbers

### `excel_create_table`

Create a table in the Excel sheet

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name where the table is created
- `range`
    - Range to be a table (e.g., "A1:C10")
- `tableName`
    - Table name to be created

### `excel_copy_sheet`

Copy existing sheet to a new sheet

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `srcSheetName`
    - Source sheet name in the Excel file
- `dstSheetName`
    - Sheet name to be copied

### `excel_rename_sheet`

Rename a sheet. Cell values, formatting, merged cells, column widths and row heights are preserved, and formulas referring to the sheet are updated to point at the new name.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Current name of the sheet to rename
- `newName`
    - New name for the sheet. Maximum 31 characters, and it cannot contain `: \ / ? * [ ]`

### `excel_delete_sheet`

Delete a sheet. Refuses to delete the last remaining sheet, and refuses to delete a sheet that formulas on other sheets still refer to unless `force` is set. Defined names pointing at the deleted sheet are removed.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Name of the sheet to delete
- `force` (optional, default: `false`)
    - Delete even when formulas on other sheets refer to this sheet, leaving those formulas broken. The affected cells are listed in the tool output.

### `excel_move_sheet`

Move a sheet to another position in the tab order. Content and formatting are untouched.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Name of the sheet to move
- `index`
    - Zero-based target position. `0` makes the sheet the first tab, which is the one shown when the workbook is opened.

### `excel_delete_rows`

Delete rows and shift the rows below them up. Formatting of the remaining rows, merged cells, column widths, conditional formatting and data validation are preserved. Formulas are adjusted the way Excel adjusts them: references below the deleted rows move up, ranges spanning them shrink, and references pointing only into the deleted rows become `#REF!`.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `startRow`
    - First row to delete, one-based and inclusive
- `endRow`
    - Last row to delete, one-based and inclusive

### `excel_insert_rows`

Insert empty rows and shift the rows at and below the insertion point down. Merged cells, conditional formatting, data validation and formulas move with the rows.

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `beforeRow`
    - One-based row number to insert before
- `count` (optional, default: `1`)
    - Number of rows to insert

### `excel_format_range`

Format cells in the Excel sheet with style information

**Arguments:**
- `fileAbsolutePath`
    - Absolute path to the Excel file
- `sheetName`
    - Sheet name in the Excel file
- `range`
    - Range of cells in the Excel sheet (e.g., "A1:C3")
- `styles`
    - 2D array of style objects for each cell. If a cell does not change style, use null. The number of items of the array must match the range size.
    - Style object properties:
        - `border`: Array of border styles (type, color, style)
        - `font`: Font styling (bold, italic, underline, size, strike, color, vertAlign)
        - `fill`: Fill/background styling (type, pattern, color, shading)
        - `alignment`: Cell alignment settings
            - `horizontal`: Horizontal alignment (left, center, right, fill, justify, centerContinuous, distributed)
            - `vertical`: Vertical alignment (top, center, bottom, justify, distributed)
            - `wrapText`: Wrap text in cell (boolean)
            - `shrinkToFit`: Shrink text to fit cell width (boolean)
            - `textRotation`: Text rotation angle (0-180, or 255 for vertical)
            - `indent`: Indent level (0-250)
        - `numFmt`: Custom number format string
        - `decimalPlaces`: Number of decimal places (0-30)

<h2 id="configuration">Configuration</h2>

You can change the MCP Server behaviors by the following environment variables:

### `EXCEL_MCP_PAGING_CELLS_LIMIT`

The maximum number of cells to read in a single paging operation.  
[default: 4000]

## License

Copyright (c) 2025 Kazuki Negoro

excel-mcp-server is released under the [MIT License](LICENSE)