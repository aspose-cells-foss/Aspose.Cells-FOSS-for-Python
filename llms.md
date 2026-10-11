# Aspose.Cells FOSS for Python — LLM Context

Aspose.Cells FOSS for Python is a pure-Python library for creating, loading, editing, and
exporting Excel `.xlsx` workbooks without Microsoft Excel. The current package version is
`26.10.0`, and the import namespace is `aspose.cells_foss`.

This document is a high-confidence usage guide, not an exhaustive API reference. When it
conflicts with the repository, use the following sources in order:

1. `aspose/cells_foss/` for behavior
2. `aspose/cells_foss/__init__.py::__all__` for supported top-level imports
3. `examples/` for executable public workflows
4. `README.md` for installation and scope

## Installation

```bash
pip install aspose-cells-foss
```

Python 3.10 or later is required. Core package metadata declares `pycryptodome>=3.15.0` and
`olefile>=0.46`.

PDF rendering additionally requires `skia-python`:

```bash
pip install "skia-python>=144.0.post2,<145"
```

## Rules for Generated Code

- Use A1 strings for direct cell access: `worksheet.cells["A1"]`.
- Do not generate tuple keys such as `worksheet.cells[0, 0]`.
- Prefer snake_case methods. Some classes provide PascalCase compatibility aliases, but this is
  not universal; never invent an alias.
- Use names exported from `aspose.cells_foss.__all__` for top-level imports. Loader, saver, CFB,
  and renderer implementation classes are internal unless explicitly exported there.
- Worksheet indexes and row/column coordinate APIs are generally zero-based. A1 references are
  the preferred user-facing form.
- Preserve formulas as strings beginning with `=`.
- Use `SaveFormat` or a supported filename extension to select an output format.
- Do not claim support for `.xls`, `.ods`, or `.xlsb`.
- The library preserves many unmodified Open XML parts for round-trip fidelity. Do not advise
  users to rebuild unchanged workbook XML.

## Core Workbook Usage

```python
from aspose.cells_foss import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]

worksheet.cells["A1"].value = "Revenue"
worksheet.cells["B1"].value = 1250.50
worksheet.cells["B2"].formula = "=SUM(B1:B1)"

workbook.save("output.xlsx")
```

Load an existing or encrypted workbook:

```python
from aspose.cells_foss import Workbook

workbook = Workbook("input.xlsx")
protected = Workbook("protected.xlsx", password="secret")
```

Workbook worksheet management uses methods on `Workbook`, not `workbook.worksheets.add()`:

```python
worksheet = workbook.add_worksheet("Data")
same_sheet = workbook.get_worksheet("Data")
copy = workbook.copy_worksheet("Data")
workbook.set_active_worksheet("Data")
workbook.remove_worksheet(copy.name)
```

## Cells and Styles

Assign through `Cell.value`; `Cell.put_value()` also exists, but property assignment is preferred.

```python
cell = worksheet.cells["A1"]
cell.value = "Quarterly Report"

style = cell.get_style()
style.font.bold = True
style.font.size = 14
style.font.color = "#FF0000"
style.alignment.horizontal = "center"
cell.apply_style(style)
```

Do not generate `font.is_bold`; the implemented property is `font.bold`. Cell-style color values
accept normalized hexadecimal strings used by the examples, including `#RRGGBB`.

Merge cells with an A1 range when possible:

```python
worksheet.cells.merge_range("A1:C1")
worksheet.cells.unmerge_range("A1:C1")
```

Comments use `set_comment`, not `add_comment`:

```python
worksheet.cells["A1"].set_comment("Review this value", author="Analyst")
```

## Common Worksheet Features

### Data validation

```python
from aspose.cells_foss import DataValidationType

validation = worksheet.data_validations.add("A1:A10")
validation.type = DataValidationType.LIST
validation.formula1 = '"Yes,No"'
```

### Hyperlinks

```python
worksheet.hyperlinks.add("A1", "https://example.com")
worksheet.hyperlinks.add("A2", sub_address="Sheet2!B5")
```

### Auto-filter and print area

```python
worksheet.auto_filter.range = "A1:E20"
worksheet.print_area = "A1:H40"
```

### Page breaks

```python
worksheet.horizontal_page_breaks.add(19)
worksheet.vertical_page_breaks.add(3)
```

These integer APIs use zero-based row and column indexes.

### Protection and encryption

Workbook or worksheet protection controls edit permissions; file encryption protects the package
with a password. They are separate features.

```python
workbook.protect(password="structure-password")
worksheet.protect(password="sheet-password")

workbook.save("encrypted.xlsx", password="file-password")
loaded = Workbook("encrypted.xlsx", password="file-password")
```

### Document properties

```python
workbook.document_properties.core.title = "Quarterly Report"
workbook.document_properties.core.creator = "Finance Team"
workbook.document_properties.core.subject = "Q4 Results"
```

## Charts, Pictures, Shapes, Tables, and Sparklines

Feature collections are exposed on each worksheet:

- `worksheet.charts`
- `worksheet.pictures`
- `worksheet.shapes`
- `worksheet.tables`
- `worksheet.sparkline_groups`
- `worksheet.conditional_formats`
- `worksheet.data_validations`
- `worksheet.hyperlinks`

Example chart creation:

```python
chart = worksheet.charts.add_line(0, 4, 20, 12)
chart.title = "Monthly Sales"
chart.n_series.add("B2:B7", category_data="A2:A7", name="Sales")
```

Example table creation:

```python
table = worksheet.tables.add(
    start_row=0,
    start_col=0,
    end_row=9,
    end_col=3,
    has_headers=True,
    name="SalesTable",
)
table.table_style_info.name = "TableStyleMedium9"
table.table_style_info.show_row_stripes = True
```

Example sparkline creation:

```python
from aspose.cells_foss import SparklineType

group = worksheet.sparkline_groups.add(
    sparkline_type=SparklineType.LINE,
    data_range="Sheet1!B2:F6",
    is_vertical=False,
    location_range="G2:G6",
)
```

Consult the matching `examples/test_<feature>.py` before generating advanced chart, conditional
formatting, shape, picture, table, or sparkline code. These APIs have feature-specific arguments
that should not be guessed.

## Load and Save Formats

`Workbook.save()` infers the format from the extension or accepts an explicit `SaveFormat`:

```python
from aspose.cells_foss import SaveFormat

workbook.save("output.xlsx")
workbook.save("output.csv")
workbook.save("output.tsv")
workbook.save("output.json")
workbook.save("output.md")
workbook.save("output.pdf")  # Requires skia-python
workbook.save("data.txt", SaveFormat.CSV)
```

Supported directions:

| Format | Load | Save |
|---|:---:|:---:|
| XLSX | Yes | Yes |
| Encrypted XLSX | Yes | Yes |
| CSV | Yes | Yes |
| TSV | No | Yes |
| JSON | No | Yes |
| Markdown | No | Yes |
| PDF | No | Yes |

CSV import is explicit:

```python
from aspose.cells_foss import Workbook, load_csv_workbook

workbook = load_csv_workbook("input.csv")

another = Workbook()
another.load_csv("input.csv")
```

PDF options are available through `PdfSaveOptions`:

```python
from aspose.cells_foss import PdfSaveOptions

options = PdfSaveOptions()
options.worksheet_indices = [0]
options.export_hidden_worksheets = False
workbook.save_as_pdf("output.pdf", options)
```

## Public API Groups

The top-level package exports these main groups:

- Core: `Workbook`, `SaveFormat`, `Worksheet`, `Cell`, `Cells`
- Styling and text: `Style`, `Font`, `NumberFormat`, `StyleResolver`, `ValueFormatter`,
  `RichTextRun`, text measurement types
- Encryption: encryption parameters, algorithms, `encrypt_xlsx`, `decrypt_xlsx`
- Data validation and text exports: validation enums and collections, CSV/JSON/Markdown handlers
  and options
- PDF: `PdfSaveOptions`, layout models, `PdfLayoutEngine`, `PdfExporter`, `SkiaPdfRenderer`
- Drawings and charts: chart, picture, and shape models and enums
- Tables and sparklines: table and sparkline models, collections, and enums

Always inspect `aspose/cells_foss/__init__.py` before asserting that a symbol supports a top-level
import.

## Limitations

- This is not a full Excel calculation engine. Formula strings and cached values round-trip, and
  a limited internal evaluator covers some expressions.
- Native workbook I/O targets `.xlsx`; legacy binary and OpenDocument spreadsheet formats are not
  supported.
- Some advanced Open XML content is preserved without an editable object model.
- PDF appearance depends on installed fonts and `skia-python`.

## Development and Verification

Use the repository examples as tests:

```bash
python -m pytest examples/test_<feature>.py -v
python -m pytest examples --collect-only -q -p no:cacheprovider
```

Examples should write generated files through `examples.output_path_helper.examples_output_path()`.
Do not commit generated files from `examples/outputfiles/`.
