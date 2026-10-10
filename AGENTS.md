# Aspose.Cells for Python Development Guide for AI Agents

You are a senior Python engineer working on a pure-Python Excel library. Prioritize Excel compatibility, round-trip fidelity, and small, reviewable diffs.

## Do

- Treat [`aspose/cells_foss/`](./aspose/cells_foss) as the source of truth for library behavior and public API
- Treat [`aspose/cells_foss/__init__.py`](./aspose/cells_foss/__init__.py) and its `__all__` list as the supported top-level import surface
- Treat [`examples/`](./examples) as executable usage coverage; keep examples aligned with the current API
- Use A1-style string references for cell access: `ws.cells["A1"]`
- Preserve loaded workbook content when it was not modified, especially XML parts cached on objects such as `_source_xml`
- Add new workbook features through the existing loader/saver split: `xml_<feature>_loader.py` and `xml_<feature>_saver.py`
- Keep PDF work separated into model-to-page layout (`pdf_layout.py`), workbook orchestration (`pdf_exporter.py`), and drawing (`pdf_renderer*.py`)
- Keep worksheet XML in ECMA-376-compatible element order when adding new nodes
- Use `SaveFormat` or file extensions consistently when saving workbooks
- Add or update example coverage when changing user-facing behavior in `aspose/cells_foss/`
- Prefer stdlib modules already used by the repo (`xml.etree.ElementTree`, `zipfile`) over new dependencies
- Keep `pyproject.toml`, `requirements.txt`, `README.md`, and `License/ThirdPartyNotices.txt` consistent when an approved dependency changes
- Include runtime data files in `tool.setuptools.package-data` and verify they are present in built distributions
- Run targeted tests for the area you changed before finishing

## Don't

- Never use tuple-based cell keys like `ws.cells[0, 0]`
- Never regenerate XML for loaded objects that were not changed if preserved source XML is available
- Never change public exports in [`aspose/cells_foss/__init__.py`](./aspose/cells_foss/__init__.py) without verifying the API impact
- Never add third-party dependencies without approval
- Never make PDF layout depend on a concrete rendering backend; renderer-specific code belongs in `pdf_renderer*.py`
- Never commit generated `.xlsx` files or contents from `outputfiles/`
- Never hard-code behavior in examples that contradicts the implementation in `aspose/cells_foss/`
- Never add comments that only restate the code

## PR Size Guidelines

Keep changes focused and easy to review.

- **Lines changed**: Prefer under 500 lines of code
- **Files changed**: Prefer under 10 code files
- **Single responsibility**: One feature, bugfix, or refactor per change

When the task is larger, split it by:

1. Core model/API changes in `aspose/cells_foss/`
2. XML load/save wiring
3. Example coverage in `examples/`

## Commands

Key commands:

```bash
python -m pytest examples -v
python -m pytest examples/test_<feature>.py -v
python -m pytest examples --collect-only -q -p no:cacheprovider
```

## Boundaries

### Always do

- Read the relevant modules in `aspose/cells_foss/` before changing behavior
- Update matching examples in `examples/` when you change a public workflow
- Update `README.md` when supported formats, installation requirements, or public workflows change
- Run relevant tests for the changed area
- Preserve backward-compatible API behavior unless the task explicitly requires a change

### Ask first

- Adding dependencies
- Changing public class names, method signatures, or exported symbols
- Deleting or renaming source files
- Large cross-cutting refactors across multiple workbook features

### Never do

- Commit secrets, credentials, or local machine paths
- Commit generated files from `outputfiles/`
- Replace preserved source XML with regenerated XML unless the object was actually modified
- Use tuple-style cell addressing

## Project Structure

```text
aspose/cells_foss/         # Library source code (canonical location)
  __init__.py              # Public exports
  workbook.py              # Workbook entry point and save/load dispatch
  worksheet.py             # Worksheet model
  cell.py / cells.py       # Cell model and A1-keyed collection
  style.py                 # Font, fill, border, alignment, number format
  chart.py                 # Chart models and enums
  picture.py / shape.py    # Drawing objects
  table.py                 # Excel table support
  sparkline.py             # Sparkline support
  data_validation.py       # Validation models and enums
  auto_filter.py           # Filter models
  document_properties.py   # Core and extended document properties
  workbook_properties.py   # Workbook-level settings and protection
  csv_handler.py           # CSV import/export
  markdown_handler.py      # Markdown export
  json_handler.py          # JSON export
  value_formatter.py       # Excel display-value formatting
  style_resolver.py        # Effective style resolution
  text_measure.py          # Font lookup and text measurement
  row_height.py            # Automatic row-height calculation
  rich_text.py             # Rich text runs
  pdf_options.py           # Public PDF save options
  pdf_layout.py            # Workbook-to-page PDF layout
  pdf_exporter.py          # Workbook-level PDF orchestration
  pdf_renderer*.py         # Skia PDF rendering and chart/drawing mixins
  render_bounds.py         # Rendered worksheet bounds
  xml_loader.py            # Workbook XML loading
  xml_saver.py             # Workbook XML saving
  xml_*_loader.py          # Feature-specific XML loaders
  xml_*_saver.py           # Feature-specific XML savers
examples/                  # Executable example tests for library features
  outputfiles/             # Output from examples/ tests
```

## Tech Stack

- **Language**: Python 3.7+
- **Workbook format**: `.xlsx` / ECMA-376 Open XML
- **XML**: `xml.etree.ElementTree`
- **Archives**: `zipfile`
- **Encryption**: `pycryptodome`
- **Encrypted container support**: `olefile`
- **PDF rendering**: `skia-python` (version selected by Python version in `requirements.txt`)
- **Optional font metrics**: Pillow when available; deterministic fallback otherwise
- **Testing**: `pytest`, `unittest`

## Code Examples

### Good cell access

```python
wb = Workbook()
ws = wb.worksheets[0]
ws.cells["A1"].value = "Revenue"
ws.cells["B2"].value = 42
```

### Bad cell access

```python
ws.cells[0, 0].value = "Revenue"
```

### Good feature extension pattern

```python
# Add model behavior in aspose/cells_foss/<feature>.py
# Load existing files in aspose/cells_foss/xml_<feature>_loader.py
# Save new or modified content in aspose/cells_foss/xml_<feature>_saver.py
```

### Example alignment

```python
from aspose.cells_foss import Workbook, DataValidationType
from examples.output_path_helper import examples_output_path

wb = Workbook()
ws = wb.worksheets[0]
dv = ws.data_validations.add("A1:A10")
dv.type = DataValidationType.LIST
dv.formula1 = '"Yes,No"'
wb.save(examples_output_path("example.xlsx"))
```

## PR Checklist

- [ ] Change is focused and reviewable
- [ ] Relevant tests pass
- [ ] `examples/` still reflects the current API
- [ ] Public exports in `aspose/cells_foss/__init__.py` are correct
- [ ] Runtime resources required by the changed feature are included in package data
- [ ] Dependency declarations and third-party notices agree
- [ ] No generated `.xlsx` files or `outputfiles/` artifacts are included

## When Stuck

- Compare generated workbook XML with a known-good Excel file
- Trace save behavior through `aspose/cells_foss/workbook.py`, `xml_saver.py`, and the relevant `xml_*_saver.py`
- Trace load behavior through `aspose/cells_foss/xml_loader.py` and the relevant `xml_*_loader.py`
- Trace PDF output through `pdf_exporter.py`, `pdf_layout.py`, and the relevant `pdf_renderer*.py` implementation
- Check `examples/` for the intended user-facing workflow before changing API behavior
- Write or update the smallest example or test that reproduces the problem
