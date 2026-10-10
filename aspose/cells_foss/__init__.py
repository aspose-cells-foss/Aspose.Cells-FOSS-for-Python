"""
Aspose.Cells for Python

A Python library for creating, reading, and modifying Excel files (.xlsx format).
This library provides a simple API compatible with Aspose.Cells for .NET structure.

Main Classes:
    - Workbook: Represents an Excel workbook
    - Worksheet: Represents a worksheet within a workbook
    - Cell: Represents a single cell
    - Cells: Collection of cells in a worksheet
    - Style: Represents cell formatting styles
"""

from .workbook import Workbook, SaveFormat
from .worksheet import Worksheet
from .cell import Cell
from .cells import Cells
from .style import Style, Font, NumberFormat
from .style_resolver import StyleResolver
from .text_measure import FontStrategy, TextMeasurer, FontResolution, MeasuredText, TextRun
from .value_formatter import ValueFormatter, DisplayValueOptions
from .rich_text import RichTextRun
from .encryption_params import (
    AgileEncryptionParameters,
    StandardEncryptionParameters,
    CipherAlgorithm,
    HashAlgorithm,
    get_default_encryption_params
)
from .xlsx_encryptor import encrypt_xlsx, decrypt_xlsx
from .data_validation import (
    DataValidation,
    DataValidationCollection,
    DataValidationType,
    DataValidationOperator,
    DataValidationAlertStyle,
    DataValidationImeMode
)
from .csv_handler import (
    CSVHandler,
    CSVLoadOptions,
    CSVSaveOptions,
    load_csv_workbook,
    save_workbook_as_csv
)
from .markdown_handler import (
    MarkdownHandler,
    MarkdownSaveOptions,
    save_workbook_as_markdown
)
from .json_handler import (
    JsonHandler,
    JsonSaveOptions,
    save_workbook_as_json
)
from .page_break import HorizontalPageBreakCollection, VerticalPageBreakCollection
from .render_bounds import RenderBounds, WorksheetBoundsResolver
from .row_height import RowHeightCalculator
from .pdf_options import PdfSaveOptions
from .pdf_layout import (
    PdfCellLayout, PdfDataBarLayout, PdfIconSetLayout,
    PdfPageLayout, PdfTextRunLayout,
    PdfLayoutEngine,
)
from .pdf_exporter import PdfExporter
from .pdf_renderer import SkiaPdfRenderer
from .chart import ChartType, Chart, ChartCollection, NSeries, ChartSeries, ChartAxis, ChartErrorBars, ChartView3D
from .picture import Picture, PictureCollection
from .shape import (
    MsoDrawingType, FillType, MsoLineDashStyle,
    TextAlignmentType, TextAnchorType,
    MsoFillFormat, MsoLineFormat, ShapeFont,
    Shape, ShapeCollection,
)
from .table import (
    TableStyleInfo, TableColumn, Table, TableCollection,
)
from .sparkline import (
    SparklineType, SparklineEmptyCells, Sparkline,
    SparklineGroup, SparklineGroupCollection,
)

__version__ = "26.10.0"
__all__ = [
    "Workbook",
    "SaveFormat",
    "Worksheet",
    "Cell",
    "Cells",
    "Style",
    "Font",
    "NumberFormat",
    "StyleResolver",
    "FontStrategy",
    "TextMeasurer",
    "FontResolution",
    "MeasuredText",
    "TextRun",
    "ValueFormatter",
    "DisplayValueOptions",
    "RichTextRun",
    "AgileEncryptionParameters",
    "StandardEncryptionParameters",
    "CipherAlgorithm",
    "HashAlgorithm",
    "get_default_encryption_params",
    "encrypt_xlsx",
    "decrypt_xlsx",
    "DataValidation",
    "DataValidationCollection",
    "DataValidationType",
    "DataValidationOperator",
    "DataValidationAlertStyle",
    "DataValidationImeMode",
    "CSVHandler",
    "CSVLoadOptions",
    "CSVSaveOptions",
    "load_csv_workbook",
    "save_workbook_as_csv",
    "MarkdownHandler",
    "MarkdownSaveOptions",
    "save_workbook_as_markdown",
    "JsonHandler",
    "JsonSaveOptions",
    "save_workbook_as_json",
    "HorizontalPageBreakCollection",
    "VerticalPageBreakCollection",
    "RenderBounds",
    "WorksheetBoundsResolver",
    "RowHeightCalculator",
    "PdfSaveOptions",
    "PdfCellLayout",
    "PdfDataBarLayout",
    "PdfIconSetLayout",
    "PdfPageLayout",
    "PdfTextRunLayout",
    "PdfLayoutEngine",
    "PdfExporter",
    "SkiaPdfRenderer",
    "ChartType",
    "Chart",
    "ChartCollection",
    "NSeries",
    "ChartSeries",
    "ChartAxis",
    "ChartErrorBars",
    "ChartView3D",
    "Picture",
    "PictureCollection",
    "MsoDrawingType",
    "FillType",
    "MsoLineDashStyle",
    "TextAlignmentType",
    "TextAnchorType",
    "MsoFillFormat",
    "MsoLineFormat",
    "ShapeFont",
    "Shape",
    "ShapeCollection",
    "TableStyleInfo",
    "TableColumn",
    "Table",
    "TableCollection",
    "SparklineType",
    "SparklineEmptyCells",
    "Sparkline",
    "SparklineGroup",
    "SparklineGroupCollection",
]
