"""
Aspose.Cells for Python - Worksheet Module

This module provides the Worksheet class which represents a single worksheet in an Excel workbook.
The Worksheet class provides methods to manage cells, ranges, and worksheet properties.

Compatible with Aspose.Cells for .NET API structure.
"""

from .cells import Cells
from .cell import Cell
from .auto_filter import AutoFilter
from .conditional_format import ConditionalFormatCollection
from .worksheet_properties import WorksheetProperties
from .hyperlink import Hyperlinks
from .data_validation import DataValidationCollection
from .page_break import HorizontalPageBreakCollection, VerticalPageBreakCollection
from .chart import ChartCollection
from .picture import PictureCollection
from .shape import ShapeCollection
from .table import TableCollection
from .sparkline import SparklineGroupCollection
from .render_bounds import WorksheetBoundsResolver
from .style_resolver import StyleResolver
from .text_measure import TextMeasurer
from .row_height import RowHeightCalculator


class SheetProtectionDictWrapper:
    """
    Dictionary-like wrapper around SheetProtection for backward compatibility.

    This wrapper allows accessing SheetProtection properties using dictionary syntax
    while maintaining backward compatibility with existing code that uses ws.protection['key'].
    """

    def __init__(self, sheet_protection):
        """
        Initialize wrapper with a SheetProtection object.

        Args:
            sheet_protection: The SheetProtection object to wrap.
        """
        self._protection = sheet_protection

    def __getitem__(self, key):
        """Get protection setting by key."""
        if key == 'protected':
            return self._protection.sheet
        elif key == 'password':
            return self._protection.password
        elif key == 'sheet':
            return self._protection.sheet
        elif key == 'objects':
            return self._protection.objects
        elif key == 'scenarios':
            return self._protection.scenarios
        elif key == 'format_cells':
            return self._protection.format_cells
        elif key == 'format_columns':
            return self._protection.format_columns
        elif key == 'format_rows':
            return self._protection.format_rows
        elif key == 'insert_columns':
            return self._protection.insert_columns
        elif key == 'insert_rows':
            return self._protection.insert_rows
        elif key == 'insert_hyperlinks':
            return self._protection.insert_hyperlinks
        elif key == 'delete_columns':
            return self._protection.delete_columns
        elif key == 'delete_rows':
            return self._protection.delete_rows
        elif key == 'select_locked_cells':
            return self._protection.select_locked_cells
        elif key == 'select_unlocked_cells':
            return self._protection.select_unlocked_cells
        elif key == 'sort':
            return self._protection.sort
        elif key == 'auto_filter':
            return self._protection.auto_filter
        elif key == 'pivot_tables':
            return self._protection.pivot_tables
        else:
            raise KeyError(f"Unknown protection key: {key}")

    def __setitem__(self, key, value):
        """Set protection setting by key."""
        if key == 'protected':
            self._protection.sheet = value
        elif key == 'password':
            self._protection.password = value
        elif key == 'sheet':
            self._protection.sheet = value
        elif key == 'objects':
            self._protection.objects = value
        elif key == 'scenarios':
            self._protection.scenarios = value
        elif key == 'format_cells':
            self._protection.format_cells = value
        elif key == 'format_columns':
            self._protection.format_columns = value
        elif key == 'format_rows':
            self._protection.format_rows = value
        elif key == 'insert_columns':
            self._protection.insert_columns = value
        elif key == 'insert_rows':
            self._protection.insert_rows = value
        elif key == 'insert_hyperlinks':
            self._protection.insert_hyperlinks = value
        elif key == 'delete_columns':
            self._protection.delete_columns = value
        elif key == 'delete_rows':
            self._protection.delete_rows = value
        elif key == 'select_locked_cells':
            self._protection.select_locked_cells = value
        elif key == 'select_unlocked_cells':
            self._protection.select_unlocked_cells = value
        elif key == 'sort':
            self._protection.sort = value
        elif key == 'auto_filter':
            self._protection.auto_filter = value
        elif key == 'pivot_tables':
            self._protection.pivot_tables = value
        else:
            raise KeyError(f"Unknown protection key: {key}")

    def get(self, key, default=None):
        """Get protection setting with default value."""
        try:
            return self[key]
        except KeyError:
            return default


class Worksheet:
    """
    Represents a single worksheet in an Excel workbook.
    
    The Worksheet class provides access to cells, ranges, and worksheet properties
    such as name, visibility, protection, and page setup settings.
    
    Examples:
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> ws.name = "MySheet"
        >>> ws.cells["A1"].value = "Hello"
    """
    
    def __init__(self, name="Sheet1"):
        """
        Initializes a new instance of the Worksheet class.
        
        Args:
            name (str, optional): Name of the worksheet. Defaults to "Sheet1".
        """
        self._name = name
        self._cells = Cells(self)
        self._visible = True
        self._tab_color = None
        self._auto_filter = AutoFilter()  # Auto filter settings
        self._conditional_formats = ConditionalFormatCollection()  # Conditional formatting
        self._data_validations = DataValidationCollection()  # Data validation rules
        self._properties = WorksheetProperties()  # Worksheet properties
        self._hyperlinks = Hyperlinks(self)  # Hyperlinks collection
        self._charts = ChartCollection(self)  # Charts collection
        self._pictures = PictureCollection(self)  # Pictures collection
        self._shapes = ShapeCollection(self)  # Drawing shapes collection
        self._tables = TableCollection(self)  # Structured tables (ListObjects)
        self._tables_dirty = False
        self._sparkline_groups = SparklineGroupCollection(self)  # Sparklines
        self._source_sparkline_xml = None   # raw <ext ...>...</ext> for round-trip
        self._sparklines_dirty = False
        self._source_drawing_xml = None
        self._source_drawing_rels_xml = None
        self._source_drawing_extra_parts = []
        self._drawing_dirty = False
        self._row_heights = {}  # Row index -> height (points)
        self._custom_height_rows = set()  # Rows whose height must not be auto-fitted
        self._loaded_height_rows = set()  # Rows with effective ht loaded from XLSX
        self._column_widths = {}  # Column index -> width (characters)
        self._best_fit_columns = set()  # Columns auto-fitted by Excel
        self._column_styles = {}  # Column index -> xf style index
        self._hidden_rows = set()  # Set of hidden row indices
        self._hidden_columns = set()  # Set of hidden column indices
        self._merged_cells = []  # List of merged ranges in A1 notation (e.g., "A1:B2")
        self._print_area = None  # Print area in A1 notation (e.g., "A1:D20" or "A1:D20,F1:F20")
        self._horizontal_page_breaks = set()  # Set of 0-based row indices with manual breaks
        self._vertical_page_breaks = set()  # Set of 0-based column indices with manual breaks
        self._horizontal_page_breaks_collection = HorizontalPageBreakCollection(self)
        self._vertical_page_breaks_collection = VerticalPageBreakCollection(self)
        
    # Properties
    
    @property
    def name(self):
        """
        Gets or sets the name of the worksheet.
        
        Returns:
            str: The name of the worksheet.
        """
        return self._name
    
    @name.setter
    def name(self, value):
        """
        Sets the name of the worksheet.
        
        Args:
            value (str): The new name for the worksheet.
        """
        self._name = value
    
    def rename(self, new_name):
        """
        Renames the worksheet.
        
        Args:
            new_name (str): The new name for the worksheet.
            
        Examples:
            >>> ws.rename("NewSheetName")
        """
        self._name = new_name
    
    @property
    def cells(self):
        """
        Gets the Cells collection for this worksheet.
        
        Returns:
            Cells: The Cells collection containing all cells in the worksheet.
        """
        return self._cells
    
    @property
    def visible(self):
        """
        Gets or sets the visibility state of the worksheet.
        
        Returns:
            bool or str: True for visible, False for hidden, 'veryHidden' for very hidden.
        """
        return self._visible
    
    @visible.setter
    def visible(self, value):
        """
        Sets the visibility state of the worksheet.

        Args:
            value (bool or str): True for visible, False for hidden, 'veryHidden' for very hidden.
        """
        self._visible = value

    def set_visibility(self, value):
        """Set worksheet visibility. Accepts True, False, or 'veryHidden'."""
        if value is True or value is False or value == 'veryHidden':
            self._visible = value
        else:
            raise ValueError(f"Invalid visibility value: {value!r}. Use True, False, or 'veryHidden'.")

    def get_visibility(self):
        """Return current worksheet visibility (True, False, or 'veryHidden')."""
        return self._visible

    @property
    def tab_color(self):
        """
        Gets or sets the tab color of the worksheet.
        
        Returns:
            str: The tab color in RGB hex format (AARRGGBB), or None if not set.
        """
        return self._tab_color
    
    @tab_color.setter
    def tab_color(self, value):
        """
        Sets the tab color of the worksheet.

        Args:
            value (str): The tab color in RGB hex format (AARRGGBB), or None to clear.
        """
        self._tab_color = value

    def set_tab_color(self, color):
        """Set the worksheet tab color. color must be an 8-char AARRGGBB hex string."""
        if color is not None and len(color) != 8:
            raise ValueError(f"Invalid tab color format: {color!r}. Expected 8-char AARRGGBB hex string.")
        self._tab_color = color.upper() if color else None

    def get_tab_color(self):
        """Return the worksheet tab color (8-char AARRGGBB hex) or None."""
        return self._tab_color

    def clear_tab_color(self):
        """Clear the worksheet tab color."""
        self._tab_color = None

    def set_page_orientation(self, orientation):
        """Set page orientation. orientation must be 'portrait' or 'landscape'."""
        self._properties.page_setup.orientation = orientation

    def get_page_orientation(self):
        """Return the page orientation ('portrait', 'landscape', or None)."""
        return self._properties.page_setup.orientation

    def set_paper_size(self, paper_size):
        """Set the paper size (integer code, e.g. 9 = A4)."""
        self._properties.page_setup.paper_size = paper_size

    def get_paper_size(self):
        """Return the paper size integer code."""
        return self._properties.page_setup.paper_size

    def set_page_margins(self, left=None, right=None, top=None, bottom=None,
                         header=None, footer=None):
        """Set page margins (in inches). Only provided values are updated."""
        margins = self._properties.page_margins
        if left is not None: margins.left = left
        if right is not None: margins.right = right
        if top is not None: margins.top = top
        if bottom is not None: margins.bottom = bottom
        if header is not None: margins.header = header
        if footer is not None: margins.footer = footer

    def get_page_margins(self):
        """Return a copy of the page margins dict."""
        return dict(self._properties.page_margins)

    def set_fit_to_pages(self, width=1, height=1):
        """Set fit-to-pages: width and height are page counts (0 = auto)."""
        page_setup = self._properties.page_setup
        page_setup.fit_to_width = width
        page_setup.fit_to_height = height

    def set_print_scale(self, scale):
        """Set print scale percentage (10–400)."""
        self._properties.page_setup.scale = scale

    @property
    def auto_filter(self):
        """
        Gets the AutoFilter object for this worksheet.
        
        Returns:
            AutoFilter: The AutoFilter object.
        """
        return self._auto_filter
    
    @property
    def conditional_formats(self):
        """
        Gets the collection of conditional formats for this worksheet.

        Returns:
            ConditionalFormatCollection: The collection of conditional formats.
        """
        return self._conditional_formats

    @property
    def data_validations(self):
        """
        Gets the collection of data validations for this worksheet.

        Data validation allows you to control what data can be entered into cells.
        You can create rules to restrict data entry to specific types, values,
        or ranges, and display messages or errors when invalid data is entered.

        Returns:
            DataValidationCollection: The collection of data validations.

        Examples:
            >>> from aspose_cells import Workbook
            >>> from aspose_cells.data_validation import DataValidationType, DataValidationOperator
            >>>
            >>> wb = Workbook()
            >>> ws = wb.worksheets[0]
            >>>
            >>> # Add a whole number validation
            >>> dv = ws.data_validations.add("A1:A10")
            >>> dv.type = DataValidationType.WHOLE_NUMBER
            >>> dv.operator = DataValidationOperator.BETWEEN
            >>> dv.formula1 = "1"
            >>> dv.formula2 = "100"
            >>> dv.show_error_message = True
            >>> dv.error_message = "Please enter a number between 1 and 100"
            >>>
            >>> # Add a dropdown list validation
            >>> dv2 = ws.data_validations.add("B1:B10")
            >>> dv2.type = DataValidationType.LIST
            >>> dv2.formula1 = '"Red,Green,Blue"'
            >>> dv2.show_dropdown = True
        """
        return self._data_validations

    @property
    def hyperlinks(self):
        """
        Gets the collection of hyperlinks for this worksheet.

        Returns:
            Hyperlinks: The collection of hyperlinks.

        Examples:
            >>> # Add external hyperlink
            >>> link = ws.hyperlinks.add("A1", "https://www.example.com")
            >>> link.text_to_display = "Visit Website"

            >>> # Add internal hyperlink
            >>> link = ws.hyperlinks.add("B2", sub_address="Sheet2!A1")
            >>> link.text_to_display = "Go to Sheet2"
        """
        return self._hyperlinks

    @property
    def charts(self):
        """
        Gets the collection of charts for this worksheet.

        Returns:
            ChartCollection: The chart collection.
        """
        return self._charts

    @property
    def shapes(self):
        """
        Gets the collection of drawing shapes for this worksheet.

        Returns:
            ShapeCollection: The shapes collection.
        """
        return self._shapes

    @property
    def tables(self):
        """
        Gets the collection of structured tables (ListObjects) for this worksheet.

        Returns:
            TableCollection: The tables collection.
        """
        return self._tables

    @property
    def sparkline_groups(self):
        """
        Gets the collection of sparkline groups for this worksheet.

        Returns:
            SparklineGroupCollection: The sparkline groups collection.
        """
        return self._sparkline_groups

    @property
    def pictures(self):
        """
        Gets the collection of pictures for this worksheet.

        Returns:
            PictureCollection: The picture collection.
        """
        return self._pictures

    def _mark_drawing_dirty(self):
        self._source_drawing_xml = None
        self._source_drawing_rels_xml = None
        self._source_drawing_extra_parts = []
        self._drawing_dirty = True

    @property
    def protection(self):
        """
        Gets the protection settings for this worksheet.

        Returns:
            SheetProtectionDictWrapper: Dictionary-like object containing protection settings.
        """
        return SheetProtectionDictWrapper(self._properties.protection)
    
    @property
    def page_setup(self):
        """
        Gets the page setup settings for this worksheet.
        
        Returns:
            PageSetup: The shared page setup model. Dictionary-style access is
                supported for backward compatibility.
        """
        return self._properties.page_setup
    
    @property
    def page_margins(self):
        """
        Gets the page margins for this worksheet.

        Returns:
            PageMargins: The shared page margins model. Dictionary-style access
                is supported for backward compatibility.
        """
        return self._properties.page_margins

    @property
    def properties(self):
        """
        Gets the worksheet properties.

        Returns:
            WorksheetProperties: The worksheet properties object containing
            view, selection, pane, format, protection, page setup, margins,
            header/footer, and print options.

        Examples:
            >>> ws.properties.view.show_grid_lines = False
            >>> ws.properties.view.zoom_scale = 80
            >>> ws.properties.page_setup.orientation = "landscape"
        """
        return self._properties

    @property
    def horizontal_page_breaks(self):
        """
        Gets manual horizontal page break collection.

        Returns:
            HorizontalPageBreakCollection: Collection for manual row page breaks.
        """
        return self._horizontal_page_breaks_collection

    @property
    def vertical_page_breaks(self):
        """
        Gets manual vertical page break collection.

        Returns:
            VerticalPageBreakCollection: Collection for manual column page breaks.
        """
        return self._vertical_page_breaks_collection

    @property
    def merged_cells(self):
        """
        Gets merged cell ranges in A1 notation.

        Returns:
            list[str]: Merged ranges (e.g., ["A1:B2", "D4:F4"]).
        """
        return list(self._merged_cells)

    @property
    def print_area(self):
        """
        Gets or sets print area in A1 notation.

        Returns:
            str or None: Print area string (e.g., "A1:D20" or "A1:D20,F1:F20"), or None.
        """
        return self._print_area

    @print_area.setter
    def print_area(self, value):
        """Sets print area in A1 notation."""
        if value is None:
            self._print_area = None
            return
        self._print_area = self._normalize_print_area(value)

    def set_print_area(self, print_area):
        """
        Sets the worksheet print area.

        Args:
            print_area (str): A1-style range(s), e.g. "A1:D20" or "A1:D20,F1:F20".
        """
        self.print_area = print_area

    def clear_print_area(self):
        """Clears the worksheet print area."""
        self._print_area = None

    def SetPrintArea(self, print_area):
        return self.set_print_area(print_area)

    def ClearPrintArea(self):
        return self.clear_print_area()

    @property
    def print_title_rows(self):
        """Rows repeated at the top of each printed page, such as ``1:3``."""
        return self.page_setup.print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value):
        self.page_setup.print_title_rows = value

    @property
    def print_title_columns(self):
        """Columns repeated at the left of each printed page, such as ``A:C``."""
        return self.page_setup.print_title_columns

    @print_title_columns.setter
    def print_title_columns(self, value):
        self.page_setup.print_title_columns = value

    def set_print_titles(self, rows=None, columns=None):
        """Set repeated print-title rows and columns; ``None`` clears an axis."""
        self.print_title_rows = rows
        self.print_title_columns = columns

    def clear_print_titles(self):
        """Clear repeated rows and columns."""
        self.set_print_titles()

    def get_print_title_ranges(self):
        """Return 1-based inclusive row/column tuples for pagination."""
        rows = self._parse_print_title_indexes(self.print_title_rows, rows=True)
        columns = self._parse_print_title_indexes(
            self.print_title_columns, rows=False
        )
        return rows, columns

    def SetPrintTitles(self, rows=None, columns=None):
        return self.set_print_titles(rows, columns)

    def ClearPrintTitles(self):
        return self.clear_print_titles()

    @staticmethod
    def _parse_print_title_indexes(value, rows):
        if value is None:
            return None
        start, end = value.split(':', 1)
        if rows:
            return int(start), int(end)
        return Cells.column_index_from_string(start), Cells.column_index_from_string(end)

    def _normalize_print_area(self, print_area):
        """
        Validates and normalizes print area string into uppercase A1 notation.
        """
        if not isinstance(print_area, str) or not print_area.strip():
            raise ValueError("print_area must be a non-empty string")

        parts = []
        for token in print_area.split(','):
            part = token.strip().upper().replace('$', '')
            if not part:
                continue
            if ':' in part:
                start_ref, end_ref = part.split(':', 1)
                Cells.coordinate_from_string(start_ref)
                Cells.coordinate_from_string(end_ref)
                parts.append(f"{start_ref}:{end_ref}")
            else:
                Cells.coordinate_from_string(part)
                parts.append(part)

        if not parts:
            raise ValueError("print_area must contain at least one valid range")
        return ','.join(parts)

    def get_content_bounds(self, include_drawings=True):
        """Return bounds of live renderable content, or None for an empty sheet."""
        return WorksheetBoundsResolver.get_content_bounds(
            self, include_drawings=include_drawings
        )

    def get_printable_bounds(self, include_drawings=True):
        """Return printable areas, preserving discontiguous explicit print areas."""
        return WorksheetBoundsResolver.get_printable_bounds(
            self, include_drawings=include_drawings
        )

    def get_render_bounds(self, include_drawings=True):
        """Return the final printable envelope consumed by layout and pagination."""
        return WorksheetBoundsResolver.get_render_bounds(
            self, include_drawings=include_drawings
        )

    def get_effective_style(
        self,
        cell_ref=None,
        row=None,
        column=None,
        include_conditional_formats=False,
    ):
        """
        Returns the resolved visual style for a cell.

        Args:
            cell_ref (str, optional): Cell reference in A1 notation.
            row (int, optional): 1-based row number when cell_ref is omitted.
            column (int or str, optional): 1-based column number or column
                letter when cell_ref is omitted.
            include_conditional_formats (bool): Whether matching conditional
                formatting differential styles are applied over the base style.
        """
        if cell_ref is not None:
            row, column = self.cells.coordinate_from_string(str(cell_ref).upper())
        return StyleResolver.resolve_cell_style(
            self,
            row,
            column,
            include_conditional_formats=include_conditional_formats,
        )

    def measure_text(self, text, style=None, width_points=None, wrap_text=False):
        """
        Measure text using the workbook's shared font strategy when available.
        """
        font = getattr(style, "font", None)
        if font is None:
            workbook = getattr(self, "_workbook", None)
            if workbook is not None and getattr(workbook, "_styles", None):
                font = workbook._styles[0].font
            else:
                font = Cell().style.font

        workbook = getattr(self, "_workbook", None)
        measurer = workbook.text_measurer if workbook is not None else TextMeasurer()
        return measurer.measure_text(
            text,
            font,
            width_points=width_points,
            wrap_text=wrap_text,
        )

    def measure_cell_text(self, cell_ref, width_points=None):
        """
        Measure a cell's display text using its resolved effective style.
        """
        row, column = self.cells.coordinate_from_string(str(cell_ref).upper())
        cell = self.cells[cell_ref]
        style = self.get_effective_style(row=row, column=column)
        text = cell.get_display_text()
        return self.measure_text(
            text,
            style=style,
            width_points=width_points,
            wrap_text=style.alignment.wrap_text,
        )

    def get_effective_default_row_height(self):
        """
        Return the workbook-font-derived default row height unless the sheet marks
        the stored default row height as custom/authoritative.
        """
        workbook = getattr(self, "_workbook", None)
        if workbook is None:
            return float(self.properties.format.default_row_height)
        font = workbook._styles[0].font
        resolution = workbook.font_strategy.resolve(font)
        if (
            self.properties.format.dy_descent is not None
            and resolution.resolved_family.lower() == 'calibri'
            and not self.properties.format.custom_height
        ):
            measured = workbook.text_measurer.measure_text('0', font)
            return round(measured.line_height_points, 2)
        return workbook.text_measurer.derive_default_row_height(
            font,
            stored_height=self.properties.format.default_row_height,
            is_custom=self.properties.format.custom_height,
        )

    def calculate_row_heights(self, bounds=None):
        """
        Return layout-time row heights in points without changing the worksheet.

        The returned mapping uses 1-based row numbers. Explicit custom heights
        are preserved; cached non-custom heights are recomputed from display text.
        """
        return RowHeightCalculator.calculate(self, bounds=bounds)

    def _calculate_row_heights_for_auto_fit(self, bounds, rows):
        """Internal helper that allows auto-fit to replace explicit row heights."""
        return RowHeightCalculator.calculate(
            self, bounds=bounds, ignore_custom_rows=rows
        )

    # Methods
    
    def is_protected(self):
        """
        Checks if the worksheet is protected.

        Returns:
            bool: True if the worksheet is protected, False otherwise.
        """
        return self._properties.protection.sheet
    
    def protect(self, password=None, format_cells=None, format_columns=None, format_rows=None,
                insert_columns=None, insert_rows=None, delete_columns=None, delete_rows=None,
                sort=None, auto_filter=None, insert_hyperlinks=None, pivot_tables=None,
                select_locked_cells=None, select_unlocked_cells=None, objects=None, scenarios=None):
        """
        Protects the worksheet with optional password and protection options.

        Args:
            password (str, optional): Password for worksheet protection. Defaults to None.
            format_cells (bool, optional): Allow formatting cells. Defaults to None (keeps current).
            format_columns (bool, optional): Allow formatting columns. Defaults to None (keeps current).
            format_rows (bool, optional): Allow formatting rows. Defaults to None (keeps current).
            insert_columns (bool, optional): Allow inserting columns. Defaults to None (keeps current).
            insert_rows (bool, optional): Allow inserting rows. Defaults to None (keeps current).
            delete_columns (bool, optional): Allow deleting columns. Defaults to None (keeps current).
            delete_rows (bool, optional): Allow deleting rows. Defaults to None (keeps current).
            sort (bool, optional): Allow sorting. Defaults to None (keeps current).
            auto_filter (bool, optional): Allow auto filter. Defaults to None (keeps current).
            insert_hyperlinks (bool, optional): Allow inserting hyperlinks. Defaults to None (keeps current).
            pivot_tables (bool, optional): Allow pivot tables. Defaults to None (keeps current).
            select_locked_cells (bool, optional): Allow selecting locked cells. Defaults to None (keeps current).
            select_unlocked_cells (bool, optional): Allow selecting unlocked cells. Defaults to None (keeps current).
            objects (bool, optional): Protect objects. Defaults to None (keeps current).
            scenarios (bool, optional): Protect scenarios. Defaults to None (keeps current).

        Examples:
            >>> ws.protect()  # Protect without password
            >>> ws.protect("mypassword")  # Protect with password
            >>> ws.protect("secure", format_cells=True, sort=False)  # Protect with options
        """
        # Enable sheet protection
        self._properties.protection.sheet = True

        # Set password (will be hashed when saving to XML)
        if password is not None:
            self._properties.protection.password = password

        # Set protection options if provided
        if format_cells is not None:
            self._properties.protection.format_cells = format_cells
        if format_columns is not None:
            self._properties.protection.format_columns = format_columns
        if format_rows is not None:
            self._properties.protection.format_rows = format_rows
        if insert_columns is not None:
            self._properties.protection.insert_columns = insert_columns
        if insert_rows is not None:
            self._properties.protection.insert_rows = insert_rows
        if delete_columns is not None:
            self._properties.protection.delete_columns = delete_columns
        if delete_rows is not None:
            self._properties.protection.delete_rows = delete_rows
        if sort is not None:
            self._properties.protection.sort = sort
        if auto_filter is not None:
            self._properties.protection.auto_filter = auto_filter
        if insert_hyperlinks is not None:
            self._properties.protection.insert_hyperlinks = insert_hyperlinks
        if pivot_tables is not None:
            self._properties.protection.pivot_tables = pivot_tables
        if select_locked_cells is not None:
            self._properties.protection.select_locked_cells = select_locked_cells
        if select_unlocked_cells is not None:
            self._properties.protection.select_unlocked_cells = select_unlocked_cells
        if objects is not None:
            self._properties.protection.objects = objects
        if scenarios is not None:
            self._properties.protection.scenarios = scenarios
    
    def unprotect(self, password=None):
        """
        Unprotects the worksheet.

        Args:
            password (str, optional): Password to unprotect the worksheet. Defaults to None.
                Note: Password validation is not currently implemented.

        Examples:
            >>> ws.unprotect()  # Unprotect without password
            >>> ws.unprotect("mypassword")  # Unprotect with password
        """
        # Note: We don't validate the password here - just disable protection
        # In a full implementation, we would check if the password matches
        self._properties.protection.sheet = False
        self._properties.protection.password = None
    
    def set_view(self, zoom=None, show_grid_lines=None, show_row_col_headers=None):
        """
        Sets view options for the worksheet.
        
        Args:
            zoom (int, optional): Zoom percentage (10-400). Defaults to None.
            show_grid_lines (bool, optional): Whether to show grid lines. Defaults to None.
            show_row_col_headers (bool, optional): Whether to show row and column headers. Defaults to None.
        """
        # This is a placeholder for future implementation
        pass
    
    def copy(self, name=None):
        """
        Creates a copy of the worksheet.
        
        Args:
            name (str, optional): Name for the copied worksheet. If None, a default name is generated.
            
        Returns:
            Worksheet: The copied worksheet.
        """
        new_ws = Worksheet(name if name else f"{self._name} (copy)")
        # Copy cells
        for ref, cell in self._cells._cells.items():
            new_ws._cells._cells[ref] = Cell(cell.value, cell.formula, worksheet=new_ws)
            new_ws._cells._cells[ref]._rich_text_runs = tuple(cell.rich_text_runs)
            if cell.style:
                new_ws._cells._cells[ref].style = cell.style.copy()
        new_ws._merged_cells = list(self._merged_cells)
        new_ws._row_heights = dict(self._row_heights)
        new_ws._custom_height_rows = set(self._custom_height_rows)
        new_ws._loaded_height_rows = set(self._loaded_height_rows)
        new_ws._print_area = self._print_area
        new_ws.page_setup.print_title_rows = self.print_title_rows
        new_ws.page_setup.print_title_columns = self.print_title_columns
        new_ws._horizontal_page_breaks = set(self._horizontal_page_breaks)
        new_ws._vertical_page_breaks = set(self._vertical_page_breaks)
        new_ws._charts = self._charts.copy(new_ws)
        new_ws._pictures = self._pictures.copy(new_ws)
        new_ws._shapes = self._shapes.copy(new_ws)
        new_ws._tables = self._tables.copy(new_ws)
        new_ws._tables_dirty = self._tables_dirty
        new_ws._sparkline_groups = self._sparkline_groups.copy(new_ws)
        new_ws._source_sparkline_xml = self._source_sparkline_xml
        new_ws._sparklines_dirty = self._sparklines_dirty
        new_ws._drawing_dirty = self._drawing_dirty
        return new_ws
    
    def delete(self):
        """
        Deletes the worksheet from the workbook.
        
        Note: This method is a placeholder. The actual deletion is handled by the Workbook class.
        """
        pass
    
    def move(self, index):
        """
        Moves the worksheet to the specified position in the workbook.
        
        Args:
            index (int): Zero-based index where to move the worksheet.
            
        Note: This method is a placeholder. The actual move is handled by the Workbook class.
        """
        pass
    
    def select(self):
        """
        Selects the worksheet.
        
        Note: This method is a placeholder. The actual selection is handled by the Workbook class.
        """
        pass
    
    def activate(self):
        """
        Activates the worksheet (makes it the active worksheet).
        
        Note: This method is a placeholder. The actual activation is handled by the Workbook class.
        """
        pass
    
    def calculate_formula(self):
        """
        Calculates formulas in the worksheet.
        
        Note: This method is a placeholder for future implementation.
        """
        pass
    
    def get_range(self, start_cell, end_cell=None):
        """
        Gets a range of cells.
        
        Args:
            start_cell (str): The starting cell reference (e.g., "A1").
            end_cell (str, optional): The ending cell reference. If None, returns a single cell.
            
        Returns:
            The range or cell.
            
        Note: This method is a placeholder for future implementation.
        """
        pass
