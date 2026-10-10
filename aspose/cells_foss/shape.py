"""
Aspose.Cells for Python - Shape/Drawing Module

Provides APIs for worksheet drawing shapes (rectangles, ovals, text boxes,
arrows, stars, etc.) aligned with ECMA-376 Part 1 §20.5 (SpreadsheetML
drawings), the Excel VBA Shape/Shapes object model, and the Aspose.Cells
for .NET Shape/ShapeCollection APIs.

Shape XML element: xdr:sp inside xdr:twoCellAnchor / xdr:oneCellAnchor.
"""

from enum import IntEnum


class MsoDrawingType(IntEnum):
    """
    Shape preset geometry types (maps to ECMA-376 a:prstGeom prst attributes).

    Aligned with Excel VBA MsoAutoShapeType and Aspose.Cells .NET AutoShapeType.
    """
    UNKNOWN = 0
    RECTANGLE = 1             # prst="rect"
    ROUNDED_RECTANGLE = 2     # prst="roundRect"
    OVAL = 3                  # prst="ellipse"
    DIAMOND = 4               # prst="diamond"
    TRIANGLE = 5              # prst="triangle"      (isosceles)
    RIGHT_TRIANGLE = 6        # prst="rtTriangle"
    PARALLELOGRAM = 7         # prst="parallelogram"
    TRAPEZOID = 8             # prst="trapezoid"
    HEXAGON = 9               # prst="hexagon"
    OCTAGON = 10              # prst="octagon"
    CROSS = 11                # prst="plus"
    STAR_4 = 12               # prst="star4"
    STAR_5 = 13               # prst="star5"
    STAR_6 = 14               # prst="star6"
    STAR_7 = 15               # prst="star7"
    STAR_8 = 16               # prst="star8"
    RIGHT_ARROW = 17          # prst="rightArrow"
    LEFT_ARROW = 18           # prst="leftArrow"
    UP_ARROW = 19             # prst="upArrow"
    DOWN_ARROW = 20           # prst="downArrow"
    TEXT_BOX = 21             # prst="rect" with txBox=1 and noFill (Excel TextBox)
    CALLOUT = 22              # prst="wedgeRectCallout"
    PENTAGON = 23             # prst="homePlate"
    CLOUD = 24                # prst="cloud"
    HEART = 25                # prst="heart"
    LIGHTNING_BOLT = 26       # prst="lightningBolt"
    SMILEY_FACE = 27          # prst="smileyFace"
    LEFT_RIGHT_ARROW = 28     # prst="leftRightArrow"
    UP_DOWN_ARROW = 29        # prst="upDownArrow"
    CUBE = 30                 # prst="cube"
    BEVEL = 31                # prst="bevel"


class FillType(IntEnum):
    """Shape fill type (ECMA-376 a:spPr fill child elements)."""
    NONE = 0        # <a:noFill/>
    SOLID = 1       # <a:solidFill>
    GRADIENT = 2    # <a:gradFill> (round-trip only; generation defaults to solid)
    PATTERN = 3     # <a:pattFill> (round-trip only)
    AUTOMATIC = 4   # no explicit fill element (inherit from theme)


class MsoLineDashStyle(IntEnum):
    """Shape border/line dash style (ECMA-376 a:prstDash val attribute)."""
    SOLID = 0          # no <a:prstDash> element (solid line)
    DASH = 1           # val="dash"
    DOT = 2            # val="dot"
    DASH_DOT = 3       # val="dashDot"
    DASH_DOT_DOT = 4   # val="lgDashDotDot"
    LONG_DASH = 5      # val="lgDash"
    LONG_DASH_DOT = 6  # val="lgDashDot"


class TextAlignmentType(IntEnum):
    """Horizontal text alignment inside a shape (ECMA-376 a:pPr algn attribute)."""
    LEFT = 0        # algn="l"
    CENTER = 1      # algn="ctr"
    RIGHT = 2       # algn="r"
    JUSTIFY = 3     # algn="just"
    DISTRIBUTED = 4 # algn="dist"


class TextAnchorType(IntEnum):
    """Vertical text anchor inside a shape (ECMA-376 a:bodyPr anchor attribute)."""
    TOP = 0         # anchor="t"
    MIDDLE = 1      # anchor="ctr"
    BOTTOM = 2      # anchor="b"
    JUSTIFIED = 3   # anchor="just"


# -----------------------------------------------------------------------
# ECMA-376 preset geometry name ↔ MsoDrawingType mappings
# -----------------------------------------------------------------------

_DRAWING_TYPE_TO_PRST = {
    MsoDrawingType.RECTANGLE:       "rect",
    MsoDrawingType.ROUNDED_RECTANGLE: "roundRect",
    MsoDrawingType.OVAL:            "ellipse",
    MsoDrawingType.DIAMOND:         "diamond",
    MsoDrawingType.TRIANGLE:        "triangle",
    MsoDrawingType.RIGHT_TRIANGLE:  "rtTriangle",
    MsoDrawingType.PARALLELOGRAM:   "parallelogram",
    MsoDrawingType.TRAPEZOID:       "trapezoid",
    MsoDrawingType.HEXAGON:         "hexagon",
    MsoDrawingType.OCTAGON:         "octagon",
    MsoDrawingType.CROSS:           "plus",
    MsoDrawingType.STAR_4:          "star4",
    MsoDrawingType.STAR_5:          "star5",
    MsoDrawingType.STAR_6:          "star6",
    MsoDrawingType.STAR_7:          "star7",
    MsoDrawingType.STAR_8:          "star8",
    MsoDrawingType.RIGHT_ARROW:     "rightArrow",
    MsoDrawingType.LEFT_ARROW:      "leftArrow",
    MsoDrawingType.UP_ARROW:        "upArrow",
    MsoDrawingType.DOWN_ARROW:      "downArrow",
    MsoDrawingType.TEXT_BOX:        "rect",        # same prst as rectangle
    MsoDrawingType.CALLOUT:         "wedgeRectCallout",
    MsoDrawingType.PENTAGON:        "homePlate",
    MsoDrawingType.CLOUD:           "cloud",
    MsoDrawingType.HEART:           "heart",
    MsoDrawingType.LIGHTNING_BOLT:  "lightningBolt",
    MsoDrawingType.SMILEY_FACE:     "smileyFace",
    MsoDrawingType.LEFT_RIGHT_ARROW: "leftRightArrow",
    MsoDrawingType.UP_DOWN_ARROW:   "upDownArrow",
    MsoDrawingType.CUBE:            "cube",
    MsoDrawingType.BEVEL:           "bevel",
}

# Reverse mapping: prst → MsoDrawingType.
# TEXT_BOX is excluded because it shares "rect" with RECTANGLE; txBox="1"
# is used during loading to distinguish TEXT_BOX from RECTANGLE.
_PRST_TO_DRAWING_TYPE = {
    v: k for k, v in _DRAWING_TYPE_TO_PRST.items()
    if k != MsoDrawingType.TEXT_BOX
}

# Dash style value strings for a:prstDash
_DASH_STYLE_TO_VAL = {
    MsoLineDashStyle.DASH:         "dash",
    MsoLineDashStyle.DOT:          "dot",
    MsoLineDashStyle.DASH_DOT:     "dashDot",
    MsoLineDashStyle.DASH_DOT_DOT: "lgDashDotDot",
    MsoLineDashStyle.LONG_DASH:    "lgDash",
    MsoLineDashStyle.LONG_DASH_DOT: "lgDashDot",
}

_DASH_VAL_TO_STYLE = {v: k for k, v in _DASH_STYLE_TO_VAL.items()}

# TextAlignmentType → algn attribute value
_TEXT_ALIGN_TO_VAL = {
    TextAlignmentType.LEFT:        "l",
    TextAlignmentType.CENTER:      "ctr",
    TextAlignmentType.RIGHT:       "r",
    TextAlignmentType.JUSTIFY:     "just",
    TextAlignmentType.DISTRIBUTED: "dist",
}
_TEXT_ALIGN_VAL_TO_TYPE = {v: k for k, v in _TEXT_ALIGN_TO_VAL.items()}

# TextAnchorType → anchor attribute value
_ANCHOR_TO_VAL = {
    TextAnchorType.TOP:        "t",
    TextAnchorType.MIDDLE:     "ctr",
    TextAnchorType.BOTTOM:     "b",
    TextAnchorType.JUSTIFIED:  "just",
}
_ANCHOR_VAL_TO_TYPE = {v: k for k, v in _ANCHOR_TO_VAL.items()}


# -----------------------------------------------------------------------
# Format helper classes
# -----------------------------------------------------------------------

class MsoFillFormat:
    """
    Fill format properties for a shape.

    Attributes:
        fill_type (FillType): How the shape is filled.
        fore_color (str): Primary fill colour as a 6-hex-digit RRGGBB string.
        transparency (float): Opacity 0.0 (opaque) – 1.0 (fully transparent).
    """

    def __init__(self):
        self.fill_type = FillType.SOLID
        self.fore_color = "FFFFFF"
        self.transparency = 0.0

    def copy(self):
        f = MsoFillFormat()
        f.fill_type = self.fill_type
        f.fore_color = self.fore_color
        f.transparency = self.transparency
        return f


class MsoLineFormat:
    """
    Border/outline format properties for a shape.

    Attributes:
        is_visible (bool): Whether the border is drawn.
        color (str): Border colour as a 6-hex-digit RRGGBB string.
        weight (int): Line thickness in EMU (9525 = 0.75 pt; 12700 = 1 pt).
        dash_style (MsoLineDashStyle): Dash pattern.
        transparency (float): Opacity 0.0 (opaque) – 1.0 (fully transparent).
    """

    def __init__(self):
        self.is_visible = True
        self.color = "000000"
        self.weight = 9525        # 0.75 pt default (Excel default hairline border)
        self.dash_style = MsoLineDashStyle.SOLID
        self.transparency = 0.0

    def copy(self):
        l = MsoLineFormat()
        l.is_visible = self.is_visible
        l.color = self.color
        l.weight = self.weight
        l.dash_style = self.dash_style
        l.transparency = self.transparency
        return l


class ShapeFont:
    """
    Font properties for text inside a shape.

    Attributes:
        name (str): Typeface name (e.g. "Calibri").
        size (float): Font size in points.
        bold (bool): Bold formatting.
        italic (bool): Italic formatting.
        underline (bool): Underline formatting.
        color (str): Font colour as a 6-hex-digit RRGGBB string.
    """

    def __init__(self):
        self.name = "Calibri"
        self.size = 11.0
        self.bold = False
        self.italic = False
        self.underline = False
        self.color = "000000"

    def copy(self):
        f = ShapeFont()
        f.name = self.name
        f.size = self.size
        f.bold = self.bold
        f.italic = self.italic
        f.underline = self.underline
        f.color = self.color
        return f


# -----------------------------------------------------------------------
# Shape
# -----------------------------------------------------------------------

class Shape:
    """
    Represents a drawing shape (rectangle, oval, text box, arrow, etc.) on
    a worksheet.

    Aligned with ECMA-376 xdr:sp element, Excel VBA Shape object, and
    Aspose.Cells for .NET Shape class.

    Anchor coordinates use the same row/column system as Picture:
    - upper_left_row / upper_left_column: top-left cell (0-based)
    - lower_right_row / lower_right_column: bottom-right cell (0-based, exclusive)
    - Offsets are in EMU (English Metric Units).
    """

    def __init__(self, drawing_type=MsoDrawingType.RECTANGLE,
                 upper_left_row=0, upper_left_column=0,
                 lower_right_row=5, lower_right_column=5):
        self.name = ""
        self.drawing_type = drawing_type

        # Cell anchor
        self._upper_left_row = upper_left_row
        self._upper_left_column = upper_left_column
        self._lower_right_row = lower_right_row
        self._lower_right_column = lower_right_column
        self._upper_left_row_offset = 0
        self._upper_left_column_offset = 0
        self._lower_right_row_offset = 0
        self._lower_right_column_offset = 0

        # Formatting
        self.fill = MsoFillFormat()
        self.line = MsoLineFormat()

        # Text
        self.text = ""
        self.font = ShapeFont()
        self.text_horizontal_alignment = TextAlignmentType.LEFT
        self.text_vertical_alignment = TextAnchorType.TOP
        self.is_text_wrapped = True

        # Behaviour
        self.is_locked = True
        self.is_hidden = False
        self.placement = "twoCell"   # editAs attribute on the anchor element

        # Hyperlink URL (if set, a hyperlink relationship is written to drawing rels)
        self.hyperlink = None

        # Round-trip: raw xdr:sp XML string preserved from loaded file.
        # When set, the saver writes this XML verbatim inside the anchor element.
        self._source_xml = None
        self._preset_geometry = _DRAWING_TYPE_TO_PRST.get(
            drawing_type, "rect"
        )
        self._is_connector = False
        self._shape_id = None
        self._start_connection_id = None
        self._start_connection_index = None
        self._end_connection_id = None
        self._end_connection_index = None
        self._rotation = 0.0
        self._flip_horizontal = False
        self._flip_vertical = False
        self._text_direction = None
        self._tail_end = None
        self._transform_x = None
        self._transform_y = None
        self._transform_width = None
        self._transform_height = None

    # ---- anchor properties ----

    @property
    def upper_left_row(self):
        return self._upper_left_row

    @upper_left_row.setter
    def upper_left_row(self, v):
        self._upper_left_row = v

    @property
    def upper_left_column(self):
        return self._upper_left_column

    @upper_left_column.setter
    def upper_left_column(self, v):
        self._upper_left_column = v

    @property
    def lower_right_row(self):
        return self._lower_right_row

    @lower_right_row.setter
    def lower_right_row(self, v):
        self._lower_right_row = v

    @property
    def lower_right_column(self):
        return self._lower_right_column

    @lower_right_column.setter
    def lower_right_column(self, v):
        self._lower_right_column = v

    @property
    def upper_left_row_offset(self):
        return self._upper_left_row_offset

    @upper_left_row_offset.setter
    def upper_left_row_offset(self, v):
        self._upper_left_row_offset = v

    @property
    def upper_left_column_offset(self):
        return self._upper_left_column_offset

    @upper_left_column_offset.setter
    def upper_left_column_offset(self, v):
        self._upper_left_column_offset = v

    @property
    def lower_right_row_offset(self):
        return self._lower_right_row_offset

    @lower_right_row_offset.setter
    def lower_right_row_offset(self, v):
        self._lower_right_row_offset = v

    @property
    def lower_right_column_offset(self):
        return self._lower_right_column_offset

    @lower_right_column_offset.setter
    def lower_right_column_offset(self, v):
        self._lower_right_column_offset = v

    def copy(self):
        s = Shape(
            self.drawing_type,
            self._upper_left_row, self._upper_left_column,
            self._lower_right_row, self._lower_right_column,
        )
        s.name = self.name
        s._upper_left_row_offset = self._upper_left_row_offset
        s._upper_left_column_offset = self._upper_left_column_offset
        s._lower_right_row_offset = self._lower_right_row_offset
        s._lower_right_column_offset = self._lower_right_column_offset
        s.fill = self.fill.copy()
        s.line = self.line.copy()
        s.text = self.text
        s.font = self.font.copy()
        s.text_horizontal_alignment = self.text_horizontal_alignment
        s.text_vertical_alignment = self.text_vertical_alignment
        s.is_text_wrapped = self.is_text_wrapped
        s.is_locked = self.is_locked
        s.is_hidden = self.is_hidden
        s.placement = self.placement
        s.hyperlink = self.hyperlink
        s._source_xml = self._source_xml
        s._preset_geometry = self._preset_geometry
        s._is_connector = self._is_connector
        s._shape_id = self._shape_id
        s._start_connection_id = self._start_connection_id
        s._start_connection_index = self._start_connection_index
        s._end_connection_id = self._end_connection_id
        s._end_connection_index = self._end_connection_index
        s._rotation = self._rotation
        s._flip_horizontal = self._flip_horizontal
        s._flip_vertical = self._flip_vertical
        s._text_direction = self._text_direction
        s._tail_end = self._tail_end
        s._transform_x = self._transform_x
        s._transform_y = self._transform_y
        s._transform_width = self._transform_width
        s._transform_height = self._transform_height
        return s


# -----------------------------------------------------------------------
# ShapeCollection
# -----------------------------------------------------------------------

class ShapeCollection:
    """
    Collection of Shape objects on a worksheet.

    Accessed via worksheet.shapes.  Mirrors PictureCollection in design:
    - add() marks the drawing dirty so the drawing XML is regenerated on save.
    - _add_loaded() is used internally by the loader without marking dirty.
    """

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._shapes = []

    # ---- public API ----

    def add(self, drawing_type, upper_left_row, upper_left_column,
            lower_right_row, lower_right_column):
        """
        Adds a shape of the specified type to the worksheet.

        Args:
            drawing_type (MsoDrawingType): Preset geometry type.
            upper_left_row (int): 0-based row of the top-left anchor cell.
            upper_left_column (int): 0-based column of the top-left anchor cell.
            lower_right_row (int): 0-based row of the bottom-right anchor cell.
            lower_right_column (int): 0-based column of the bottom-right anchor cell.

        Returns:
            Shape: The newly created shape.
        """
        if isinstance(drawing_type, int):
            drawing_type = MsoDrawingType(drawing_type)
        shape = Shape(drawing_type, upper_left_row, upper_left_column,
                      lower_right_row, lower_right_column)
        shape.name = f"Shape {len(self._shapes) + 1}"
        self._shapes.append(shape)
        self._worksheet._mark_drawing_dirty()
        return shape

    def add_text_box(self, upper_left_row, upper_left_column,
                     lower_right_row, lower_right_column):
        """
        Adds a TextBox shape (transparent fill, no border) to the worksheet.

        Text boxes are rendered using an xdr:sp element with txBox="1" and
        no fill / no line, matching Excel's Insert > Text Box behaviour.

        Returns:
            Shape: The new TextBox shape.
        """
        shape = self.add(MsoDrawingType.TEXT_BOX,
                         upper_left_row, upper_left_column,
                         lower_right_row, lower_right_column)
        shape.fill.fill_type = FillType.NONE
        shape.line.is_visible = False
        return shape

    def AddTextBox(self, upper_left_row, upper_left_column,
                   lower_right_row, lower_right_column):
        """PascalCase alias of add_text_box()."""
        return self.add_text_box(upper_left_row, upper_left_column,
                                 lower_right_row, lower_right_column)

    # ---- internal ----

    def _add_loaded(self, shape):
        """Appends a shape loaded from an XLSX file (does not mark drawing dirty)."""
        self._shapes.append(shape)

    # ---- standard collection ----

    @property
    def count(self):
        return len(self._shapes)

    def __len__(self):
        return len(self._shapes)

    def __iter__(self):
        return iter(self._shapes)

    def __getitem__(self, index):
        return self._shapes[index]

    def copy(self, new_worksheet):
        """Returns a deep copy of this collection bound to new_worksheet."""
        new_coll = ShapeCollection(new_worksheet)
        for shape in self._shapes:
            new_coll._shapes.append(shape.copy())
        return new_coll
