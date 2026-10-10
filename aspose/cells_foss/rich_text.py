"""Rich-text model used by loaded Excel shared strings."""

from dataclasses import dataclass


@dataclass(frozen=True)
class RichTextRun:
    """One formatted text run from an OOXML shared string."""

    text: str
    font_name: str = None
    font_size: float = None
    font_color: str = None
    bold: bool = None
    italic: bool = None
    underline: bool = None
    underline_type: str = None
    strikethrough: bool = None
    vertical_alignment: str = "baseline"
