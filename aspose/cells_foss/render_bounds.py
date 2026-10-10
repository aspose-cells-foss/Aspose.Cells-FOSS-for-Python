"""Worksheet content and printable-bound resolution for rendering."""

import math
import xml.etree.ElementTree as ET
from typing import NamedTuple

from .style import Style


class RenderBounds(NamedTuple):
    """A 1-based, inclusive rectangular range in worksheet coordinates."""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def start_row(self):
        return self.min_row

    @property
    def start_column(self):
        return self.min_col

    @property
    def end_row(self):
        return self.max_row

    @property
    def end_column(self):
        return self.max_col

    @property
    def row_count(self):
        return self.max_row - self.min_row + 1

    @property
    def column_count(self):
        return self.max_col - self.min_col + 1

    def to_a1(self):
        from .cells import Cells

        start = Cells.coordinate_to_string(self.min_row, self.min_col)
        end = Cells.coordinate_to_string(self.max_row, self.max_col)
        return start if start == end else f"{start}:{end}"

    @property
    def a1_range(self):
        return self.to_a1()

    def union(self, other):
        if other is None:
            return self
        return RenderBounds(
            min(self.min_row, other.min_row),
            min(self.min_col, other.min_col),
            max(self.max_row, other.max_row),
            max(self.max_col, other.max_col),
        )


class WorksheetBoundsResolver:
    """Resolve live worksheet bounds without relying on cached XLSX dimensions."""

    _default_style_state = None

    @classmethod
    def get_content_bounds(cls, worksheet, include_drawings=True):
        bounds = None

        for ref, cell in worksheet.cells.get_all_cells().items():
            if cls._cell_has_renderable_content(cell):
                bounds = cls._include(bounds, cls._from_a1(worksheet, ref))

        for merge_ref in getattr(worksheet, '_merged_cells', ()):
            bounds = cls._include(bounds, cls._from_a1(worksheet, merge_ref))

        for table in getattr(worksheet, 'tables', ()):
            bounds = cls._include(bounds, cls._from_a1(worksheet, getattr(table, 'ref', None)))

        for group in getattr(worksheet, 'sparkline_groups', ()):
            for sparkline in getattr(group, 'sparklines', ()):
                bounds = cls._include(
                    bounds,
                    cls._from_a1(worksheet, getattr(sparkline, 'cell_reference', None)),
                )

        if include_drawings:
            bounds = cls._include_collection(bounds, getattr(worksheet, 'charts', ()))
            bounds = cls._include_collection(bounds, getattr(worksheet, 'pictures', ()))
            visible_shapes = (
                shape for shape in getattr(worksheet, 'shapes', ())
                if not getattr(shape, 'is_hidden', False)
            )
            bounds = cls._include_collection(bounds, visible_shapes)

        return bounds

    @classmethod
    def get_printable_bounds(cls, worksheet, include_drawings=True):
        print_area = getattr(worksheet, 'print_area', None)
        if print_area:
            return tuple(
                cls._from_a1(worksheet, token.strip())
                for token in print_area.split(',')
                if token.strip()
            )

        content = cls.get_content_bounds(worksheet, include_drawings=include_drawings)
        return (content,) if content is not None else ()

    @classmethod
    def get_render_bounds(cls, worksheet, include_drawings=True):
        bounds = None
        for area in cls.get_printable_bounds(worksheet, include_drawings=include_drawings):
            bounds = cls._include(bounds, area)
        return bounds

    @classmethod
    def _cell_has_renderable_content(cls, cell):
        if cell.value is not None or cell.formula is not None or cell.comment is not None:
            return True
        if getattr(cell, '_source_style_idx', 0):
            return True
        if cls._default_style_state is None:
            cls._default_style_state = cls._object_state(Style())
        return cls._object_state(cell.style) != cls._default_style_state

    @classmethod
    def _object_state(cls, value):
        if hasattr(value, '__dict__'):
            return tuple(
                (key, cls._object_state(item))
                for key, item in sorted(vars(value).items())
            )
        if isinstance(value, (list, tuple)):
            return tuple(cls._object_state(item) for item in value)
        if isinstance(value, dict):
            return tuple(
                (key, cls._object_state(item))
                for key, item in sorted(value.items())
            )
        return value

    @staticmethod
    def _from_a1(worksheet, ref):
        if not ref:
            return None
        parts = str(ref).replace('$', '').split(':', 1)
        start_row, start_col = worksheet.cells.coordinate_from_string(parts[0])
        end_row, end_col = worksheet.cells.coordinate_from_string(parts[-1])
        return RenderBounds(
            min(start_row, end_row),
            min(start_col, end_col),
            max(start_row, end_row),
            max(start_col, end_col),
        )

    @classmethod
    def _include_collection(cls, bounds, collection):
        for drawing in collection:
            bounds = cls._include(bounds, cls._drawing_bounds(drawing))
        return bounds

    @staticmethod
    def _drawing_bounds(drawing):
        upper_row = int(getattr(drawing, '_upper_left_row'))
        upper_col = int(getattr(drawing, '_upper_left_column'))
        lower_row = int(getattr(drawing, '_lower_right_row'))
        lower_col = int(getattr(drawing, '_lower_right_column'))

        min_row = min(upper_row, lower_row) + 1
        min_col = min(upper_col, lower_col) + 1
        max_row = max(upper_row, lower_row)
        max_col = max(upper_col, lower_col)

        if lower_row >= upper_row and getattr(drawing, '_lower_right_row_offset', 0) > 0:
            max_row += 1
        if lower_col >= upper_col and getattr(drawing, '_lower_right_column_offset', 0) > 0:
            max_col += 1

        source = getattr(drawing, '_source_chart_xml', None)
        if source:
            try:
                root = ET.fromstring(source)
                ns = {
                    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                }
                glow = root.find('c:spPr/a:effectLst/a:glow', ns)
                soft_edge = root.find(
                    'c:spPr/a:effectLst/a:softEdge', ns
                )
                scatter = root.find('.//c:scatterChart', ns)
                radius = (
                    float(glow.get('rad', 0.0)) / 12700.0
                    if glow is not None
                    and not (soft_edge is not None and scatter is not None)
                    else 0.0
                )
                column_padding = math.ceil(radius / 48.0)
                row_padding = math.ceil(radius / 15.0)
                min_col = max(1, min_col - column_padding)
                max_col += column_padding
                min_row = max(1, min_row - row_padding)
                max_row += row_padding
            except (ET.ParseError, TypeError, ValueError):
                pass

        return RenderBounds(min_row, min_col, max(min_row, max_row), max(min_col, max_col))

    @staticmethod
    def _include(bounds, addition):
        if addition is None:
            return bounds
        return addition if bounds is None else bounds.union(addition)
