"""Excel-like row auto-fit calculations for worksheet layout."""

import math


class RowHeightCalculator:
    """Computes effective row heights without mutating the worksheet."""

    CELL_HORIZONTAL_PADDING = 4.0
    CELL_VERTICAL_PADDING = 0.12

    @classmethod
    def calculate(cls, worksheet, bounds=None, ignore_custom_rows=None):
        bounds = bounds or worksheet.get_content_bounds(include_drawings=False)
        if bounds is None:
            return {}

        ignored = set(ignore_custom_rows or ())
        loaded_rows = set(getattr(worksheet, "_loaded_height_rows", ()))
        loaded_rotated_rows = cls._loaded_rows_with_rotation(
            worksheet, loaded_rows
        )
        loaded_rich_text_rows = cls._loaded_rows_with_rich_text(
            worksheet, loaded_rows
        )
        loaded_font_aligned_rows = cls._loaded_rows_with_font_alignment(
            worksheet, loaded_rows
        )
        fixed_rows = (
            set(getattr(worksheet, "_custom_height_rows", ()))
            | loaded_rotated_rows
            | loaded_rich_text_rows
            | loaded_font_aligned_rows
        ) - ignored
        default_height = worksheet.get_effective_default_row_height()
        loaded_height_scale = cls._loaded_height_scale(
            worksheet, default_height, bool(loaded_font_aligned_rows)
        )
        heights = {}
        for row in range(bounds.min_row, bounds.max_row + 1):
            if row in worksheet._hidden_rows:
                heights[row] = 0.0
            elif row in fixed_rows:
                height = float(worksheet._row_heights[row])
                if row in loaded_rows:
                    height *= loaded_height_scale
                heights[row] = height
            else:
                heights[row] = float(default_height)

        merges, covered = cls._merge_map(worksheet)
        pending_multirow_merges = []
        for ref in worksheet.cells.get_all_cells():
            row, column = worksheet.cells.coordinate_from_string(ref)
            if row < bounds.min_row or row > bounds.max_row or (row, column) in covered:
                continue

            merge = merges.get((row, column))
            min_row, min_col, max_row, max_col = merge or (row, column, row, column)
            if max_row == min_row and (
                row in fixed_rows or row in worksheet._hidden_rows
            ):
                continue
            width = cls._range_width_points(worksheet, min_col, max_col)
            required = cls._cell_height(worksheet, ref, width)
            if max_row == min_row:
                heights[row] = max(heights.get(row, default_height), required)
            else:
                pending_multirow_merges.append((min_row, max_row, required))

        for min_row, max_row, required in pending_multirow_merges:
            span_rows = [
                row for row in range(min_row, max_row + 1)
                if bounds.min_row <= row <= bounds.max_row
            ]
            current = sum(heights.get(row, default_height) for row in span_rows)
            adjustable = [
                row for row in span_rows
                if row not in fixed_rows and row not in worksheet._hidden_rows
            ]
            if required > current and adjustable:
                addition = (required - current) / len(adjustable)
                for row in adjustable:
                    heights[row] = heights.get(row, default_height) + addition

        return {row: round(height, 2) for row, height in sorted(heights.items())}

    @staticmethod
    def _loaded_rows_with_rotation(worksheet, loaded_rows):
        rotated_rows = set()
        for ref in worksheet.cells.get_all_cells():
            row, column = worksheet.cells.coordinate_from_string(ref)
            if row not in loaded_rows:
                continue
            style = worksheet.get_effective_style(row=row, column=column)
            if int(style.alignment.text_rotation or 0):
                rotated_rows.add(row)
        return rotated_rows

    @staticmethod
    def _loaded_rows_with_rich_text(worksheet, loaded_rows):
        return {
            worksheet.cells.coordinate_from_string(ref)[0]
            for ref, cell in worksheet.cells.get_all_cells().items()
            if (
                worksheet.cells.coordinate_from_string(ref)[0] in loaded_rows
                and cell.rich_text_runs
            )
        }

    @staticmethod
    def _loaded_rows_with_font_alignment(worksheet, loaded_rows):
        aligned_rows = set()
        for ref in worksheet.cells.get_all_cells():
            row, column = worksheet.cells.coordinate_from_string(ref)
            if row not in loaded_rows:
                continue
            style = worksheet.get_effective_style(row=row, column=column)
            if getattr(
                style.font, 'vertical_alignment', 'baseline'
            ) in ('superscript', 'subscript'):
                aligned_rows.add(row)
        return aligned_rows

    @staticmethod
    def _loaded_height_scale(
        worksheet, default_height, has_font_aligned_rows
    ):
        if (
            getattr(worksheet.page_setup, '_source_present', False)
            or not has_font_aligned_rows
        ):
            return 1.0
        source_height = float(
            worksheet.properties.format.default_row_height or 0.0
        )
        if source_height <= 0:
            return 1.0
        return min(1.0, float(default_height) / source_height)

    @classmethod
    def column_width_to_points(cls, worksheet, column):
        """Convert an Excel character width to points using the Normal font MDW."""
        if column in worksheet._hidden_columns:
            return 0.0
        width = worksheet.cells.get_column_width(column - 1)
        workbook = getattr(worksheet, "_workbook", None)
        if workbook is not None:
            font = workbook._styles[0].font
            mdw_points = workbook.text_measurer.measure_mdw(font)
            resolution = workbook.font_strategy.resolve(font)
        else:
            style = worksheet.get_effective_style(row=1, column=column)
            mdw_points = worksheet.measure_text("0", style=style).mdw_points
            resolution = worksheet.measure_text("0", style=style).font_resolution
        properties = worksheet.properties.format
        is_dengxian = resolution.resolved_family.lower() in (
            '\u7b49\u7ebf', 'dengxian'
        )
        is_calibri = resolution.resolved_family.lower() == 'calibri'
        if (
            is_calibri
            and properties.default_col_width is None
            and column not in worksheet._column_widths
        ):
            # baseColWidth uses Excel's integer device-pixel MDW path rather
            # than the explicit <col width> conversion below. Pillow's hinted
            # Calibri metric rounds 7.4px down to 7px, so retain the next pixel
            # before adding Excel's default three-pixel internal allowance.
            mdw_pixels = math.ceil(mdw_points * 96.0 / 72.0 + 1e-6)
            pixels = math.trunc(
                float(properties.base_col_width) * mdw_pixels
            ) + 3.0
            return pixels * 72.0 / 96.0
        if (
            is_calibri
            and column in getattr(worksheet, '_best_fit_columns', ())
        ):
            measure_unhinted = getattr(
                getattr(workbook, 'text_measurer', None),
                'measure_unhinted_mdw',
                None,
            )
            print_mdw_points = (
                measure_unhinted(font) if callable(measure_unhinted) else None
            )
            if print_mdw_points is not None:
                print_mdw_pixels = max(
                    print_mdw_points * 96.0 / 72.0, 0.01
                )
                pixels = (
                    (
                        256.0 * float(width)
                        + math.trunc(128.0 / print_mdw_pixels)
                    )
                    / 256.0
                    * print_mdw_pixels
                    + 1.0
                )
                return pixels * 72.0 / 96.0
        if is_dengxian and properties.default_col_width is not None:
            # Explicit default widths retain Excel's source character metric.
            return float(width) * 5.76
        if is_dengxian and column in worksheet._column_widths:
            return float(width) * 5.76
        if (
            is_dengxian
            and column not in worksheet._column_widths
            and properties.default_col_width is None
        ):
            return float(properties.base_col_width) * mdw_points + 3.24
        mdw_pixels = max(mdw_points * 96.0 / 72.0, 0.01)
        pixels = math.trunc(
            ((256.0 * float(width) + math.trunc(128.0 / mdw_pixels)) / 256.0)
            * mdw_pixels
        )
        # Excel adds 2px padding on each side plus a 1px gridline allowance.
        pixels += 5.0
        return pixels * 72.0 / 96.0

    @classmethod
    def _cell_height(cls, worksheet, ref, width_points):
        cell = worksheet.cells.get_all_cells()[ref]
        row, column = worksheet.cells.coordinate_from_string(ref)
        style = worksheet.get_effective_style(
            row=row,
            column=column,
            include_conditional_formats=True,
        )
        text = cell.get_display_text()
        if text == "":
            return 0.0

        available_width = cls.available_text_width(
            worksheet, style, width_points
        )
        measured = worksheet.measure_text(
            text,
            style=style,
            width_points=available_width,
            wrap_text=cls.should_wrap(style),
        )
        text_height = cls._rotated_height(measured, style.alignment.text_rotation)
        if (
            style.alignment.shrink_to_fit
            and not cls.should_wrap(style)
            and measured.width_points > available_width
        ):
            text_height *= available_width / measured.width_points
        padding = cls.CELL_VERTICAL_PADDING
        if (
            not cls.should_wrap(style)
            and worksheet.properties.format.dy_descent is not None
            and measured.font_resolution.resolved_family.lower() == 'calibri'
        ):
            padding = 0.0
        return text_height + padding

    @classmethod
    def available_text_width(cls, worksheet, style, width_points):
        """Return the Excel text box width after padding and indentation."""
        indent_width = cls.indent_width_points(worksheet, style)
        if style.alignment.horizontal in ('distributed', 'justify'):
            indent_width *= 2.0
        return max(
            0.1, width_points - cls.CELL_HORIZONTAL_PADDING - indent_width
        )

    @staticmethod
    def indent_width_points(worksheet, style):
        """Convert an Excel indentation level to a point offset."""
        indent = max(0, int(getattr(style.alignment, "indent", 0) or 0))
        if not indent:
            return 0.0
        mdw = worksheet.measure_text("0", style=style).mdw_points
        return indent * (mdw + 2.0)

    @staticmethod
    def should_wrap(style):
        alignment = style.alignment
        return bool(
            alignment.wrap_text
            or (
                alignment.indent
                and alignment.horizontal in ('distributed', 'justify')
            )
        )

    @staticmethod
    def _rotated_height(measured, rotation):
        rotation = int(rotation or 0)
        if rotation == 255:
            character_count = max(
                (len(part) for part in measured.text.split()), default=1
            )
            return character_count * measured.line_height_points
        if rotation:
            angle = math.radians(rotation if rotation <= 90 else rotation - 90)
            return (
                abs(measured.width_points * math.sin(angle))
                + abs(measured.height_points * math.cos(angle))
            )
        return measured.height_points

    @classmethod
    def _range_width_points(cls, worksheet, min_col, max_col):
        return sum(
            cls.column_width_to_points(worksheet, column)
            for column in range(min_col, max_col + 1)
        )

    @staticmethod
    def _merge_map(worksheet):
        origins = {}
        covered = set()
        for merge_ref in getattr(worksheet, "_merged_cells", ()):
            start_ref, _, end_ref = merge_ref.partition(":")
            end_ref = end_ref or start_ref
            min_row, min_col = worksheet.cells.coordinate_from_string(start_ref)
            max_row, max_col = worksheet.cells.coordinate_from_string(end_ref)
            origins[(min_row, min_col)] = (min_row, min_col, max_row, max_col)
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    if (row, column) != (min_row, min_col):
                        covered.add((row, column))
        return origins, covered
