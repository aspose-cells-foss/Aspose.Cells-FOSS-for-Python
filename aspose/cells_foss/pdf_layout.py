"""Backend-neutral worksheet layout and pagination for PDF export."""

import copy
import math
import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace
from datetime import datetime

from .chart import ChartType
from .render_bounds import RenderBounds
from .row_height import RowHeightCalculator
from .shape import ShapeFont
from .style_resolver import StyleResolver


@dataclass(frozen=True)
class PdfTextRunLayout:
    text: str
    width: float
    font_resolution: object
    font_size: float = 11.0
    color: str = 'FF000000'
    underline: bool = False
    underline_type: str = 'none'
    strikethrough: bool = False
    baseline_shift: float = 0.0
    advance_adjustment: float = 0.0
    synthetic_italic: bool = False


@dataclass(frozen=True)
class PdfDataBarLayout:
    x: float
    y: float
    width: float
    height: float
    gradient_start_x: float
    gradient_end_x: float
    color: str
    gradient: bool = True
    show_value: bool = True


@dataclass(frozen=True)
class PdfIconSetLayout:
    x: float
    y: float
    width: float
    height: float
    icon_set_type: str
    icon_index: int
    show_value: bool = True


@dataclass(frozen=True)
class PdfSparklineLayout:
    kind: str
    x: float
    y: float
    width: float
    height: float
    values: tuple
    color_series: str
    color_negative: str
    color_axis: str
    empty_cells: str
    line_width: float = 0.75


@dataclass(frozen=True)
class PdfCellLayout:
    row: int
    column: int
    x: float
    y: float
    width: float
    height: float
    text: str
    lines: tuple
    line_runs: tuple
    style: object
    font_resolution: object
    font_scale: float
    line_height: float
    indent_width: float
    text_clip_width: float
    numeric: bool
    rich_text: bool = False
    hyperlink: str | None = None
    merged: bool = False
    data_bar: PdfDataBarLayout | None = None
    icon_set: PdfIconSetLayout | None = None


@dataclass(frozen=True)
class PdfPictureLayout:
    name: str
    image_bytes: bytes
    image_extension: str
    x: float
    y: float
    width: float
    height: float


@dataclass(frozen=True)
class PdfShapeLayout:
    name: str
    preset_geometry: str
    is_connector: bool
    rotation: float
    flip_horizontal: bool
    flip_vertical: bool
    tail_end: str
    x: float
    y: float
    width: float
    height: float
    fill_visible: bool
    fill_color: str
    line_visible: bool
    line_color: str
    line_width: float
    text: str
    text_wrapped: bool
    text_direction: str
    text_horizontal_alignment: object
    text_vertical_alignment: object
    font: object
    font_resolution: object
    connector_start: tuple
    connector_end: tuple


@dataclass(frozen=True)
class PdfChartSeriesLayout:
    name: str
    values: tuple
    color: str
    line_width: float
    line_cap: str
    marker_symbol: str
    marker_size: float
    marker_fill_color: str | None
    marker_border_color: str | None
    marker_border_width: float
    display_values: tuple
    error_bars: tuple
    pattern_type: str | None = None
    pattern_foreground: str | None = None
    pattern_background: str | None = None
    hidden: bool = False
    x_values: tuple = ()
    line_visible: bool = True
    bubble_sizes: tuple = ()
    bubble_3d: bool = False


@dataclass(frozen=True)
class PdfChartLayout:
    name: str
    chart_kind: str
    bar_direction: str
    bar_shape: str
    gap_width: float
    overlap: float
    show_data_labels: bool
    data_label_position: str
    data_label_color: str
    x: float
    y: float
    width: float
    height: float
    background_fill_color: str | None
    background_gradient_stops: tuple
    background_gradient_angle: float
    background_pattern_type: str | None
    background_pattern_foreground: str
    background_pattern_background: str
    chart_border_color: str | None
    chart_border_width: float
    back_wall_fill_image: bytes | None
    side_wall_fill_image: bytes | None
    plot_x: float
    plot_y: float
    plot_width: float
    plot_height: float
    title: str
    title_layout_x: float | None
    title_layout_y: float | None
    title_color: str
    title_fill_image: bytes | None
    title_border_color: str | None
    categories: tuple
    category_positions: tuple
    axis_categories: tuple
    category_rotation: float
    show_vertical_gridlines: bool
    category_major_gridline_color: str
    value_major_gridline_color: str
    plot_border_color: str | None
    show_minor_horizontal_gridlines: bool
    minor_gridline_color: str
    value_axis_color: str | None
    value_axis_position: str
    value_major_tick_mark: str
    value_minor_tick_mark: str
    value_axis_text_color: str
    value_axis_fill_image: bytes | None
    category_axis_text_color: str
    category_axis_fill_image: bytes | None
    category_axis_title: str
    value_axis_title: str
    legend_text_color: str
    legend_fill_image: bytes | None
    legend_layout_x: float | None
    legend_layout_y: float | None
    is_3d: bool
    rotation_x: float
    rotation_y: float
    depth_percent: float
    perspective: float
    series: tuple
    grouping: str
    value_min: float
    value_max: float
    value_step: float
    value_axis_log_base: float | None
    axis_number_format: str
    value_axis_display_unit: float
    value_axis_display_unit_label: str
    value_axis_display_unit_layout_x: float | None
    value_axis_display_unit_layout_y: float | None
    show_legend: bool
    show_data_table: bool
    data_table_show_keys: bool
    show_series_axis: bool
    title_font_resolution: object
    axis_font_resolution: object
    pie_point_styles: tuple = ()
    pie_series_point_styles: tuple = ()
    pie_first_slice_angle: float = 0.0
    pie_variant: str = 'pie'
    pie_hole_size: float = 0.0
    pie_secondary_size: float = 75.0
    pie_secondary_indices: tuple = ()
    pie_other_style: tuple = ('0D3A4E', None, 0.0)
    pie_connector_color: str = 'A6A6A6'
    pie_connector_width: float = 0.75
    title_gradient_stops: tuple = ()
    title_gradient_angle: float = 0.0
    legend_gradient_stops: tuple = ()
    legend_gradient_angle: float = 0.0
    chart_glow_color: str | None = None
    chart_glow_opacity: float = 0.0
    chart_glow_radius: float = 0.0
    chart_soft_edge_radius: float = 0.0
    chart_bevel_top_width: float = 0.0
    radar_style: str = 'standard'
    legend_position: str = 'b'
    box_show_mean_line: bool = False
    box_show_mean_marker: bool = True
    box_show_inner_points: bool = False
    box_show_outlier_points: bool = True
    box_gap_width: float = 1.0
    box_quartile_method: str = 'exclusive'
    histogram_variant: str = 'histogram'
    histogram_bins: tuple = ()
    map_value_colors: tuple = ()
    stock_style: str = 'high_low_close'
    stock_has_high_low_lines: bool = False
    stock_has_up_down_bars: bool = False
    stock_has_volume: bool = False
    stock_axes_hidden: bool = False
    rounded_corners: bool = False
    stock_price_min: float = 0.0
    stock_price_max: float = 1.0
    stock_price_step: float = 1.0
    sunburst_paths: tuple = ()
    treemap_paths: tuple = ()
    waterfall_subtotals: tuple = ()
    waterfall_gap_width: float = 0.35
    scatter_style: str = 'lineMarker'
    scatter_x_min: float = 0.0
    scatter_x_max: float = 1.0
    scatter_x_step: float = 1.0
    scatter_x_number_format: str = '0'
    scatter_text_outline_gradient_stops: tuple = ()
    scatter_text_outline_gradient_angle: float = 0.0
    scatter_vary_colors: bool = False
    bubble_scale: float = 100.0
    surface_wireframe: bool = False
    surface_band_colors: tuple = ()
    surface_axes_hidden: bool = False


@dataclass(frozen=True)
class PdfPageLayout:
    worksheet_name: str
    page_width: float
    page_height: float
    content_left: float
    content_top: float
    content_width: float
    content_height: float
    primary_bounds: RenderBounds
    repeated_rows: tuple
    repeated_columns: tuple
    scale: float
    cells: tuple
    sparklines: tuple
    pictures: tuple
    shapes: tuple
    charts: tuple
    render_gridlines: bool


class PdfLayoutEngine:
    """Materializes all worksheet pages before the renderer starts drawing."""

    _DATA_BAR_HORIZONTAL_INSET = 1.44
    _DATA_BAR_TOP_INSET = 1.44
    _DATA_BAR_BOTTOM_INSET = 2.40
    _ICON_SET_HORIZONTAL_INSET = 1.44
    _ICON_SET_TOP_INSET = 2.16
    _ICON_SET_BOTTOM_INSET = 0.36
    _ICON_SET_SIZE = 12.0

    PAPER_SIZES = {
        1: (612.0, 792.0),       # Letter
        5: (612.0, 1008.0),      # Legal
        8: (841.89, 1190.55),    # A3
        9: (595.28, 841.89),     # A4
        11: (419.53, 595.28),    # A5
    }

    @classmethod
    def build_pages(cls, worksheet, options):
        areas = worksheet.get_printable_bounds(include_drawings=True)
        if not areas:
            areas = (RenderBounds(1, 1, 1, 1),)
        elif not getattr(worksheet, 'print_area', None):
            areas = tuple(
                RenderBounds(1, 1, area.max_row, area.max_col)
                for area in areas
            )

        page_width, page_height = cls._page_size(worksheet, options)
        margins = worksheet.page_margins
        left = float(margins.left) * 72.0
        right = float(margins.right) * 72.0
        top = float(margins.top) * 72.0
        bottom = float(margins.bottom) * 72.0
        usable_width = page_width - left - right
        usable_height = page_height - top - bottom
        if usable_width <= 0 or usable_height <= 0:
            raise ValueError("page margins leave no printable PDF area")

        pages = []
        for area in areas:
            pages.extend(
                cls._build_area_pages(
                    worksheet,
                    area,
                    options,
                    page_width,
                    page_height,
                    left,
                    top,
                    usable_width,
                    usable_height,
                )
            )
        return pages

    @classmethod
    def _build_area_pages(
        cls, worksheet, area, options, page_width, page_height,
        left, top, usable_width, usable_height,
    ):
        title_rows, title_columns = worksheet.get_print_title_ranges()
        measure_bounds = cls._union_titles(area, title_rows, title_columns)
        row_heights = worksheet.calculate_row_heights(measure_bounds)
        column_widths = {
            column: RowHeightCalculator.column_width_to_points(worksheet, column)
            for column in range(measure_bounds.min_col, measure_bounds.max_col + 1)
        }
        map_chart_sheet = any(
            chart.type == ChartType.MAP for chart in worksheet.charts
        )
        if map_chart_sheet:
            # Modern geography-chart templates retain printer-grid metrics
            # that differ from the chartEx drawing canvas. Preserve the source
            # row dimensions and calibrate explicit column widths separately;
            # drawing objects use their own EMU mapping below.
            source_default_height = float(
                worksheet.properties.format.default_row_height
            )
            row_heights = {
                row: (
                    0.0 if row in worksheet._hidden_rows else
                    float(worksheet._row_heights.get(
                        row, source_default_height
                    )) * 0.994
                )
                for row in row_heights
            }

        scale = cls._resolve_scale(
            worksheet,
            area,
            row_heights,
            column_widths,
            usable_width,
            usable_height,
            options,
        )
        row_chunks = cls._split_axis(
            area.min_row,
            area.max_row,
            row_heights,
            usable_height / scale,
            title_rows,
            getattr(worksheet, '_horizontal_page_breaks', set()),
        )
        column_chunks = cls._split_axis(
            area.min_col,
            area.max_col,
            column_widths,
            usable_width / scale,
            title_columns,
            getattr(worksheet, '_vertical_page_breaks', set()),
        )

        # Geography charts use a drawing transform whose physical bounds fit
        # on one page even when its two-cell anchor extends into later rows.
        # Do not turn that anchor-only tail into a second, blank vertical page.
        if map_chart_sheet:
            row_chunks = ((area.min_row, area.max_row),)

        plans = []
        for row_index, row_chunk in enumerate(row_chunks):
            for column_index, column_chunk in enumerate(column_chunks):
                repeated_rows = cls._repeat_indexes(title_rows, row_chunk, row_index)
                repeated_columns = cls._repeat_indexes(
                    title_columns, column_chunk, column_index
                )
                plans.append((row_index, column_index, row_chunk, column_chunk,
                              repeated_rows, repeated_columns))

        if worksheet.page_setup.page_order == 'downThenOver':
            plans.sort(key=lambda item: (item[1], item[0]))

        return [
            cls._materialize_page(
                worksheet,
                options,
                page_width,
                page_height,
                left,
                top,
                left + usable_width,
                scale,
                row_heights,
                column_widths,
                *plan[2:],
            )
            for plan in plans
        ]

    @classmethod
    def _materialize_page(
        cls, worksheet, options, page_width, page_height, left, top,
        printable_right, scale, row_heights, column_widths, row_chunk, column_chunk,
        repeated_rows, repeated_columns,
    ):
        drawing_column_widths = column_widths
        map_chart_sheet = any(
            chart.type == ChartType.MAP for chart in worksheet.charts
        )
        if map_chart_sheet:
            column_widths = {
                column: width * (0.9803 if column == 1 else 0.9189)
                for column, width in column_widths.items()
            }
        rows = repeated_rows + tuple(
            row for row in range(row_chunk[0], row_chunk[1] + 1)
            if row not in repeated_rows
        )
        columns = repeated_columns + tuple(
            column for column in range(column_chunk[0], column_chunk[1] + 1)
            if column not in repeated_columns
        )
        x_positions = cls._positions(columns, column_widths, left, scale)
        cell_top = top
        if any(chart.type == ChartType.MAP for chart in worksheet.charts):
            cell_top += 41.28 * scale
        y_positions = cls._positions(rows, row_heights, cell_top, scale)
        cells = []
        all_cells = worksheet.cells.get_all_cells()
        covered, merges = cls._merge_map(worksheet)
        hyperlinks = cls._hyperlink_map(worksheet)
        for ref, cell in all_cells.items():
            row, column = worksheet.cells.coordinate_from_string(ref)
            if row not in y_positions or column not in x_positions:
                continue
            if (row, column) in covered:
                continue
            merge = merges.get((row, column))
            merge_rows = [row]
            merge_columns = [column]
            if merge:
                merge_rows = [value for value in rows if merge[0] <= value <= merge[2]]
                merge_columns = [value for value in columns if merge[1] <= value <= merge[3]]
            width = sum(column_widths[value] for value in merge_columns) * scale
            height = sum(row_heights[value] for value in merge_rows) * scale
            style = worksheet.get_effective_style(
                row=row,
                column=column,
                include_conditional_formats=True,
            )
            text = cell.get_display_text()
            available_width = RowHeightCalculator.available_text_width(
                worksheet, style, width / scale
            )
            should_wrap = RowHeightCalculator.should_wrap(style)
            measured = worksheet.measure_text(
                text,
                style=style,
                width_points=available_width,
                wrap_text=should_wrap,
            )
            rich_text = bool(cell.rich_text_runs) and not should_wrap and not (
                '\n' in text
                or int(style.alignment.text_rotation or 0)
            )
            if rich_text:
                line_runs, text_width, line_height = cls._rich_line_runs(
                    worksheet, cell.rich_text_runs, style, scale
                )
                lines = (text,)
            else:
                line_runs, run_text_width = cls._line_runs(
                    worksheet, measured.lines, style, scale
                )
                lines = measured.lines
                text_width = run_text_width
                line_height = measured.line_height_points
            font_scale = 1.0
            if (
                style.alignment.shrink_to_fit
                and not should_wrap
                and text_width > available_width
            ):
                font_scale = max(0.1, available_width / text_width)
            indent_width = RowHeightCalculator.indent_width_points(
                worksheet, style
            ) * scale
            text_clip_width = cls._text_clip_width(
                worksheet=worksheet,
                row=row,
                column=column,
                merge_columns=merge_columns,
                page_columns=columns,
                column_widths=column_widths,
                scale=scale,
                width=width,
                measured_width=text_width * scale * font_scale,
                indent_width=indent_width,
                max_clip_width=max(width, printable_right - x_positions[column]),
                style=style,
                all_cells=all_cells,
                covered=covered,
                merges=merges,
            )
            data_bar = cls._materialize_data_bar(
                worksheet=worksheet,
                row=row,
                column=column,
                x=x_positions[column],
                y=y_positions[row],
                width=width,
                height=height,
                scale=scale,
            )
            icon_set = cls._materialize_icon_set(
                worksheet=worksheet,
                row=row,
                column=column,
                x=x_positions[column],
                y=y_positions[row],
                width=width,
                height=height,
                scale=scale,
            )
            cells.append(PdfCellLayout(
                row=row,
                column=column,
                x=x_positions[column],
                y=y_positions[row],
                width=width,
                height=height,
                text=text,
                lines=lines,
                line_runs=line_runs,
                style=style,
                font_resolution=measured.font_resolution,
                font_scale=font_scale,
                line_height=line_height * scale * font_scale,
                indent_width=indent_width,
                text_clip_width=text_clip_width,
                numeric=isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool),
                rich_text=rich_text,
                hyperlink=hyperlinks.get((row, column)),
                merged=bool(merge),
                data_bar=data_bar,
                icon_set=icon_set,
            ))

        cells.sort(key=lambda item: (item.y, item.x, item.row, item.column))
        sparklines = cls._materialize_sparklines(
            worksheet,
            x_positions,
            y_positions,
            row_heights,
            column_widths,
            scale,
        )
        pictures = cls._materialize_pictures(
            worksheet,
            page_width,
            page_height,
            left,
            top,
            scale,
            row_heights,
            drawing_column_widths,
            row_chunk,
            column_chunk,
        )
        shapes = cls._materialize_shapes(
            worksheet,
            page_width,
            page_height,
            left,
            top,
            scale,
            row_heights,
            drawing_column_widths,
            row_chunk,
            column_chunk,
        )
        charts = cls._materialize_charts(
            worksheet,
            left,
            top,
            scale,
            row_heights,
            drawing_column_widths,
            row_chunk,
            column_chunk,
        )
        stock_chart_page = bool(charts) and all(
            chart.chart_kind == 'stock' for chart in charts
        )
        content_width = sum(
            column_widths[column] for column in columns
        ) * scale
        content_height = sum(row_heights[row] for row in rows) * scale
        if stock_chart_page:
            content_width = printable_right - left
            content_height = (
                page_height - top
                - float(worksheet.page_margins.bottom) * 72.0
            )
        return PdfPageLayout(
            worksheet_name=worksheet.name,
            page_width=page_width,
            page_height=page_height,
            content_left=left,
            content_top=top,
            content_width=content_width,
            content_height=content_height,
            primary_bounds=RenderBounds(
                row_chunk[0], column_chunk[0], row_chunk[1], column_chunk[1]
            ),
            repeated_rows=repeated_rows,
            repeated_columns=repeated_columns,
            scale=scale,
            cells=tuple(cells),
            sparklines=sparklines,
            pictures=pictures,
            shapes=shapes,
            charts=charts,
            render_gridlines=(
                bool(worksheet.properties.print_options.print_grid_lines)
                if options.render_gridlines is None
                else bool(options.render_gridlines)
            ),
        )

    @classmethod
    def _materialize_sparklines(
        cls, worksheet, x_positions, y_positions,
        row_heights, column_widths, scale,
    ):
        type_names = {0: 'line', 1: 'column', 2: 'win_loss'}
        empty_names = {0: 'zero', 1: 'gap', 2: 'connected'}
        layouts = []
        for group in getattr(worksheet, 'sparkline_groups', ()):
            for sparkline in group.sparklines:
                try:
                    row, column = worksheet.cells.coordinate_from_string(
                        sparkline.cell_reference.replace('$', '')
                    )
                except ValueError:
                    continue
                if row not in y_positions or column not in x_positions:
                    continue
                values = cls._sparkline_range_values(
                    worksheet, sparkline.data_range
                )
                empty_cells = empty_names.get(
                    int(group.display_empty_cells_as), 'gap'
                )
                if empty_cells == 'zero':
                    values = tuple(
                        0.0 if value is None else value for value in values
                    )
                layouts.append(PdfSparklineLayout(
                    kind=type_names.get(int(group.type), 'line'),
                    x=x_positions[column],
                    y=y_positions[row],
                    width=column_widths[column] * scale,
                    height=row_heights[row] * scale,
                    values=values,
                    color_series=group.color_series,
                    color_negative=group.color_negative,
                    color_axis=group.color_axis,
                    empty_cells=empty_cells,
                ))
        layouts.sort(key=lambda item: (item.y, item.x, item.kind))
        return tuple(layouts)

    @staticmethod
    def _sparkline_range_values(worksheet, formula):
        if not formula:
            return ()
        source = worksheet
        reference = str(formula).strip()
        if '!' in reference:
            sheet_name, reference = reference.rsplit('!', 1)
            sheet_name = sheet_name.strip().strip("'")
            workbook = getattr(worksheet, '_workbook', None)
            if workbook is not None and sheet_name != worksheet.name:
                try:
                    source = workbook.get_worksheet(sheet_name)
                except (IndexError, ValueError):
                    return ()
        reference = reference.replace('$', '')
        start, _, end = reference.partition(':')
        try:
            start_row, start_column = source.cells.coordinate_from_string(start)
            end_row, end_column = source.cells.coordinate_from_string(
                end or start
            )
        except ValueError:
            return ()
        cells = source.cells.get_all_cells()
        values = []
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for column in range(
                min(start_column, end_column),
                max(start_column, end_column) + 1,
            ):
                ref = source.cells.coordinate_to_string(row, column)
                cell = cells.get(ref)
                if cell is None or cell.value is None or cell.value == '':
                    values.append(None)
                elif (
                    isinstance(cell.value, (int, float))
                    and not isinstance(cell.value, bool)
                ):
                    values.append(float(cell.value))
                else:
                    # Excel treats text inside a sparkline source as zero.
                    values.append(0.0)
        return tuple(values)

    @classmethod
    def _materialize_data_bar(
        cls, worksheet, row, column, x, y, width, height, scale
    ):
        resolved = StyleResolver.resolve_data_bar(
            worksheet, row, column
        )
        if resolved is None:
            return None

        horizontal_inset = cls._DATA_BAR_HORIZONTAL_INSET * scale
        top_inset = cls._DATA_BAR_TOP_INSET * scale
        bottom_inset = cls._DATA_BAR_BOTTOM_INSET * scale
        available_width = max(0.0, width - 2.0 * horizontal_inset)
        bar_width = available_width * resolved['length_ratio']
        bar_height = max(0.0, height - top_inset - bottom_inset)
        if bar_width <= 0.0 or bar_height <= 0.0:
            return None

        inner_left = x + horizontal_inset
        if resolved['direction'] == 'right-to-left':
            bar_x = x + width - horizontal_inset - bar_width
            gradient_start_x = x + width - horizontal_inset
            gradient_end_x = gradient_start_x - width
        else:
            bar_x = inner_left
            gradient_start_x = inner_left
            gradient_end_x = gradient_start_x + width
        return PdfDataBarLayout(
            x=bar_x,
            y=y + top_inset,
            width=bar_width,
            height=bar_height,
            gradient_start_x=gradient_start_x,
            gradient_end_x=gradient_end_x,
            color=resolved['color'],
            gradient=resolved['gradient'],
            show_value=resolved['show_value'],
        )

    @classmethod
    def _materialize_icon_set(
        cls, worksheet, row, column, x, y, width, height, scale
    ):
        resolved = StyleResolver.resolve_icon_set(
            worksheet, row, column
        )
        if resolved is None:
            return None

        horizontal_inset = cls._ICON_SET_HORIZONTAL_INSET * scale
        top_inset = cls._ICON_SET_TOP_INSET * scale
        bottom_inset = cls._ICON_SET_BOTTOM_INSET * scale
        size = min(
            cls._ICON_SET_SIZE * scale,
            max(0.0, width - 2.0 * horizontal_inset),
            max(0.0, height - top_inset - bottom_inset),
        )
        if size <= 0.0:
            return None
        return PdfIconSetLayout(
            x=x + horizontal_inset,
            y=y + top_inset,
            width=size,
            height=size,
            icon_set_type=resolved['icon_set_type'],
            icon_index=resolved['icon_index'],
            show_value=resolved['show_value'],
        )

    @classmethod
    def _materialize_pictures(
        cls, worksheet, page_width, page_height, left, top, scale,
        row_heights, column_widths, row_chunk, column_chunk,
    ):
        if not worksheet.pictures:
            return ()
        default_row_height = worksheet.get_effective_default_row_height()

        def column_width(column):
            return column_widths.get(
                column,
                RowHeightCalculator.column_width_to_points(
                    worksheet, column
                ),
            )

        def row_height(row):
            if row in row_heights:
                return row_heights[row]
            if row in worksheet._hidden_rows:
                return 0.0
            return float(worksheet._row_heights.get(row, default_row_height))

        column_origin = sum(
            column_width(column)
            for column in range(1, column_chunk[0])
        )
        row_origin = sum(
            row_height(row)
            for row in range(1, row_chunk[0])
        )
        layouts = []
        for picture in worksheet.pictures:
            if getattr(picture, '_transform_is_group_absolute', False):
                # Group children use coordinates in the group's child space.
                # The XML loader has composed that space into drawing-absolute
                # EMUs, so retain the physical aspect ratio here instead of
                # independently stretching it to the cell anchor on each axis.
                wide_map_table = column_width(3) > 0.0
                base_x = 34.6 if wide_map_table else 42.0
                page_origins = (
                    (0.0, 482.76, 1013.88)
                    if wide_map_table else (0.0, 491.03, 996.58)
                )
                page_index = 0 if column_chunk[0] <= 1 else (
                    1 if column_chunk[0] <= 12 else 2
                )
                x = (
                    picture._transform_x / 12700.0
                    + base_x - page_origins[page_index]
                ) * scale
                y = (
                    picture._transform_y / 12700.0 + 93.5
                ) * scale
                width = (
                    picture._transform_width / 12700.0 * 0.9592 * scale
                )
                height = (
                    picture._transform_height / 12700.0 * 0.994 * scale
                )
                if (
                    x >= page_width
                    or y >= page_height
                    or x + width <= 0
                    or y + height <= 0
                ):
                    continue
                layouts.append(PdfPictureLayout(
                    name=picture.name,
                    image_bytes=picture.image_bytes,
                    image_extension=picture.image_extension,
                    x=x,
                    y=y,
                    width=width,
                    height=height,
                ))
                continue
            picture_left = sum(
                column_width(column)
                for column in range(1, picture._upper_left_column + 1)
            ) + picture._upper_left_column_offset / 12700.0
            picture_right = sum(
                column_width(column)
                for column in range(1, picture._lower_right_column + 1)
            ) + picture._lower_right_column_offset / 12700.0
            picture_top = sum(
                row_height(row)
                for row in range(1, picture._upper_left_row + 1)
            ) + picture._upper_left_row_offset / 12700.0
            picture_bottom = sum(
                row_height(row)
                for row in range(1, picture._lower_right_row + 1)
            ) + picture._lower_right_row_offset / 12700.0

            transform_x = picture._transform_x
            transform_y = picture._transform_y
            transform_width = picture._transform_width
            transform_height = picture._transform_height
            if transform_x and transform_width:
                horizontal_ratio = picture_left / (transform_x / 12700.0)
                picture_left = transform_x / 12700.0 * horizontal_ratio
                picture_right = (
                    transform_x + transform_width
                ) / 12700.0 * horizontal_ratio
            if transform_y and transform_height:
                vertical_ratio = picture_top / (transform_y / 12700.0)
                picture_top = transform_y / 12700.0 * vertical_ratio
                picture_bottom = (
                    transform_y + transform_height
                ) / 12700.0 * vertical_ratio

            x = left + (picture_left - column_origin) * scale
            y = top + (picture_top - row_origin) * scale
            width = (picture_right - picture_left) * scale
            height = (picture_bottom - picture_top) * scale
            if (
                x >= page_width
                or y >= page_height
                or x + width <= 0
                or y + height <= 0
            ):
                continue
            layouts.append(PdfPictureLayout(
                name=picture.name,
                image_bytes=picture.image_bytes,
                image_extension=picture.image_extension,
                x=x,
                y=y,
                width=width,
                height=height,
            ))
        return tuple(layouts)

    @classmethod
    def _materialize_charts(
        cls, worksheet, left, top, scale, row_heights, column_widths,
        row_chunk, column_chunk,
    ):
        if not worksheet.charts:
            return ()
        default_row_height = worksheet.get_effective_default_row_height()

        def column_width(column):
            return column_widths.get(
                column,
                RowHeightCalculator.column_width_to_points(worksheet, column),
            )

        def row_height(row):
            return row_heights.get(row, default_row_height)

        column_origin = sum(
            column_width(column) for column in range(1, column_chunk[0])
        )
        row_origin = sum(
            row_height(row) for row in range(1, row_chunk[0])
        )
        clip_right = left + sum(
            column_width(column)
            for column in range(column_chunk[0], column_chunk[1] + 1)
        ) * scale
        clip_bottom = top + sum(
            row_height(row)
            for row in range(row_chunk[0], row_chunk[1] + 1)
        ) * scale
        workbook = getattr(worksheet, '_workbook', None)
        axis_font = ShapeFont()
        axis_font.name = 'Aptos Narrow'
        axis_font.size = 9.0
        title_font = axis_font.copy()
        title_font.size = 14.0
        axis_resolution = workbook.font_strategy.resolve(axis_font)
        default_title_resolution = workbook.font_strategy.resolve(title_font)
        layouts = []
        for index, chart in enumerate(worksheet.charts):
            chart_kind = 'line'
            if chart.type == ChartType.BAR:
                chart_kind = (
                    'column' if chart.bar_direction == 'col' else 'bar'
                )
            elif chart.type == ChartType.PIE:
                chart_kind = 'pie'
            elif chart.type == ChartType.RADAR:
                chart_kind = 'radar'
            elif chart.type == ChartType.AREA:
                chart_kind = 'area'
            elif chart.type == ChartType.BOX_WHISKER:
                chart_kind = 'box_whisker'
            elif chart.type == ChartType.FUNNEL:
                chart_kind = 'funnel'
            elif chart.type == ChartType.HISTOGRAM:
                chart_kind = 'histogram'
            elif chart.type == ChartType.MAP:
                chart_kind = 'map'
            elif chart.type == ChartType.STOCK:
                chart_kind = 'stock'
            elif chart.type == ChartType.SUNBURST:
                chart_kind = 'sunburst'
            elif chart.type == ChartType.TREEMAP:
                chart_kind = 'treemap'
            elif chart.type == ChartType.WATERFALL:
                chart_kind = 'waterfall'
            elif chart.type == ChartType.SCATTER:
                chart_kind = (
                    'bubble' if getattr(chart, '_is_bubble', False)
                    else 'scatter'
                )
            elif chart.type == ChartType.SURFACE:
                chart_kind = 'surface'
            waterfall_subtotals = (
                cls._chart_waterfall_subtotals(chart)
                if chart_kind == 'waterfall' else ()
            )
            waterfall_gap_width = (
                cls._chart_waterfall_gap_width(chart)
                if chart_kind == 'waterfall' else 0.35
            )
            if chart_kind == 'stock' and not (
                row_chunk[0] <= chart._lower_right_row + 1 <= row_chunk[1]
                and column_chunk[0]
                <= chart._lower_right_column + 1 <= column_chunk[1]
            ):
                continue
            if chart_kind == 'map':
                anchor_center = (
                    chart._upper_left_column + chart._lower_right_column
                ) / 2.0
                if not column_chunk[0] <= anchor_center <= column_chunk[1]:
                    continue
            histogram_variant = (
                cls._chart_histogram_variant(chart)
                if chart_kind == 'histogram' else 'histogram'
            )
            stock_features = (
                cls._chart_stock_features(chart)
                if chart_kind == 'stock'
                else (False, False, False, False, False)
            )
            surface_features = (
                cls._chart_surface_features(chart)
                if chart_kind == 'surface'
                else (False, False)
            )
            chart_left = sum(
                column_width(column)
                for column in range(1, chart._upper_left_column + 1)
            ) + chart._upper_left_column_offset / 12700.0
            chart_right = sum(
                column_width(column)
                for column in range(1, chart._lower_right_column + 1)
            ) + chart._lower_right_column_offset / 12700.0
            chart_top = sum(
                row_height(row)
                for row in range(1, chart._upper_left_row + 1)
            ) + chart._upper_left_row_offset / 12700.0
            chart_bottom = sum(
                row_height(row)
                for row in range(1, chart._lower_right_row + 1)
            ) + chart._lower_right_row_offset / 12700.0
            anchor_extent_width = getattr(
                chart, '_anchor_extent_width', None
            )
            anchor_extent_height = getattr(
                chart, '_anchor_extent_height', None
            )
            if anchor_extent_width is not None:
                chart_right = chart_left + anchor_extent_width / 12700.0
            if anchor_extent_height is not None:
                chart_bottom = chart_top + anchor_extent_height / 12700.0
            # Graphic-frame anchors use Excel's drawing pixel boundary rather
            # than the cell text metric used by the worksheet grid.
            x = left + (chart_left - column_origin) * scale + 1.2 * scale
            y = top + (chart_top - row_origin) * scale - 1.2 * scale
            width = (chart_right - chart_left) * scale
            height = max(0.0, (chart_bottom - chart_top) * scale - 0.16 * scale)
            if anchor_extent_width is not None:
                x -= 2.4 * scale
                height += 0.16 * scale
            chart_glow_color, chart_glow_opacity, chart_glow_radius = (
                cls._chart_glow_style(chart)
            )
            chart_soft_edge_radius = cls._chart_soft_edge_radius(chart)
            chart_bevel_top_width = (
                cls._chart_bevel_top_width(chart) * scale
            )
            if (
                x - chart_glow_radius >= clip_right
                or y - chart_glow_radius >= clip_bottom
                or x + width + chart_glow_radius <= left
                or y + height + chart_glow_radius <= top
            ):
                continue
            manual_layout = cls._chart_has_manual_layout(chart)
            if not manual_layout and anchor_extent_width is None:
                default_columns = chart._lower_right_column - chart._upper_left_column
                width += default_columns * 0.24 * scale
                height += 0.5 * scale
                if default_columns >= 8:
                    width -= 1.08 * scale
                if height >= 250.0:
                    height -= 0.12 * scale
                if chart_kind == 'pie' and default_columns < 7:
                    # Compact pie frames align to Excel's drawing boundary,
                    # which differs slightly from the generic chart inset.
                    x -= 0.72 * scale
                    y += 1.06 * scale
                    width += 0.13 * scale
                    height = max(0.0, height - 1.30 * scale)
            if chart_kind == 'radar':
                width = max(0.0, width - 0.96 * scale)
                height += 0.62 * scale
                if str(
                    getattr(chart, 'radar_style', 'standard')
                ).lower() in ('standard', 'filled'):
                    x -= (1.74 if x >= 0.0 else 0.93) * scale
                    y += 0.96 * scale
                    width = max(0.0, width - 1.26 * scale)
                    height = max(0.0, height - 2.68 * scale)
            if chart_kind == 'funnel' and not manual_layout:
                x += 0.06 * scale
                y += 0.06 * scale
                width = max(0.0, width - 1.08 * scale)
                height = max(0.0, height - 0.36 * scale)
            if chart_kind == 'histogram' and not manual_layout:
                if histogram_variant == 'pareto':
                    x += 0.66 * scale
                    y += 0.28 * scale
                    width = max(0.0, width - 3.225 * scale)
                    height += 0.14 * scale
                else:
                    x += 0.06 * scale
                    y += 0.06 * scale
                    width = max(0.0, width - 1.08 * scale)
                    height = max(0.0, height - 0.36 * scale)
            if chart_kind == 'map' and not manual_layout:
                # chartEx geography frames use their xfrm physical size; the
                # fallback two-cell anchor is substantially taller and wider.
                x = left + (
                    17.34
                    + (chart._upper_left_column - 12) * 31.92
                    + (
                        174.0
                        if column_width(3) > 0.0 else 0.0
                    )
                ) * scale
                y = top + (
                    114.06
                    + (chart._upper_left_row_offset - 146684) / 12700.0
                ) * scale
                width = 345.24 * scale
                height = 240.0 * scale
            if (
                chart_kind == 'surface'
                and not bool(getattr(chart, 'wireframe', False))
                and not manual_layout
            ):
                y += 0.72 * scale
                height += 0.48 * scale
                if not bool(getattr(chart, 'is_3d', False)):
                    width = max(0.0, width - 2.88 * scale)
            if (
                chart_kind == 'treemap'
                and not manual_layout
                and height >= 250.0
            ):
                x -= 1.08 * scale
                width = max(0.0, width - 1.08 * scale)
            elif (
                chart_kind == 'treemap'
                and not manual_layout
                and height >= 210.0
            ):
                x -= 0.54 * scale
                width = max(0.0, width - 1.08 * scale)
            elif (
                chart_kind == 'waterfall'
                and not manual_layout
                and chart._upper_left_column == 0
            ):
                width = max(0.0, width - 2.0 * scale)
            compact_percent_scatter = (
                chart_kind == 'scatter'
                and not bool(chart.show_legend)
                and '%' in cls._chart_scatter_y_number_format(chart)
                and anchor_extent_width is None
            )
            if compact_percent_scatter:
                x -= 2.25 * scale
                width = max(0.0, width - 9.0 * scale)
                height = max(0.0, height - 14.0 * scale)
            scatter_vary_colors = (
                chart_kind == 'scatter' and bool(chart.vary_colors)
            )
            plot = cls._chart_plot_ratios(chart, width, height)
            if chart_kind == 'box_whisker' and not manual_layout:
                # Modern chartEx box charts reserve a top legend band and a
                # narrower value-label gutter than legacy charts.
                plot = (
                    38.58 / width,
                    80.94 / height,
                    (width - 45.12) / width,
                    91.44 / height,
                )
            elif chart_kind == 'funnel' and not manual_layout:
                plot = (
                    38.52 / width,
                    59.16 / height,
                    (width - 45.24) / width,
                    (height - 65.76) / height,
                )
            elif chart_kind == 'histogram' and not manual_layout:
                if histogram_variant == 'pareto':
                    plot = (
                        45.96 / width,
                        32.64 / height,
                        (width - 85.08) / width,
                        (height - 55.68) / height,
                    )
                else:
                    plot = (
                        23.40 / width,
                        63.60 / height,
                        (width - 30.00) / width,
                        (height - 86.64) / height,
                    )
            elif chart_kind == 'treemap' and not manual_layout:
                # chartEx treemaps reserve a compact title/legend band.
                plot_top = (
                    58.56 if height >= 250.0
                    else 57.36 if height >= 210.0
                    else 56.04
                )
                bottom_reserve = 5.56 if height >= 210.0 else 6.36
                horizontal_reserve = 12.12 if height >= 250.0 else 13.20
                plot = (
                    6.60 / width,
                    plot_top / height,
                    (width - horizontal_reserve) / width,
                    (height - plot_top - bottom_reserve) / height,
                )
            elif chart_kind == 'waterfall' and not manual_layout:
                plot = (
                    37.0 / width,
                    61.0 / height,
                    (width - 43.5) / width,
                    (height - 84.5) / height,
                )
            elif chart_kind in ('scatter', 'bubble') and not manual_layout:
                if scatter_vary_colors:
                    plot = (
                        11.0 / width,
                        11.0 / height,
                        (width - 61.0) / width,
                        (height - 22.0) / height,
                    )
                elif compact_percent_scatter:
                    plot = (
                        37.0 / width,
                        39.23 / height,
                        (width - 55.0) / width,
                        (height - 60.98) / height,
                    )
                else:
                    plot = (
                        50.75 / width,
                        37.09 / height,
                        (width - 73.75) / width,
                        (height - 83.08) / height,
                    )
            elif chart_kind == 'surface' and not manual_layout:
                # A top-down contour surface reserves the right side for its
                # series axis and the bottom band for interval legend items.
                if bool(getattr(chart, 'wireframe', False)):
                    plot = (
                        70.08 / width,
                        37.92 / height,
                        218.88 / width,
                        109.44 / height,
                    )
                else:
                    plot = (
                        95.76 / width,
                        39.36 / height,
                        174.48 / width,
                        131.28 / height,
                    )
            series = []
            categories = ()
            sunburst_paths = ()
            treemap_paths = ()
            colors = ('156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E')
            series_styles = cls._chart_series_styles(chart)
            series_line_visibility = cls._chart_series_line_visibility(chart)
            series_patterns = cls._chart_series_pattern_styles(chart)
            series_names = cls._chart_series_names(chart, worksheet)
            use_default_radar_markers = (
                chart_kind == 'radar'
                and str(getattr(chart, 'radar_style', 'standard')).lower()
                == 'marker'
                and cls._chart_radar_uses_default_markers(chart)
            )
            for series_index, source_series in enumerate(chart.n_series):
                values = cls._chart_range_values(
                    worksheet, source_series.values, display=False
                )
                x_values = cls._chart_range_values(
                    worksheet, getattr(source_series, 'x_values', None),
                    display=False,
                )
                bubble_sizes = cls._chart_range_values(
                    worksheet, getattr(source_series, 'bubble_sizes', None),
                    display=False,
                )
                display_values = cls._chart_range_values(
                    worksheet, source_series.values, display=True
                )
                if not categories:
                    categories = cls._chart_range_values(
                        worksheet, source_series.category_data, display=True
                    )
                default_color = colors[series_index % len(colors)]
                (
                    series_color, line_width, line_cap,
                    marker_symbol, marker_size, marker_fill_color,
                    marker_border_color, marker_border_width,
                ) = (
                    series_styles[series_index]
                    if series_index < len(series_styles)
                    else (
                        default_color, 2.25, 'butt',
                        'none', 0.0, None, None, 0.0,
                    )
                )
                if use_default_radar_markers:
                    if marker_symbol == 'none':
                        marker_symbol = 'circle'
                    if marker_size <= 0.0:
                        marker_size = 5.0
                radar_style = str(
                    getattr(chart, 'radar_style', 'standard')
                ).lower()
                if chart_kind == 'radar' and radar_style == 'standard':
                    if marker_symbol == 'none':
                        marker_symbol = (
                            'diamond', 'square', 'triangle'
                        )[series_index % 3]
                    if marker_size <= 0.0:
                        marker_size = 9.0
                    if marker_border_width <= 0.0:
                        marker_border_width = 1.0
                    line_width = 3.0
                if chart_kind == 'stock' and not stock_features[0]:
                    volume_offset = 1 if stock_features[2] else 0
                    if series_index >= volume_offset:
                        marker_symbol = (
                            'diamond', 'square', 'triangle', 'x', 'circle'
                        )[(series_index - volume_offset) % 5]
                        marker_size = 9.0
                        marker_fill_color = series_color or default_color
                        marker_border_color = series_color or default_color
                        marker_border_width = max(1.0, marker_border_width)
                        line_width = 3.0
                series.append(PdfChartSeriesLayout(
                    name=str(
                        series_names[series_index]
                        if series_index < len(series_names)
                        else source_series.name or f'Series {series_index + 1}'
                    ),
                    values=tuple(values),
                    color=series_color or default_color,
                    line_width=line_width,
                    line_cap=line_cap,
                    marker_symbol=marker_symbol,
                    marker_size=marker_size,
                    marker_fill_color=marker_fill_color,
                    marker_border_color=marker_border_color,
                    marker_border_width=marker_border_width,
                    display_values=tuple(
                        str(value).strip() for value in display_values
                    ),
                    error_bars=tuple(
                        (
                            error.val_type,
                            error.bar_type,
                            bool(error.no_end_cap),
                            (
                                float(error.line_width) / 12700.0
                                if error.line_width is not None else 0.75
                            ),
                            error.line_color or '595959',
                            float(error.val),
                        )
                        for error in source_series.error_bars
                    ),
                    pattern_type=(
                        series_patterns[series_index][0]
                        if series_index < len(series_patterns) else None
                    ),
                    pattern_foreground=(
                        series_patterns[series_index][1]
                        if series_index < len(series_patterns) else None
                    ),
                    pattern_background=(
                        series_patterns[series_index][2]
                        if series_index < len(series_patterns) else None
                    ),
                    hidden=bool(getattr(source_series, 'hidden', False)),
                    x_values=tuple(x_values),
                    line_visible=(
                        series_line_visibility[series_index]
                        if series_index < len(series_line_visibility)
                        else True
                    ),
                    bubble_sizes=tuple(bubble_sizes),
                    bubble_3d=bool(getattr(source_series, '_bubble_3d', False)),
                ))
            if chart_kind == 'sunburst' and chart.n_series:
                sunburst_paths = cls._chart_sunburst_paths(
                    worksheet, chart.n_series[0]
                )
            if chart_kind == 'treemap':
                visible_series = next(
                    (
                        item for item in chart.n_series
                        if not bool(getattr(item, 'hidden', False))
                    ),
                    None,
                )
                if visible_series is not None:
                    treemap_paths = cls._chart_sunburst_paths(
                        worksheet, visible_series
                    )
            histogram_bins = (
                cls._histogram_bins(chart, series)
                if chart_kind == 'histogram'
                and histogram_variant == 'histogram'
                else ()
            )
            if chart_kind == 'pie' and not categories and series:
                categories = tuple(
                    str(value) for value in range(
                        1, len(series[0].values) + 1
                    )
                )
            (
                pie_variant,
                pie_hole_size,
                pie_secondary_size,
                pie_secondary_indices,
                pie_other_style,
                pie_connector_color,
                pie_connector_width,
            ) = cls._chart_pie_settings(
                chart,
                tuple(series[0].values) if series else (),
            )
            if pie_variant in ('pie_of_pie', 'bar_of_pie'):
                top_reserve = min(37.09, max(0.0, height))
                bottom_reserve = min(
                    33.87, max(0.0, height - top_reserve)
                )
                plot = (
                    11.0 / max(1e-12, width),
                    top_reserve / max(1e-12, height),
                    max(0.0, width - 22.0) / max(1e-12, width),
                    max(
                        0.0, height - top_reserve - bottom_reserve
                    ) / max(1e-12, height),
                )
            grouping = cls._chart_grouping(chart)
            bar_shape = cls._chart_bar_shape(chart)
            show_data_table, data_table_show_keys = (
                cls._chart_data_table(chart)
            )
            (
                show_data_labels,
                data_label_position,
                data_label_color,
            ) = cls._chart_data_labels(chart)
            if chart_kind == 'waterfall':
                if show_data_labels:
                    data_label_position = 'outEnd'
                elif not manual_layout:
                    plot = (
                        plot[0],
                        63.5 / height,
                        plot[2],
                        (height - 87.0) / height,
                    )
            (
                is_3d, rotation_x, rotation_y,
                depth_percent, perspective,
            ) = cls._chart_3d_view(chart)
            if (
                is_3d
                and chart_kind == 'bar'
                and grouping not in ('stacked', 'percentStacked')
                and perspective < 65.0
                and not manual_layout
            ):
                # Excel's 3-D bar frame uses the drawing anchor boundary
                # without the full graphic-frame inset used by 2-D charts.
                # Compact charts also carry a small left inset that disappears
                # once the chart spans a full print column band.
                if width < 420.0:
                    x -= 0.72 * scale
                    width -= 0.96 * scale
                else:
                    width -= 1.08 * scale
                plot = cls._chart_plot_ratios(chart, width, height)
            elif (
                is_3d
                and chart_kind == 'area'
                and len(series) > 3
                and not manual_layout
            ):
                width -= 1.08 * scale
                plot = cls._chart_plot_ratios(chart, width, height)
            show_series_axis = cls._chart_has_series_axis(chart)
            if is_3d and grouping == 'standard' and show_series_axis:
                series = cls._disambiguate_depth_series_names(series)
            (
                background_fill_color,
                background_gradient_stops,
                background_gradient_angle,
                background_pattern_type,
                background_pattern_foreground,
                background_pattern_background,
                chart_border_color,
                chart_border_width,
            ) = cls._chart_background_style(chart)
            back_wall_fill_image, side_wall_fill_image = (
                cls._chart_wall_fill_images(chart)
            )
            (
                show_vertical_gridlines,
                category_major_gridline_color,
                value_major_gridline_color,
                plot_border_color,
            ) = cls._chart_grid_style(chart)
            if scatter_vary_colors:
                show_vertical_gridlines = True
                category_major_gridline_color = '000000'
                value_major_gridline_color = '000000'
                plot_border_color = '000000'
            (
                show_minor_horizontal_gridlines,
                minor_gridline_color,
                value_axis_color,
                value_major_tick_mark,
                value_minor_tick_mark,
            ) = cls._chart_value_axis_style(chart)
            (
                value_axis_display_unit,
                value_axis_display_unit_label,
                value_axis_display_unit_layout_x,
                value_axis_display_unit_layout_y,
            ) = cls._chart_value_axis_display_unit(chart)
            value_axis_log_base = cls._chart_value_axis_log_base(chart)
            value_axis_position = cls._chart_value_axis_position(chart)
            (
                title_family, title_color,
                title_fill_image, title_border_color,
            ) = (
                cls._chart_title_style(chart)
            )
            (
                title_size, title_bold, title_italic, title_all_caps,
            ) = cls._chart_title_font_properties(chart)
            value_axis_text_color, value_axis_fill_image = (
                cls._chart_text_style(chart, './/c:valAx')
            )
            category_axis_text_color, category_axis_fill_image = (
                cls._chart_text_style(
                    chart, ('.//c:dateAx', './/c:catAx')
                )
            )
            legend_text_color, legend_fill_image = cls._chart_text_style(
                chart, './/c:legend'
            )
            title_gradient_stops, title_gradient_angle = (
                cls._chart_text_gradient(chart, './/c:title')
            )
            legend_gradient_stops, legend_gradient_angle = (
                cls._chart_text_gradient(chart, './/c:legend')
            )
            category_axis_title = cls._chart_axis_title(
                chart, ('.//c:dateAx', './/c:catAx')
            )
            value_axis_title = cls._chart_axis_title(
                chart, ('.//c:valAx',)
            )
            if chart_kind in ('scatter', 'bubble'):
                category_axis_title = cls._chart_scatter_axis_title(
                    chart, 'b'
                )
                value_axis_title = cls._chart_scatter_axis_title(chart, 'l')
                if value_axis_title:
                    title_gutter = 21.6 / width
                    plot = (
                        plot[0] + title_gutter,
                        plot[1],
                        max(0.0, plot[2] - title_gutter),
                        plot[3],
                    )
                if category_axis_title:
                    plot = (
                        plot[0], plot[1], plot[2],
                        max(0.0, plot[3] - 20.84 / height),
                    )
            legend_layout_x, legend_layout_y = cls._chart_legend_layout(chart)
            title_layout_x, title_layout_y = cls._chart_title_layout(chart)
            has_bitmap_text_fill = any(image is not None for image in (
                title_fill_image,
                value_axis_fill_image,
                category_axis_fill_image,
                legend_fill_image,
            ))
            if has_bitmap_text_fill and not manual_layout:
                x += 1.44 * scale
                y += 0.36 * scale
                width -= 2.52 * scale
                plot = cls._chart_plot_ratios(chart, width, height)
            if (
                chart_kind == 'bar'
                and value_axis_display_unit_label
                and not manual_layout
            ):
                plot = (
                    plot[0], plot[1], plot[2],
                    max(0.0, plot[3] - 21.21 / height),
                )
            if (
                chart_kind == 'scatter'
                and value_axis_display_unit_label
                and not manual_layout
            ):
                display_unit_gutter = 12.0 / width
                plot = (
                    plot[0] + display_unit_gutter,
                    plot[1],
                    max(0.0, plot[2] - display_unit_gutter),
                    plot[3],
                )
            bold_standard_radar_title = (
                chart_kind == 'radar'
                and str(getattr(chart, 'radar_style', 'standard')).lower()
                == 'standard'
            )
            regular_filled_radar_title = (
                chart_kind == 'radar'
                and str(getattr(chart, 'radar_style', 'standard')).lower()
                == 'filled'
            )
            if regular_filled_radar_title and not title_family:
                title_family = 'Arial'
                if title_color == '595959':
                    title_color = '000000'
            if (
                title_family
                or title_size is not None
                or title_bold is not None
                or title_italic is not None
                or bold_standard_radar_title
                or regular_filled_radar_title
            ):
                chart_title_font = title_font.copy()
                if title_family:
                    chart_title_font.name = title_family
                if title_size is not None:
                    chart_title_font.size = title_size
                if title_bold is not None:
                    chart_title_font.bold = title_bold
                if title_italic is not None:
                    chart_title_font.italic = title_italic
                if bold_standard_radar_title:
                    chart_title_font.bold = True
                elif regular_filled_radar_title:
                    chart_title_font.bold = False
                title_resolution = workbook.font_strategy.resolve(
                    chart_title_font
                )
            else:
                title_resolution = default_title_resolution
            if grouping == 'percentStacked':
                maximum = 1.0
            elif grouping == 'stacked':
                maximum = max(
                    (
                        sum(float(item.values[index]) for item in series)
                        for index in range(min(
                            (len(item.values) for item in series), default=0
                        ))
                    ),
                    default=1.0,
                )
            else:
                maximum = max(
                    (value for item in series for value in item.values),
                    default=1.0,
                )
            value_min = 0.0
            scatter_x_min = 0.0
            scatter_x_max = 1.0
            scatter_x_step = 1.0
            scatter_x_number_format = '0'
            scatter_text_outline_gradient_stops = ()
            scatter_text_outline_gradient_angle = 0.0
            if chart_kind in ('scatter', 'bubble'):
                scatter_x_number_format = cls._chart_scatter_x_number_format(
                    chart
                )
                x_numbers = tuple(
                    float(value)
                    for item in series
                    for value in item.x_values
                )
                if x_numbers:
                    x_minimum = min(x_numbers)
                    x_maximum = max(x_numbers)
                    x_span = x_maximum - x_minimum
                    is_general_axis = (
                        scatter_x_number_format.lower() == 'general'
                    )
                    is_wide_date_axis = (
                        any(
                            token in scatter_x_number_format.lower()
                            for token in ('yy', 'mmm', 'dd')
                        )
                        and x_span > 3650.0
                    )
                    if '$' in scatter_x_number_format:
                        scatter_x_step = cls._nice_chart_step(
                            max(1.0, x_maximum) / 8.0
                        )
                        scatter_x_min = 0.0
                    elif is_general_axis:
                        scatter_x_step = cls._nice_chart_step(
                            max(1.0, x_maximum) / 8.0
                        )
                        scatter_x_min = 0.0
                    elif is_wide_date_axis:
                        scatter_x_step = cls._nice_chart_step(
                            max(1.0, x_maximum) / 8.0
                        )
                        scatter_x_min = 0.0
                    else:
                        scatter_x_step = cls._nice_chart_step(
                            max(1.0, x_span) / 10.0
                        )
                        scatter_x_min = scatter_x_step * math.floor(
                            x_minimum / scatter_x_step
                        )
                        if (
                            x_minimum - scatter_x_min
                            < scatter_x_step * 0.2
                        ):
                            scatter_x_min -= scatter_x_step
                    scatter_x_max = (
                        scatter_x_step * 8.0
                        if is_wide_date_axis
                        else scatter_x_step * (
                            math.floor(x_maximum / scatter_x_step) + 1.0
                        )
                        if is_general_axis
                        else scatter_x_step * math.ceil(
                            x_maximum / scatter_x_step
                        )
                    )
                    if (
                        '$' in scatter_x_number_format
                        and scatter_x_max - x_maximum
                        < scatter_x_step * 0.1
                    ):
                        scatter_x_max += scatter_x_step
                    if chart_kind == 'bubble':
                        if '$' in scatter_x_number_format or is_general_axis:
                            scatter_x_max += scatter_x_step
                        else:
                            scatter_x_step = cls._nice_chart_step(
                                max(1.0, x_span) / 7.0
                            )
                            scatter_x_min = scatter_x_step * (
                                math.floor(x_minimum / scatter_x_step) - 1.0
                            )
                            scatter_x_max = scatter_x_step * (
                                math.ceil(x_maximum / scatter_x_step) + 1.0
                            )
                if '$' in scatter_x_number_format and not manual_layout:
                    axis_title_gutter = 21.6 if value_axis_title else 0.0
                    plot = (
                        plot[0], plot[1],
                        max(0.0, (width - 80.0 - axis_title_gutter) / width),
                        plot[3],
                    )
                (
                    scatter_text_outline_gradient_stops,
                    scatter_text_outline_gradient_angle,
                ) = cls._chart_text_outline_gradient(
                    chart
                )
            if chart_kind == 'waterfall' and series:
                running = 0.0
                levels = [0.0]
                for point_index, value in enumerate(series[0].values):
                    numeric = float(value)
                    if point_index in waterfall_subtotals:
                        running = numeric
                        levels.extend((0.0, running))
                    else:
                        levels.extend((running, running + numeric))
                        running += numeric
                minimum = min(levels, default=0.0)
                maximum = max(levels, default=1.0)
                step = cls._nice_chart_step(
                    max(1.0, maximum - minimum) / 8.0
                )
                value_min = min(0.0, step * math.floor(minimum / step))
                value_max = max(
                    step, step * math.ceil(maximum / step)
                )
            elif value_axis_log_base and grouping != 'percentStacked':
                value_min = 1.0
                maximum_exponent = max(
                    1,
                    math.ceil(math.log(
                        max(value_min, maximum), value_axis_log_base
                    )),
                )
                value_max = value_axis_log_base ** maximum_exponent
                step = 1.0
            elif grouping == 'percentStacked':
                value_axis_span = (
                    width * plot[2]
                    if chart_kind == 'bar'
                    else height * plot[3]
                )
                if is_3d and chart_kind == 'bar':
                    step = 0.2
                elif is_3d and bar_shape == 'cylinder':
                    step = 0.5
                else:
                    step = 0.1 if value_axis_span >= 130.0 else 0.2
                value_max = 1.0
            elif chart_kind == 'radar':
                radar_divisions = 10.0 if str(
                    getattr(chart, 'radar_style', 'standard')
                ).lower() in ('standard', 'filled') else 7.0
                step = cls._nice_chart_step(maximum / radar_divisions)
                if radar_divisions == 10.0:
                    value_max = step * (math.floor(maximum / step) + 1.0)
                else:
                    value_max = step * math.ceil(maximum / step)
            elif is_3d:
                if (
                    chart_kind == 'bar'
                    and perspective >= 65.0
                ):
                    step = cls._nice_chart_step(maximum / 4.0)
                    value_max = step * math.ceil(maximum / step)
                elif (
                    chart_kind == 'area'
                    and is_3d
                    and cls._chart_has_date_axis(chart)
                    and cls._chart_category_axis_reversed(chart)
                ):
                    step = cls._nice_chart_step(maximum / 2.0)
                    value_max = step * math.ceil(maximum / step)
                elif (
                    chart_kind == 'area'
                    and len(series) > 3
                    and perspective >= 65.0
                ):
                    # Excel leaves substantially more headroom for the
                    # steep, textured 3-D stacked-area camera.
                    step = cls._nice_chart_step(maximum / 2.0)
                    value_max = step * math.ceil(maximum / step)
                elif chart_kind == 'area' and len(series) > 3:
                    step = cls._nice_chart_step(maximum / 7.0)
                    value_max = step * math.ceil(maximum / step)
                elif (
                    chart_kind == 'column'
                    and grouping == 'standard'
                    and show_series_axis
                    and perspective >= 65.0
                ):
                    step = cls._nice_chart_step(maximum / 2.0)
                    value_max = step * 2.0
                elif chart_kind == 'column' and height >= 230.0:
                    step = cls._nice_chart_step(maximum / 7.0)
                    value_max = step * math.ceil(maximum / step)
                elif (
                    chart_kind == 'bar'
                    and (
                        width * plot[2] >= 350.0
                        or (
                            cls._chart_has_date_axis(chart)
                            and len(cls._date_axis_layout(categories)[1])
                            > len(categories)
                        )
                    )
                ):
                    step = cls._nice_chart_step(maximum / 7.0)
                    value_max = step * math.ceil(maximum / step)
                else:
                    step = cls._nice_chart_step(maximum / 4.0)
                    value_max = step * 4.0
            elif chart_kind == 'stock' and stock_features[2] and series:
                maximum = max(series[0].values, default=1.0)
                step = cls._nice_chart_step(maximum / 10.0)
                value_max = step * math.ceil(maximum / step)
            elif chart_kind == 'stock' and stock_features[1]:
                step = cls._nice_chart_step(maximum / 14.0)
                value_max = step * math.ceil(maximum / step)
                value_min = max(0.0, value_max - step * 7.0)
            elif chart_kind == 'stock':
                step = cls._nice_chart_step(maximum / 8.0)
                value_max = (
                    step * 8.0
                    if all(len(str(label)) <= 6 for label in categories)
                    else step * math.ceil(maximum / step)
                )
            elif chart_kind == 'scatter' and (
                category_axis_title or value_axis_title
            ):
                step = cls._nice_chart_step(maximum / 4.0)
                value_max = step * math.ceil(maximum / step)
            elif chart_kind == 'scatter' and (
                value_axis_display_unit != 1.0
                or '%' in cls._chart_scatter_y_number_format(chart)
            ):
                step = cls._nice_chart_step(maximum / 8.0)
                value_max = step * math.ceil(maximum / step)
            elif chart_kind == 'scatter':
                step = cls._nice_chart_step(maximum / 8.0)
                value_max = step * math.ceil(maximum / step)
                if value_max - maximum < step * 0.1:
                    value_max += step
            else:
                step = cls._nice_chart_step(maximum / 8.0)
                value_max = step * 8.0
            if chart_kind == 'bubble':
                value_max += step
            stock_price_min = value_min
            stock_price_max = value_max
            stock_price_step = step
            if chart_kind == 'stock' and stock_features[2] and len(series) > 1:
                price_values = tuple(
                    float(value)
                    for item in series[1:]
                    for value in item.values
                )
                price_maximum = max(price_values, default=1.0)
                if stock_features[1]:
                    stock_price_step = cls._nice_chart_step(
                        price_maximum / 8.0
                    )
                    stock_price_min = 0.0
                    stock_price_max = stock_price_step * math.ceil(
                        price_maximum / stock_price_step
                    )
                else:
                    price_minimum = min(price_values, default=0.0)
                    stock_price_step = cls._nice_chart_step(
                        max(1.0, price_maximum - price_minimum) / 10.0
                    )
                    stock_price_max = stock_price_step * math.ceil(
                        price_maximum / stock_price_step
                    ) + stock_price_step
                    stock_price_min = max(
                        0.0, stock_price_max - stock_price_step * 10.0
                    )
            if not manual_layout and grouping != 'percentStacked':
                widest_label = f'${int(round(value_max)):,}'
                label_width = workbook.text_measurer.measure_text(
                    widest_label, axis_font
                ).width_points
                baseline_width = workbook.text_measurer.measure_text(
                    '$800,000', axis_font
                ).width_points
                extra_left = max(0.0, label_width - baseline_width)
                if chart_kind == 'area' and extra_left > 0.0:
                    extra_left += 0.15
                if chart_kind == 'bar' and value_axis_log_base:
                    plot = (
                        plot[0],
                        plot[1],
                        plot[2] - extra_left / (2.0 * width),
                        plot[3],
                    )
                else:
                    plot = (
                        plot[0] + extra_left / width,
                        plot[1],
                        plot[2] - extra_left / width,
                        plot[3],
                    )
            if show_data_table and not manual_layout:
                row_label_width = max(
                    (
                        workbook.text_measurer.measure_text(
                            item.name, axis_font
                        ).width_points
                        for item in series
                    ),
                    default=0.0,
                ) + (11.05 if data_table_show_keys else 6.0)
                table_left = 11.0
                plot = (
                    (table_left + row_label_width) / width,
                    plot[1],
                    (
                        width - table_left - row_label_width - 11.0
                    ) / width,
                    max(0.0, (height - 149.15) / height),
                )
                if grouping == 'percentStacked' and height * plot[3] < 70.0:
                    step = 0.5
            has_date_axis = cls._chart_has_date_axis(chart)
            reversed_date_axis = (
                has_date_axis
                and cls._chart_category_axis_reversed(chart)
            )
            if (
                chart_kind == 'area'
                and not manual_layout
                and not reversed_date_axis
            ):
                category_positions = tuple(
                    index / max(1, len(categories) - 1)
                    for index in range(len(categories))
                )
                axis_categories = tuple(zip(categories, category_positions))
                category_rotation = 0.0
            elif manual_layout:
                category_positions = tuple(
                    (index + 0.5) / max(1, len(categories))
                    for index in range(len(categories))
                )
                axis_categories = tuple(zip(categories, category_positions))
                category_rotation = 0.0
            elif not has_date_axis:
                category_positions = tuple(
                    (index + 0.5) / max(1, len(categories))
                    for index in range(len(categories))
                )
                axis_categories = tuple(zip(categories, category_positions))
                category_rotation = 0.0
            else:
                category_positions, axis_categories = cls._date_axis_layout(
                    categories,
                    reverse=reversed_date_axis,
                )
                if chart_kind == 'bar':
                    category_rotation = 0.0
                elif chart_kind == 'column':
                    if (
                        is_3d
                        and grouping == 'standard'
                        and show_series_axis
                        and perspective >= 65.0
                    ):
                        category_rotation = -90.0
                    else:
                        category_rotation = (
                            -45.0
                            if is_3d and rotation_y >= 35.0 else 0.0
                        )
                else:
                    if chart_kind == 'stock':
                        category_rotation = (
                            0.0
                            if stock_features[3] or all(
                                len(str(label)) <= 6 for label in categories
                            )
                            else -45.0
                        )
                    else:
                        category_rotation = -90.0 if is_3d else -45.0
            if (
                grouping == 'percentStacked'
                and category_axis_fill_image is None
                and not show_data_table
                and chart_kind not in ('bar', 'column', 'area')
            ):
                axis_categories = ()
            if chart_kind == 'stock' and not stock_features[3]:
                category_rotation = (
                    0.0
                    if all(len(str(label)) <= 6 for label in categories)
                    else -45.0
                )
            layouts.append(PdfChartLayout(
                name=f'Chart {index + 1}',
                chart_kind=chart_kind,
                bar_direction=str(getattr(chart, 'bar_direction', 'col')),
                bar_shape=bar_shape,
                gap_width=float(getattr(chart, 'gap_width', 150.0)),
                overlap=float(getattr(chart, 'overlap', 0.0)),
                show_data_labels=show_data_labels,
                data_label_position=data_label_position,
                data_label_color=data_label_color,
                x=x,
                y=y,
                width=width,
                height=height,
                background_fill_color=background_fill_color,
                background_gradient_stops=background_gradient_stops,
                background_gradient_angle=background_gradient_angle,
                background_pattern_type=background_pattern_type,
                background_pattern_foreground=(
                    background_pattern_foreground
                ),
                background_pattern_background=(
                    background_pattern_background
                ),
                chart_border_color=chart_border_color,
                chart_border_width=chart_border_width,
                back_wall_fill_image=back_wall_fill_image,
                side_wall_fill_image=side_wall_fill_image,
                plot_x=x + width * plot[0],
                plot_y=y + height * plot[1],
                plot_width=width * plot[2],
                plot_height=height * plot[3],
                title=(
                    cls._chart_title(
                        chart, series[0].name if series else None
                    ).upper()
                    if title_all_caps else
                    cls._chart_title(
                        chart, series[0].name if series else None
                    )
                ),
                title_layout_x=title_layout_x,
                title_layout_y=title_layout_y,
                title_color=title_color,
                title_fill_image=title_fill_image,
                title_border_color=title_border_color,
                categories=tuple(categories),
                category_positions=category_positions,
                axis_categories=axis_categories,
                category_rotation=category_rotation,
                show_vertical_gridlines=show_vertical_gridlines,
                category_major_gridline_color=(
                    category_major_gridline_color
                ),
                value_major_gridline_color=value_major_gridline_color,
                plot_border_color=plot_border_color,
                show_minor_horizontal_gridlines=(
                    show_minor_horizontal_gridlines
                ),
                minor_gridline_color=minor_gridline_color,
                value_axis_color=value_axis_color,
                value_axis_position=value_axis_position,
                value_major_tick_mark=value_major_tick_mark,
                value_minor_tick_mark=value_minor_tick_mark,
                value_axis_text_color=value_axis_text_color,
                value_axis_fill_image=value_axis_fill_image,
                category_axis_text_color=category_axis_text_color,
                category_axis_fill_image=category_axis_fill_image,
                category_axis_title=category_axis_title,
                value_axis_title=value_axis_title,
                legend_text_color=legend_text_color,
                legend_fill_image=legend_fill_image,
                legend_layout_x=legend_layout_x,
                legend_layout_y=legend_layout_y,
                is_3d=is_3d,
                rotation_x=rotation_x,
                rotation_y=rotation_y,
                depth_percent=depth_percent,
                perspective=perspective,
                series=tuple(series),
                grouping=grouping,
                value_min=value_min,
                value_max=value_max,
                value_step=step,
                value_axis_log_base=value_axis_log_base,
                axis_number_format=(
                    '0%' if grouping == 'percentStacked'
                    else cls._chart_scatter_y_number_format(chart)
                    if chart_kind in ('scatter', 'bubble')
                    else '0' if chart_kind == 'waterfall'
                    and value_max >= 1000000.0
                    else '#,##0' if chart_kind == 'waterfall'
                    else '0' if chart_kind == 'stock'
                    and cls._chart_stock_has_general_axes(chart)
                    else '0' if chart_kind == 'surface'
                    and cls._chart_surface_has_general_axis(chart)
                    else 'mmm-yy' if chart_kind == 'box_whisker'
                    else '$#,##0'
                ),
                value_axis_display_unit=value_axis_display_unit,
                value_axis_display_unit_label=(
                    value_axis_display_unit_label
                ),
                value_axis_display_unit_layout_x=(
                    value_axis_display_unit_layout_x
                ),
                value_axis_display_unit_layout_y=(
                    value_axis_display_unit_layout_y
                ),
                show_legend=bool(chart.show_legend),
                show_data_table=show_data_table,
                data_table_show_keys=data_table_show_keys,
                show_series_axis=show_series_axis,
                title_font_resolution=title_resolution,
                axis_font_resolution=axis_resolution,
                pie_point_styles=cls._chart_pie_point_styles(
                    chart,
                    len(series[0].values) if series else 0,
                ),
                pie_series_point_styles=tuple(
                    cls._chart_pie_point_styles(
                        chart, len(item.values), series_index,
                    )
                    for series_index, item in enumerate(series)
                ),
                pie_first_slice_angle=float(
                    getattr(chart, 'first_slice_angle', 0.0)
                ),
                pie_variant=pie_variant,
                pie_hole_size=pie_hole_size,
                pie_secondary_size=pie_secondary_size,
                pie_secondary_indices=pie_secondary_indices,
                pie_other_style=pie_other_style,
                pie_connector_color=pie_connector_color,
                pie_connector_width=pie_connector_width,
                title_gradient_stops=title_gradient_stops,
                title_gradient_angle=title_gradient_angle,
                legend_gradient_stops=legend_gradient_stops,
                legend_gradient_angle=legend_gradient_angle,
                chart_glow_color=chart_glow_color,
                chart_glow_opacity=chart_glow_opacity,
                chart_glow_radius=chart_glow_radius,
                chart_soft_edge_radius=chart_soft_edge_radius * scale,
                chart_bevel_top_width=chart_bevel_top_width,
                radar_style=str(
                    getattr(chart, 'radar_style', 'standard')
                ).lower(),
                legend_position=str(
                    getattr(chart, 'legend_position', 'b') or 'b'
                ).lower(),
                box_show_mean_line=bool(
                    getattr(chart, 'box_show_mean_line', False)
                ),
                box_show_mean_marker=bool(
                    getattr(chart, 'box_show_mean_marker', True)
                ),
                box_show_inner_points=bool(
                    getattr(chart, 'box_show_inner_points', False)
                ),
                box_show_outlier_points=bool(
                    getattr(chart, 'box_show_outlier_points', True)
                ),
                box_gap_width=float(getattr(chart, 'box_gap_width', 1.0)),
                box_quartile_method=str(
                    getattr(chart, 'quartile_method', 'exclusive')
                ).lower(),
                histogram_variant=histogram_variant,
                histogram_bins=histogram_bins,
                map_value_colors=cls._chart_map_value_colors(chart),
                stock_style=str(
                    getattr(chart, 'stock_style', 'high_low_close')
                ),
                stock_has_high_low_lines=stock_features[0],
                stock_has_up_down_bars=stock_features[1],
                stock_has_volume=stock_features[2],
                stock_axes_hidden=stock_features[3],
                rounded_corners=(
                    stock_features[4]
                    or surface_features[1]
                    or cls._chart_rounded_corners(chart)
                ),
                stock_price_min=stock_price_min,
                stock_price_max=stock_price_max,
                stock_price_step=stock_price_step,
                sunburst_paths=sunburst_paths,
                treemap_paths=treemap_paths,
                waterfall_subtotals=waterfall_subtotals,
                waterfall_gap_width=waterfall_gap_width,
                scatter_style=str(
                    getattr(chart, 'scatter_style', 'lineMarker')
                ),
                scatter_x_min=scatter_x_min,
                scatter_x_max=scatter_x_max,
                scatter_x_step=scatter_x_step,
                scatter_x_number_format=scatter_x_number_format,
                scatter_text_outline_gradient_stops=(
                    scatter_text_outline_gradient_stops
                ),
                scatter_text_outline_gradient_angle=(
                    scatter_text_outline_gradient_angle
                ),
                scatter_vary_colors=scatter_vary_colors,
                bubble_scale=float(getattr(chart, '_bubble_scale', 100.0)),
                surface_wireframe=bool(getattr(chart, 'wireframe', False)),
                surface_band_colors=cls._chart_surface_band_colors(chart),
                surface_axes_hidden=surface_features[0],
            ))
        if layouts and all(item.chart_kind == 'stock' for item in layouts):
            column_origin = sum(
                column_width(column)
                for column in range(1, column_chunk[0])
            )
            stock_page_x = []
            for source_chart in worksheet.charts:
                if source_chart.type != ChartType.STOCK or not (
                    column_chunk[0]
                    <= source_chart._lower_right_column + 1
                    <= column_chunk[1]
                ):
                    continue
                chart_left = sum(
                    column_width(column)
                    for column in range(
                        1, source_chart._upper_left_column + 1
                    )
                ) + source_chart._upper_left_column_offset / 12700.0
                stock_page_x.append(
                    left + (chart_left - column_origin) * scale + 1.2 * scale
                )
            minimum_x = min(
                stock_page_x, default=min(item.x for item in layouts)
            )
            shift_x = max(0.0, left + 49.0 - minimum_x)
            minimum_y = min(item.y for item in layouts)
            shift_y = (
                top + 63.0 - minimum_y if minimum_y < top else 0.0
            )
            if shift_x or shift_y:
                layouts = [
                    replace(
                        item,
                        x=item.x + shift_x,
                        y=item.y + shift_y,
                        plot_x=item.plot_x + shift_x,
                        plot_y=item.plot_y + shift_y,
                    )
                    for item in layouts
                ]
        return tuple(layouts)

    @staticmethod
    def _chart_radar_uses_default_markers(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return True

    @staticmethod
    def _chart_waterfall_subtotals(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
        except ET.ParseError:
            return ()

        indices = []
        for subtotals in root.iter():
            if subtotals.tag.rsplit('}', 1)[-1] != 'subtotals':
                continue
            for child in subtotals:
                if child.tag.rsplit('}', 1)[-1] != 'idx':
                    continue
                try:
                    indices.append(int(child.get('val')))
                except (TypeError, ValueError):
                    continue
        return tuple(dict.fromkeys(indices))

    @staticmethod
    def _chart_waterfall_gap_width(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 0.35
        try:
            root = ET.fromstring(source)
        except ET.ParseError:
            return 0.35
        for node in root.iter():
            if node.tag.rsplit('}', 1)[-1] != 'catScaling':
                continue
            try:
                return max(0.0, float(node.get('gapWidth', 0.35)))
            except (TypeError, ValueError):
                return 0.35
        return 0.35

    @staticmethod
    def _chart_sunburst_paths(worksheet, source_series):
        category_formula = getattr(source_series, 'category_data', None)
        value_formula = getattr(source_series, 'values', None)
        if not category_formula or not value_formula:
            return ()

        category_reference = str(category_formula).rsplit('!', 1)[-1]
        category_reference = category_reference.replace('$', '')
        category_start, _, category_end = category_reference.partition(':')
        start_row, start_column = worksheet.cells.coordinate_from_string(
            category_start
        )
        end_row, end_column = worksheet.cells.coordinate_from_string(
            category_end or category_start
        )

        value_reference = str(value_formula).rsplit('!', 1)[-1]
        value_reference = value_reference.replace('$', '')
        value_start, _, value_end = value_reference.partition(':')
        value_start_row, value_column = (
            worksheet.cells.coordinate_from_string(value_start)
        )
        value_end_row, _ = worksheet.cells.coordinate_from_string(
            value_end or value_start
        )
        row_count = min(
            end_row - start_row + 1,
            value_end_row - value_start_row + 1,
        )
        cells = worksheet.cells.get_all_cells()
        current_path = [''] * (end_column - start_column + 1)
        paths = []
        for offset in range(row_count):
            row = start_row + offset
            explicit_level = None
            for level, column in enumerate(
                range(start_column, end_column + 1)
            ):
                ref = worksheet.cells.coordinate_to_string(row, column)
                cell = cells.get(ref)
                label = (
                    str(cell.get_display_text()).strip()
                    if cell is not None else ''
                )
                if label:
                    current_path[level] = label
                    current_path[level + 1:] = [''] * (
                        len(current_path) - level - 1
                    )
                    explicit_level = level
            value_ref = worksheet.cells.coordinate_to_string(
                value_start_row + offset, value_column
            )
            value_cell = cells.get(value_ref)
            if value_cell is None:
                continue
            try:
                value = float(value_cell.value or 0.0)
            except (TypeError, ValueError):
                continue
            path = tuple(label for label in current_path if label)
            if path and (explicit_level is not None or value):
                paths.append((path, value))
        return tuple(paths)

    @staticmethod
    def _chart_stock_features(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, False, False, False, False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            stock = root.find('.//c:stockChart', ns)
            axes = root.findall('.//c:dateAx', ns) + root.findall(
                './/c:catAx', ns
            ) + root.findall('.//c:valAx', ns)
            axes_hidden = bool(axes) and all(
                (deleted := axis.find('c:delete', ns)) is not None
                and deleted.get('val', '0') in ('1', 'true', 'True')
                for axis in axes
            )
            rounded = root.find('.//c:roundedCorners', ns)
            return (
                stock is not None
                and stock.find('c:hiLowLines', ns) is not None,
                stock is not None
                and stock.find('c:upDownBars', ns) is not None,
                root.find('.//c:barChart', ns) is not None,
                axes_hidden,
                rounded is not None
                and rounded.get('val', '1') in ('1', 'true', 'True'),
            )
        except ET.ParseError:
            return False, False, False, False, False

    @staticmethod
    def _chart_rounded_corners(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            rounded = root.find('.//c:roundedCorners', ns)
            return (
                rounded is not None
                and rounded.get('val', '1') in ('1', 'true', 'True')
            )
        except ET.ParseError:
            return False

    @staticmethod
    def _chart_surface_features(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            axes = (
                root.findall('.//c:catAx', ns)
                + root.findall('.//c:serAx', ns)
                + root.findall('.//c:valAx', ns)
            )
            axes_hidden = bool(axes) and all(
                (deleted := axis.find('c:delete', ns)) is not None
                and deleted.get('val', '0') in ('1', 'true', 'True')
                for axis in axes
            )
            rounded = root.find('.//c:roundedCorners', ns)
            return (
                axes_hidden,
                rounded is not None
                and rounded.get('val', '1') in ('1', 'true', 'True'),
            )
        except ET.ParseError:
            return False, False

    @staticmethod
    def _chart_stock_has_general_axes(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            formats = [
                item.get('formatCode', '')
                for item in root.findall(
                    './/c:dateAx/c:numFmt', ns
                ) + root.findall('.//c:catAx/c:numFmt', ns)
            ]
            return any(value.lower() == 'general' for value in formats)
        except ET.ParseError:
            return False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            markers = root.findall('.//c:radarChart/c:ser/c:marker', ns)
            if not markers:
                return True
            return not any(
                marker.find('c:symbol', ns) is not None
                and marker.find('c:symbol', ns).get('val') == 'none'
                for marker in markers
            )
        except ET.ParseError:
            return True

    @staticmethod
    def _chart_histogram_variant(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if source and b'paretoLine' in source:
            return 'pareto'
        return 'histogram'

    @staticmethod
    def _histogram_bins(chart, series):
        visible = next((item for item in series if not item.hidden), None)
        if visible is None:
            return ()
        values = sorted(float(value) for value in visible.values)
        if not values:
            return ()
        minimum = values[0]
        maximum = values[-1]
        requested_count = getattr(chart, 'histogram_bin_count', None)
        requested_size = getattr(chart, 'histogram_bin_size', None)
        if requested_size and requested_size > 0:
            bin_size = float(requested_size)
        elif requested_count and requested_count > 0:
            bin_size = max(1e-12, (maximum - minimum) / requested_count)
        elif len(values) > 1:
            mean = sum(values) / len(values)
            variance = sum(
                (value - mean) ** 2 for value in values
            ) / (len(values) - 1)
            scott_width = 3.5 * math.sqrt(variance) / len(values) ** (1.0 / 3.0)
            magnitude = 10.0 ** max(0, math.floor(math.log10(scott_width)) - 1)
            bin_size = round(scott_width / magnitude) * magnitude
        else:
            bin_size = 1.0
        bin_size = max(1e-12, bin_size)
        count = max(1, int(math.ceil((maximum - minimum) / bin_size)))
        closed = str(getattr(chart, 'histogram_interval_closed', 'r')).lower()
        bins = []
        for index in range(count):
            lower = minimum + bin_size * index
            upper = lower + bin_size
            if index == 0:
                frequency = sum(lower <= value <= upper for value in values)
                label = f'[{lower:g}, {upper:g}]' if closed == 'r' else f'[{lower:g}, {upper:g})'
            elif closed == 'r':
                frequency = sum(lower < value <= upper for value in values)
                label = f'({lower:g}, {upper:g}]'
            else:
                frequency = sum(lower <= value < upper for value in values)
                label = f'[{lower:g}, {upper:g})'
            bins.append((label, float(frequency), lower, upper))
        return tuple(bins)

    @staticmethod
    def _chart_data_table(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            table = root.find('.//c:plotArea/c:dTable', ns)
            if table is None:
                return False, False
            keys = table.find('c:showKeys', ns)
            return True, (
                keys is not None
                and keys.get('val', '1') not in ('0', 'false', 'False')
            )
        except ET.ParseError:
            return False, False

    @staticmethod
    def _chart_range_values(worksheet, formula, display):
        if not formula:
            return ()
        if isinstance(formula, (tuple, list)):
            return tuple(formula)
        reference = str(formula).rsplit('!', 1)[-1].replace('$', '')
        start, _, end = reference.partition(':')
        start_row, start_col = worksheet.cells.coordinate_from_string(start)
        end_row, end_col = worksheet.cells.coordinate_from_string(end or start)
        values = []
        for row in range(start_row, end_row + 1):
            for column in range(start_col, end_col + 1):
                ref = worksheet.cells.coordinate_to_string(row, column)
                cell = worksheet.cells.get_all_cells().get(ref)
                if cell is None:
                    continue
                values.append(
                    cell.get_display_text() if display else float(cell.value or 0)
                )
        return tuple(values)

    @staticmethod
    def _chart_series_names(chart, worksheet):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            names = []
            for series in root.findall('.//c:plotArea/*/c:ser', ns):
                formula = series.find('c:tx/c:strRef/c:f', ns)
                resolved = ()
                if formula is not None and formula.text:
                    try:
                        resolved = PdfLayoutEngine._chart_range_values(
                            worksheet, formula.text.strip(), display=True
                        )
                    except (TypeError, ValueError):
                        resolved = ()
                direct = series.find('c:tx/c:v', ns)
                cached = series.find(
                    'c:tx/c:strRef/c:strCache/c:pt/c:v', ns
                )
                name = (
                    resolved[0] if resolved
                    else direct.text if direct is not None and direct.text
                    else cached.text if cached is not None and cached.text
                    else None
                )
                names.append(str(name) if name is not None else '')
            return tuple(names)
        except ET.ParseError:
            return ()

    @staticmethod
    def _disambiguate_depth_series_names(series):
        if not series or len({item.name for item in series}) != 1:
            return series
        name = series[0].name
        if not name or not name[-1:].isalpha() or len(series) > 26:
            return series
        first = ord(name[-1:].upper())
        if first + len(series) - 1 > ord('Z'):
            return series
        stem = name[:-1]
        return [
            replace(item, name=stem + chr(first + index))
            for index, item in enumerate(series)
        ]

    @staticmethod
    def _chart_has_manual_layout(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            layout = root.find(
                './/c:plotArea/c:layout/c:manualLayout', ns
            )
            return (
                layout is not None
                and all(
                    layout.find(f'c:{name}', ns) is not None
                    for name in ('x', 'y', 'w', 'h')
                )
            )
        except ET.ParseError:
            return False

    @staticmethod
    def _chart_has_date_axis(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return True
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            return root.find('.//c:dateAx', ns) is not None
        except ET.ParseError:
            return True

    @staticmethod
    def _chart_grouping(chart):
        grouping = getattr(chart, 'grouping', None)
        if grouping:
            return str(grouping)
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 'standard'
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            grouping = root.find('.//c:lineChart/c:grouping', ns)
            if grouping is not None:
                return grouping.get('val', 'standard')
        except ET.ParseError:
            pass
        return 'standard'

    @staticmethod
    def _chart_bar_shape(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 'box'
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            shape = root.find('.//c:bar3DChart/c:shape', ns)
            if shape is None:
                shape = root.find('.//c:barChart/c:shape', ns)
            return shape.get('val', 'box') if shape is not None else 'box'
        except ET.ParseError:
            return 'box'

    @staticmethod
    def _chart_3d_view(chart):
        view = getattr(chart, 'view_3d', None)
        if bool(getattr(chart, 'is_3d', False)) and view is not None:
            return (
                True,
                float(view.rotation_x),
                float(view.rotation_y),
                float(view.depth_percent),
                float(view.perspective),
            )
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, 0.0, 0.0, 100.0, 30.0
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            is_3d = root.find('.//c:line3DChart', ns) is not None
            view = root.find('.//c:view3D', ns)
            rotation_x = view.find('c:rotX', ns) if view is not None else None
            rotation_y = view.find('c:rotY', ns) if view is not None else None
            depth = view.find('c:depthPercent', ns) if view is not None else None
            perspective = (
                view.find('c:perspective', ns) if view is not None else None
            )
            return (
                is_3d,
                float(rotation_x.get('val', 0)) if rotation_x is not None else 0.0,
                float(rotation_y.get('val', 0)) if rotation_y is not None else 0.0,
                float(depth.get('val', 100)) if depth is not None else 100.0,
                (
                    float(perspective.get('val', 30))
                    if perspective is not None else 30.0
                ),
            )
        except (ET.ParseError, TypeError, ValueError):
            return False, 0.0, 0.0, 100.0, 30.0

    @staticmethod
    def _chart_title_style(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, '595959', None, None
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            title = root.find('.//c:title', ns)
            if title is None:
                return None, '595959', None, None
            default_properties = title.find(
                './/c:txPr/a:p/a:pPr/a:defRPr', ns
            )
            run_properties = title.find(
                './/c:tx/c:rich/a:p/a:r/a:rPr', ns
            )
            properties = default_properties
            if run_properties is not None and (
                run_properties.find('a:solidFill', ns) is not None
                or run_properties.find('a:blipFill', ns) is not None
            ):
                properties = run_properties
            latin = (
                properties.find('a:latin', ns)
                if properties is not None else None
            )
            if latin is None and default_properties is not None:
                latin = default_properties.find('a:latin', ns)
            text_fill = (
                properties.find('a:solidFill', ns)
                if properties is not None else None
            )
            line_fill = title.find('c:spPr/a:ln/a:solidFill', ns)
            return (
                latin.get('typeface') if latin is not None else None,
                PdfLayoutEngine._chart_color(text_fill, '595959'),
                (
                    None if text_fill is not None
                    else PdfLayoutEngine._chart_blip_image(chart, properties)
                ),
                PdfLayoutEngine._chart_color(line_fill, None),
            )
        except ET.ParseError:
            return None, '595959', None, None

    @staticmethod
    def _chart_title_font_properties(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, None, None, False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            title = root.find('.//c:title', ns)
            if title is None:
                return None, None, None, False
            properties = title.find(
                './/c:txPr/a:p/a:pPr/a:defRPr', ns
            )
            if properties is None:
                properties = title.find(
                    './/c:tx/c:rich/a:p/a:r/a:rPr', ns
                )
            if properties is None:
                return None, None, None, False
            size = properties.get('sz')
            return (
                float(size) / 100.0 if size is not None else None,
                (
                    properties.get('b') == '1'
                    if properties.get('b') is not None else None
                ),
                (
                    properties.get('i') == '1'
                    if properties.get('i') is not None else None
                ),
                properties.get('cap') == 'all',
            )
        except (ET.ParseError, TypeError, ValueError):
            return None, None, None, False

    @staticmethod
    def _chart_text_style(chart, scope):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return '595959', None
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            scopes = (scope,) if isinstance(scope, str) else scope
            node = next(
                (candidate for candidate in (
                    root.find(candidate_scope, ns)
                    for candidate_scope in scopes
                ) if candidate is not None),
                None,
            )
            properties = (
                node.find('.//c:txPr/a:p/a:pPr/a:defRPr', ns)
                if node is not None else None
            )
            solid = (
                properties.find('a:solidFill', ns)
                if properties is not None else None
            )
            return (
                PdfLayoutEngine._chart_color(solid, '595959'),
                PdfLayoutEngine._chart_blip_image(chart, properties),
            )
        except ET.ParseError:
            return '595959', None

    @staticmethod
    def _chart_text_gradient(chart, scope):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return (), 0.0
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            node = root.find(scope, ns)
            properties = (
                node.find('.//c:txPr/a:p/a:pPr/a:defRPr', ns)
                if node is not None else None
            )
            gradient = (
                properties.find('a:gradFill', ns)
                if properties is not None else None
            )
            if gradient is None:
                return (), 0.0
            stops = []
            for stop in gradient.findall('a:gsLst/a:gs', ns):
                color = PdfLayoutEngine._chart_color(stop, None)
                if color:
                    stops.append((
                        float(stop.get('pos', 0.0)) / 100000.0,
                        'FF' + color,
                    ))
            linear = gradient.find('a:lin', ns)
            angle = (
                float(linear.get('ang', 0.0)) / 60000.0
                if linear is not None else 0.0
            )
            return tuple(stops), angle
        except (ET.ParseError, TypeError, ValueError):
            return (), 0.0

    @staticmethod
    def _chart_text_outline_gradient(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return (), 0.0
        try:
            root = ET.fromstring(source)
            ns = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }
            gradient = root.find(
                './/a:defRPr/a:ln/a:gradFill', ns
            )
            if gradient is None:
                return (), 0.0
            stops = []
            for stop in gradient.findall('a:gsLst/a:gs', ns):
                color = PdfLayoutEngine._chart_color(stop, None)
                if color:
                    stops.append((
                        float(stop.get('pos', 0.0)) / 100000.0,
                        'FF' + color,
                    ))
            linear = gradient.find('a:lin', ns)
            angle = (
                float(linear.get('ang', 0.0)) / 60000.0
                if linear is not None else 0.0
            )
            return tuple(stops), angle
        except (ET.ParseError, TypeError, ValueError):
            return (), 0.0

    @staticmethod
    def _chart_blip_image(chart, properties):
        if properties is None:
            return None
        blip = properties.find(
            './/{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
        )
        if blip is None:
            return None
        relationship_id = blip.get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
        )
        relationships = getattr(chart, '_source_chart_rels_xml', None)
        if not relationship_id or not relationships:
            return None
        try:
            root = ET.fromstring(relationships)
            target = None
            for relationship in root:
                if relationship.get('Id') == relationship_id:
                    target = relationship.get('Target')
                    break
            if not target:
                return None
            file_name = target.replace('\\', '/').rsplit('/', 1)[-1]
            for part_path, part_bytes, _ in getattr(
                chart, '_source_chart_extra_parts', ()
            ):
                if part_path.replace('\\', '/').endswith('/' + file_name):
                    return part_bytes
        except ET.ParseError:
            return None
        return None

    @staticmethod
    def _chart_wall_fill_images(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, None
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            back = root.find('.//c:backWall/c:spPr', ns)
            side = root.find('.//c:sideWall/c:spPr', ns)
            return (
                PdfLayoutEngine._chart_blip_image(chart, back),
                PdfLayoutEngine._chart_blip_image(chart, side),
            )
        except ET.ParseError:
            return None, None

    @staticmethod
    def _chart_has_series_axis(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            axis = root.find('.//c:serAx', ns)
            if axis is None:
                return False
            deleted = axis.find('c:delete', ns)
            return (
                deleted is None
                or deleted.get('val', '0') not in ('1', 'true', 'True')
            )
        except ET.ParseError:
            return False

    @staticmethod
    def _chart_category_axis_reversed(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            axis = root.find('.//c:dateAx', ns)
            if axis is None:
                axis = root.find('.//c:catAx', ns)
            orientation = (
                axis.find('c:scaling/c:orientation', ns)
                if axis is not None else None
            )
            return (
                orientation is not None
                and orientation.get('val') == 'maxMin'
            )
        except ET.ParseError:
            return False

    @staticmethod
    def _chart_background_style(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, (), 0.0, None, '156082', 'FFFFFF', 'D9D9D9', 0.75
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            properties = root.find('c:spPr', ns)
            if properties is None:
                return None, (), 0.0, None, '156082', 'FFFFFF', 'D9D9D9', 0.75
            line = properties.find('a:ln', ns)
            line_fill = (
                line.find('a:solidFill', ns) if line is not None else None
            )
            line_color = PdfLayoutEngine._chart_color(
                line_fill, 'D9D9D9'
            )
            if line is not None and line.find('a:noFill', ns) is not None:
                line_color = None
            line_width = (
                float(line.get('w', 9525)) / 12700.0
                if line is not None else 0.75
            )
            solid = properties.find('a:solidFill', ns)
            if solid is not None:
                return (
                    PdfLayoutEngine._chart_color(solid, None), (), 0.0,
                    None, '156082', 'FFFFFF',
                    line_color, line_width,
                )
            pattern = properties.find('a:pattFill', ns)
            if pattern is not None:
                foreground = pattern.find('a:fgClr', ns)
                background = pattern.find('a:bgClr', ns)
                return (
                    None, (), 0.0,
                    pattern.get('prst'),
                    PdfLayoutEngine._chart_color(foreground, '156082'),
                    PdfLayoutEngine._chart_color(background, 'FFFFFF'),
                    line_color, line_width,
                )
            gradient = properties.find('a:gradFill', ns)
            if gradient is None:
                return (
                    None, (), 0.0, None, '156082', 'FFFFFF',
                    line_color, line_width,
                )
            stops = []
            for stop in gradient.findall('a:gsLst/a:gs', ns):
                stops.append((
                    max(0.0, min(
                        1.0, float(stop.get('pos', 0)) / 100000.0
                    )),
                    PdfLayoutEngine._chart_color(stop, 'FFFFFF'),
                ))
            stops.sort(key=lambda item: item[0])
            linear = gradient.find('a:lin', ns)
            angle = (
                float(linear.get('ang', 0)) / 60000.0
                if linear is not None else 0.0
            )
            return (
                None, tuple(stops), angle, None, '156082', 'FFFFFF',
                line_color, line_width,
            )
        except (ET.ParseError, TypeError, ValueError):
            return None, (), 0.0, None, '156082', 'FFFFFF', 'D9D9D9', 0.75

    @staticmethod
    def _chart_glow_style(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, 0.0, 0.0
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            glow = root.find('c:spPr/a:effectLst/a:glow', ns)
            if glow is None:
                return None, 0.0, 0.0
            color_element = next(iter(glow), None)
            alpha = (
                color_element.find('.//a:alpha', ns)
                if color_element is not None else None
            )
            return (
                PdfLayoutEngine._chart_color(color_element, '156082'),
                max(0.0, min(
                    1.0,
                    float(alpha.get('val', 100000)) / 100000.0
                    if alpha is not None else 1.0,
                )),
                max(0.0, float(glow.get('rad', 0.0)) / 12700.0),
            )
        except (ET.ParseError, TypeError, ValueError):
            return None, 0.0, 0.0

    @staticmethod
    def _chart_bevel_top_width(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 0.0
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            bevel = root.find('c:spPr/a:sp3d/a:bevelT', ns)
            if bevel is None:
                return 0.0
            return max(0.0, float(bevel.get('w', 0.0)) / 12700.0)
        except (ET.ParseError, TypeError, ValueError):
            return 0.0

    @staticmethod
    def _chart_soft_edge_radius(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 0.0
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            soft_edge = root.find(
                'c:spPr/a:effectLst/a:softEdge', ns
            )
            if soft_edge is None:
                return 0.0
            return max(
                0.0, float(soft_edge.get('rad', 0.0)) / 12700.0
            )
        except (ET.ParseError, TypeError, ValueError):
            return 0.0

    @staticmethod
    def _chart_title_layout(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, None
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            layout = root.find(
                './/c:title/c:layout/c:manualLayout', ns
            )
            if layout is None:
                return None, None
            x = layout.find('c:x', ns)
            y = layout.find('c:y', ns)
            return (
                float(x.get('val')) if x is not None else None,
                float(y.get('val')) if y is not None else None,
            )
        except (ET.ParseError, TypeError, ValueError):
            return None, None

    @staticmethod
    def _chart_axis_title(chart, scopes):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ''
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            for scope in scopes:
                axis = root.find(scope, ns)
                if axis is None:
                    continue
                title = axis.find('c:title', ns)
                if title is None:
                    continue
                text = ''.join(
                    node.text or '' for node in title.findall('.//a:t', ns)
                ).strip()
                if text:
                    return text
            return ''
        except ET.ParseError:
            return ''

    @staticmethod
    def _chart_data_labels(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, 'bestFit', '404040'
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            labels = root.find('.//c:barChart/c:dLbls', ns)
            if labels is None:
                labels = root.find('.//c:bar3DChart/c:dLbls', ns)
            series_labels = root.findall(
                './/c:barChart/c:ser/c:dLbls', ns
            ) + root.findall('.//c:bar3DChart/c:ser/c:dLbls', ns)
            candidates = ([labels] if labels is not None else []) + series_labels
            if not candidates:
                chart_ex_labels = next((
                    node for node in root.iter()
                    if node.tag.rsplit('}', 1)[-1] == 'dataLabels'
                ), None)
                if chart_ex_labels is None:
                    return False, 'bestFit', '404040'
                visibility = next((
                    node for node in chart_ex_labels
                    if node.tag.rsplit('}', 1)[-1] == 'visibility'
                ), None)
                return (
                    visibility is not None
                    and visibility.get('value', '0')
                    in ('1', 'true', 'True'),
                    chart_ex_labels.get('pos', 'bestFit'),
                    '404040',
                )
            show_value = next(
                (
                    item.find('c:showVal', ns)
                    for item in candidates
                    if item.find('c:showVal', ns) is not None
                ),
                None,
            )
            position = next(
                (
                    item.find('c:dLblPos', ns)
                    for item in candidates
                    if item.find('c:dLblPos', ns) is not None
                ),
                None,
            )
            text_fill = next(
                (
                    fill
                    for item in candidates
                    for fill in [item.find(
                        'c:txPr/a:p/a:pPr/a:defRPr/a:solidFill', ns
                    )]
                    if fill is not None
                ),
                None,
            )
            return (
                show_value is not None
                and show_value.get('val', '0') in ('1', 'true', 'True'),
                position.get('val', 'bestFit')
                if position is not None else 'bestFit',
                PdfLayoutEngine._chart_color(text_fill, '404040'),
            )
        except ET.ParseError:
            return False, 'bestFit', '404040'

    @staticmethod
    def _chart_legend_layout(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None, None
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            layout = root.find(
                './/c:legend/c:layout/c:manualLayout', ns
            )
            if layout is None:
                return None, None
            x = layout.find('c:x', ns)
            y = layout.find('c:y', ns)
            return (
                float(x.get('val')) if x is not None else None,
                float(y.get('val')) if y is not None else None,
            )
        except (ET.ParseError, TypeError, ValueError):
            return None, None

    @staticmethod
    def _chart_series_styles(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            series_nodes = [
                ('line', node)
                for node in root.findall('.//c:lineChart/c:ser', ns)
            ]
            if not series_nodes:
                series_nodes = [
                    ('line', node)
                    for node in root.findall('.//c:line3DChart/c:ser', ns)
                ]
            if not series_nodes:
                series_nodes = [
                    ('line', node)
                    for node in root.findall('.//c:scatterChart/c:ser', ns)
                ]
            stock_nodes = root.findall('.//c:stockChart/c:ser', ns)
            if not series_nodes and stock_nodes:
                series_nodes = [
                    ('bar', node)
                    for node in root.findall('.//c:barChart/c:ser', ns)
                ] + [('line', node) for node in stock_nodes]
            if not series_nodes:
                series_nodes = [
                    ('bar', node)
                    for node in root.findall('.//c:barChart/c:ser', ns)
                ]
            if not series_nodes:
                series_nodes = [
                    ('bar', node)
                    for node in root.findall('.//c:bar3DChart/c:ser', ns)
                ]
            styles = []
            for series_kind, series in series_nodes:
                properties = series.find('c:spPr', ns)
                line = (
                    properties.find('a:ln', ns)
                    if properties is not None else None
                )
                if series_kind == 'bar':
                    fill = (
                        properties.find('a:solidFill', ns)
                        if properties is not None else None
                    )
                else:
                    fill = (
                        line.find('a:solidFill', ns)
                        if line is not None else None
                    )
                width = float(line.get('w', 28575)) / 12700.0 if line is not None else 2.25
                cap = line.get('cap', 'butt') if line is not None else 'butt'
                marker = series.find('c:marker', ns)
                symbol = marker.find('c:symbol', ns) if marker is not None else None
                size = marker.find('c:size', ns) if marker is not None else None
                marker_fill = (
                    marker.find('c:spPr/a:solidFill', ns)
                    if marker is not None else None
                )
                marker_line = (
                    marker.find('c:spPr/a:ln', ns)
                    if marker is not None else None
                )
                marker_line_fill = (
                    marker_line.find('a:solidFill', ns)
                    if marker_line is not None else None
                )
                styles.append((
                    PdfLayoutEngine._chart_color(fill, None),
                    width,
                    'round' if cap == 'rnd' else cap,
                    symbol.get('val', 'none') if symbol is not None else 'none',
                    float(size.get('val', 0.0)) if size is not None else 0.0,
                    PdfLayoutEngine._chart_color(marker_fill, None),
                    PdfLayoutEngine._chart_color(marker_line_fill, None),
                    (
                        float(marker_line.get('w', 0.0)) / 12700.0
                        if marker_line is not None else 0.0
                    ),
                ))
            return tuple(styles)
        except (ET.ParseError, TypeError, ValueError):
            return ()

    @staticmethod
    def _chart_series_line_visibility(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            nodes = root.findall('.//c:scatterChart/c:ser', ns)
            if not nodes:
                nodes = root.findall('.//c:lineChart/c:ser', ns)
            return tuple(
                series.find('c:spPr/a:ln/a:noFill', ns) is None
                for series in nodes
            )
        except ET.ParseError:
            return ()

    @staticmethod
    def _chart_scatter_x_number_format(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return '0'
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            for axis in root.findall('.//c:plotArea/c:valAx', ns):
                position = axis.find('c:axPos', ns)
                if position is None or position.get('val') != 'b':
                    continue
                number_format = axis.find('c:numFmt', ns)
                if number_format is not None:
                    return number_format.get('formatCode', '0')
            return '0'
        except ET.ParseError:
            return '0'

    @staticmethod
    def _chart_scatter_y_number_format(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return '0'
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            for axis in root.findall('.//c:plotArea/c:valAx', ns):
                position = axis.find('c:axPos', ns)
                if position is None or position.get('val') != 'l':
                    continue
                number_format = axis.find('c:numFmt', ns)
                if number_format is not None:
                    return number_format.get('formatCode', '0')
            return '0'
        except ET.ParseError:
            return '0'

    @staticmethod
    def _chart_scatter_axis_title(chart, position):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ''
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            for axis in root.findall('.//c:plotArea/c:valAx', ns):
                axis_position = axis.find('c:axPos', ns)
                if (
                    axis_position is None
                    or axis_position.get('val') != position
                ):
                    continue
                title = axis.find('c:title', ns)
                if title is None:
                    return ''
                text = ''.join(
                    node.text or ''
                    for node in title.findall('c:tx//a:t', ns)
                ).strip()
                return text or 'Axis Title'
            return ''
        except ET.ParseError:
            return ''

    @staticmethod
    def _chart_series_pattern_styles(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            series_nodes = (
                root.findall('.//c:areaChart/c:ser', ns)
                or root.findall('.//c:area3DChart/c:ser', ns)
                or root.findall('.//c:barChart/c:ser', ns)
                or root.findall('.//c:bar3DChart/c:ser', ns)
                or root.findall('.//c:lineChart/c:ser', ns)
                or root.findall('.//c:line3DChart/c:ser', ns)
            )
            styles = []
            for series in series_nodes:
                pattern = series.find('c:spPr/a:pattFill', ns)
                if pattern is None:
                    styles.append((None, None, None))
                    continue
                foreground = pattern.find('a:fgClr', ns)
                background = pattern.find('a:bgClr', ns)
                styles.append((
                    pattern.get('prst'),
                    PdfLayoutEngine._chart_color(foreground, None),
                    PdfLayoutEngine._chart_color(background, None),
                ))
            return tuple(styles)
        except (ET.ParseError, TypeError, ValueError):
            return ()

    @staticmethod
    def _chart_pie_point_styles(chart, point_count, series_index=0):
        palette = (
            '156082', 'E97132', '196B24',
            '0F9ED5', 'A02B93', '4EA72E',
        )
        styles = [
            (palette[index % len(palette)], None, 0.0)
            for index in range(point_count)
        ]
        source = getattr(chart, '_source_chart_xml', None)
        if not source or point_count <= 0:
            return tuple(styles)
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            series_elements = (
                root.findall('.//c:pieChart/c:ser', ns)
                or root.findall('.//c:pie3DChart/c:ser', ns)
                or root.findall('.//c:ofPieChart/c:ser', ns)
                or root.findall('.//c:doughnutChart/c:ser', ns)
            )
            if not 0 <= series_index < len(series_elements):
                return tuple(styles)
            series = series_elements[series_index]
            for point in series.findall('c:dPt', ns):
                index_element = point.find('c:idx', ns)
                if index_element is None:
                    continue
                index = int(index_element.get('val', -1))
                if not 0 <= index < point_count:
                    continue
                properties = point.find('c:spPr', ns)
                fill = (
                    properties.find('a:solidFill', ns)
                    if properties is not None else None
                )
                line = (
                    properties.find('a:ln', ns)
                    if properties is not None else None
                )
                line_fill = (
                    line.find('a:solidFill', ns)
                    if line is not None else None
                )
                no_line = (
                    line is None or line.find('a:noFill', ns) is not None
                )
                styles[index] = (
                    PdfLayoutEngine._chart_color(
                        fill, styles[index][0]
                    ),
                    None if no_line else PdfLayoutEngine._chart_color(
                        line_fill, 'FFFFFF'
                    ),
                    (
                        0.0 if no_line else
                        float(line.get('w', 19050)) / 12700.0
                    ),
                )
            return tuple(styles)
        except (ET.ParseError, TypeError, ValueError):
            return tuple(styles)

    @staticmethod
    def _chart_pie_settings(chart, values):
        source = getattr(chart, '_source_chart_xml', None)
        default = (
            'pie', 0.0, 75.0, (),
            ('0D3A4E', None, 0.0), 'A6A6A6', 0.75
        )
        if not source or not values:
            return default
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            doughnut = root.find('.//c:doughnutChart', ns)
            if doughnut is not None:
                hole = doughnut.find('c:holeSize', ns)
                hole_size = float(
                    hole.get('val', 50.0) if hole is not None else 50.0
                )
                return (
                    'doughnut', max(10.0, min(90.0, hole_size)),
                    75.0, (), ('0D3A4E', None, 0.0), 'A6A6A6', 0.75,
                )
            of_pie = root.find('.//c:ofPieChart', ns)
            if of_pie is None:
                return default
            of_pie_type = of_pie.find('c:ofPieType', ns)
            variant = (
                'pie_of_pie'
                if of_pie_type is None
                or of_pie_type.get('val', 'pie') == 'pie'
                else 'bar_of_pie'
            )
            second_size = of_pie.find('c:secondPieSize', ns)
            secondary_size = float(
                second_size.get('val', 75.0)
                if second_size is not None else 75.0
            )
            split_type_element = of_pie.find('c:splitType', ns)
            split_type = (
                split_type_element.get('val', 'auto')
                if split_type_element is not None else 'auto'
            )
            split_position_element = of_pie.find('c:splitPos', ns)
            split_position = float(
                split_position_element.get('val', 2.0)
                if split_position_element is not None else 2.0
            )
            if split_type == 'cust':
                secondary_indices = tuple(
                    int(point.get('val'))
                    for point in of_pie.findall(
                        'c:custSplit/c:secondPiePt', ns
                    )
                    if point.get('val') is not None
                )
            elif split_type == 'val':
                secondary_indices = tuple(
                    index for index, value in enumerate(values)
                    if float(value) <= split_position
                )
            elif split_type == 'percent':
                total = sum(max(0.0, float(value)) for value in values)
                secondary_indices = tuple(
                    index for index, value in enumerate(values)
                    if total > 0.0
                    and 100.0 * max(0.0, float(value)) / total
                    <= split_position
                )
            else:
                count = max(1, min(len(values) - 1, int(split_position)))
                secondary_indices = tuple(
                    range(len(values) - count, len(values))
                )
            secondary_indices = tuple(
                index for index in secondary_indices
                if 0 <= index < len(values)
            )
            if not secondary_indices or len(secondary_indices) == len(values):
                secondary_indices = tuple(
                    range(max(1, len(values) - 2), len(values))
                )

            point_styles = PdfLayoutEngine._chart_pie_point_styles(
                chart, len(values) + 1
            )
            other_style = point_styles[len(values)]
            series_lines = of_pie.find('c:serLines/c:spPr/a:ln', ns)
            line_fill = (
                series_lines.find('a:solidFill', ns)
                if series_lines is not None else None
            )
            connector_color = PdfLayoutEngine._chart_color(
                line_fill, 'A6A6A6'
            )
            connector_width = (
                float(series_lines.get('w', 9525.0)) / 12700.0
                if series_lines is not None else 0.75
            )
            return (
                variant,
                0.0,
                max(5.0, min(200.0, secondary_size)),
                secondary_indices,
                other_style,
                connector_color,
                connector_width,
            )
        except (ET.ParseError, TypeError, ValueError):
            return default

    @staticmethod
    def _chart_grid_style(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, 'D9D9D9', 'D9D9D9', None
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            category_axis = root.find('.//c:dateAx', ns)
            if category_axis is None:
                category_axis = root.find('.//c:catAx', ns)
            if category_axis is None:
                category_axis = next(
                    (
                        axis for axis in root.findall('.//c:valAx', ns)
                        if axis.find('c:axPos', ns) is not None
                        and axis.find('c:axPos', ns).get('val') == 'b'
                    ),
                    None,
                )
            category_gridlines = (
                category_axis.find('c:majorGridlines', ns)
                if category_axis is not None else None
            )
            show_vertical = category_gridlines is not None
            category_fill = (
                category_gridlines.find(
                    'c:spPr/a:ln/a:solidFill', ns
                )
                if category_gridlines is not None else None
            )
            value_axis = next(
                (
                    axis for axis in root.findall('.//c:valAx', ns)
                    if axis.find('c:axPos', ns) is not None
                    and axis.find('c:axPos', ns).get('val') == 'l'
                ),
                root.find('.//c:valAx', ns),
            )
            value_gridlines = (
                value_axis.find('c:majorGridlines', ns)
                if value_axis is not None else None
            )
            value_fill = (
                value_gridlines.find('c:spPr/a:ln/a:solidFill', ns)
                if value_gridlines is not None else None
            )
            plot_fill = root.find(
                './/c:plotArea/c:spPr/a:ln/a:solidFill', ns
            )
            return (
                show_vertical,
                PdfLayoutEngine._chart_color(
                    category_fill, 'D9D9D9'
                ),
                PdfLayoutEngine._chart_color(value_fill, 'D9D9D9'),
                PdfLayoutEngine._chart_color(plot_fill, None),
            )
        except ET.ParseError:
            return False, 'D9D9D9', 'D9D9D9', None

    @staticmethod
    def _chart_value_axis_style(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False, 'F2F2F2', None, 'none', 'none'
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            axis = root.find('.//c:valAx', ns)
            if axis is None:
                return False, 'F2F2F2', None, 'none', 'none'
            minor = axis.find('c:minorGridlines', ns)
            minor_fill = (
                minor.find('c:spPr/a:ln/a:solidFill', ns)
                if minor is not None else None
            )
            axis_fill = axis.find('c:spPr/a:ln/a:solidFill', ns)
            major_tick = axis.find('c:majorTickMark', ns)
            minor_tick = axis.find('c:minorTickMark', ns)
            return (
                minor is not None,
                PdfLayoutEngine._chart_color(minor_fill, 'F2F2F2'),
                PdfLayoutEngine._chart_color(axis_fill, None),
                major_tick.get('val', 'none') if major_tick is not None else 'none',
                minor_tick.get('val', 'none') if minor_tick is not None else 'none',
            )
        except ET.ParseError:
            return False, 'F2F2F2', None, 'none', 'none'

    @staticmethod
    def _chart_value_axis_display_unit(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 1.0, '', None, None
        units = {
            'hundreds': (100.0, 'Hundreds'),
            'thousands': (1000.0, 'Thousands'),
            'tenThousands': (10000.0, 'Ten Thousands'),
            'hundredThousands': (100000.0, 'Hundred Thousands'),
            'millions': (1000000.0, 'Millions'),
            'tenMillions': (10000000.0, 'Ten Millions'),
            'hundredMillions': (100000000.0, 'Hundred Millions'),
            'billions': (1000000000.0, 'Billions'),
            'trillions': (1000000000000.0, 'Trillions'),
        }
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            }
            display_units = root.find('.//c:valAx/c:dispUnits', ns)
            if display_units is None:
                return 1.0, '', None, None
            built_in = display_units.find('c:builtInUnit', ns)
            if built_in is None:
                return 1.0, '', None, None
            divisor, label = units.get(
                built_in.get('val'), (1.0, '')
            )
            display_label = display_units.find('c:dispUnitsLbl', ns)
            if display_label is None:
                label = ''
                return divisor, label, None, None
            manual = display_label.find('c:layout/c:manualLayout', ns)
            layout_x = None
            layout_y = None
            if manual is not None:
                x_mode = manual.find('c:xMode', ns)
                y_mode = manual.find('c:yMode', ns)
                x_value = manual.find('c:x', ns)
                y_value = manual.find('c:y', ns)
                if (
                    x_value is not None
                    and (x_mode is None or x_mode.get('val') == 'edge')
                ):
                    layout_x = float(x_value.get('val'))
                if (
                    y_value is not None
                    and (y_mode is None or y_mode.get('val') == 'edge')
                ):
                    layout_y = float(y_value.get('val'))
            return divisor, label, layout_x, layout_y
        except (ET.ParseError, TypeError, ValueError):
            return 1.0, '', None, None

    @staticmethod
    def _chart_value_axis_log_base(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return None
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            }
            element = root.find('.//c:valAx/c:scaling/c:logBase', ns)
            if element is None:
                return None
            base = float(element.get('val'))
            return base if 2.0 <= base <= 1000.0 else None
        except (ET.ParseError, TypeError, ValueError):
            return None

    @staticmethod
    def _chart_value_axis_position(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return 'b'
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            }
            element = root.find('.//c:valAx/c:axPos', ns)
            return element.get('val', 'b') if element is not None else 'b'
        except ET.ParseError:
            return 'b'

    @staticmethod
    def _chart_color(fill, default):
        import colorsys

        if fill is None:
            return default
        drawing_ns = (
            'http://schemas.openxmlformats.org/drawingml/2006/main'
        )
        rgb_tag = f'{{{drawing_ns}}}srgbClr'
        scheme_tag = f'{{{drawing_ns}}}schemeClr'
        rgb = fill if fill.tag == rgb_tag else fill.find(rgb_tag)
        if rgb is not None and rgb.get('val'):
            return rgb.get('val').upper()
        scheme = fill if fill.tag == scheme_tag else fill.find(scheme_tag)
        chart_colors = {
            'tx1': '000000', 'bg1': 'FFFFFF',
            'accent1': '156082', 'accent2': 'E97132',
            'accent3': '196B24', 'accent4': '0F9ED5',
            'accent5': 'A02B93', 'accent6': '4EA72E',
        }
        if scheme is not None:
            color = chart_colors.get(scheme.get('val'))
            if color is None:
                return default
            channels = [
                int(color[index:index + 2], 16) / 255.0
                for index in (0, 2, 4)
            ]
            lum_mod = scheme.find(
                '{http://schemas.openxmlformats.org/drawingml/2006/main}lumMod'
            )
            lum_off = scheme.find(
                '{http://schemas.openxmlformats.org/drawingml/2006/main}lumOff'
            )
            modifier = float(lum_mod.get('val', 100000)) / 100000.0 if lum_mod is not None else 1.0
            offset = float(lum_off.get('val', 0)) / 100000.0 if lum_off is not None else 0.0
            hue, luminance, saturation = colorsys.rgb_to_hls(*channels)
            luminance = max(0.0, min(1.0, luminance * modifier + offset))
            transformed = colorsys.hls_to_rgb(hue, luminance, saturation)
            return ''.join(f'{round(value * 255):02X}' for value in transformed)
        return default

    @staticmethod
    def _chart_map_value_colors(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {
                'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
            }
            colors = root.find('.//cx:valueColors', ns)
            if colors is None:
                return ()
            result = []
            map_scheme_colors = {
                'accent1': '4472C4',
                'accent6': '70AD47',
            }
            for name in ('minColor', 'midColor', 'maxColor'):
                element = colors.find(f'cx:{name}', ns)
                scheme = (
                    element.find(
                        '{http://schemas.openxmlformats.org/drawingml/2006/main}schemeClr'
                    ) if element is not None else None
                )
                result.append(
                    map_scheme_colors.get(scheme.get('val'))
                    if scheme is not None else
                    PdfLayoutEngine._chart_color(element, None)
                )
            return tuple(result) if all(result) else ()
        except ET.ParseError:
            return ()

    @staticmethod
    def _chart_surface_band_colors(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return ()
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            }
            colors = []
            for band in root.findall('.//c:surfaceChart/c:bandFmts/c:bandFmt', ns):
                fill = band.find('c:spPr/a:ln/a:solidFill', ns)
                if fill is None:
                    fill = band.find('c:spPr/a:solidFill', ns)
                color = PdfLayoutEngine._chart_color(fill, None)
                if color:
                    colors.append(color)
            return tuple(colors)
        except ET.ParseError:
            return ()

    @staticmethod
    def _chart_surface_has_general_axis(chart):
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return False
        try:
            root = ET.fromstring(source)
            ns = {
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
            }
            value_format = root.find('.//c:valAx/c:numFmt', ns)
            return (
                value_format is not None
                and value_format.get('formatCode', '').lower() == 'general'
            )
        except ET.ParseError:
            return False

    @staticmethod
    def _chart_plot_ratios(chart, width, height):
        defaults = (0.16, 0.18, 0.78, 0.68)
        if chart.type == ChartType.PIE:
            top_reserve = min(37.09, max(0.0, height))
            bottom_reserve = min(33.87, max(0.0, height - top_reserve))
            if bool(getattr(chart, 'is_3d', False)):
                side_reserve = min(22.02, max(0.0, width))
                return (
                    10.92 / max(1e-12, width),
                    top_reserve / max(1e-12, height),
                    max(0.0, width - side_reserve) / max(1e-12, width),
                    max(
                        0.0, height - top_reserve - bottom_reserve
                    ) / max(1e-12, height),
                )
            radius = max(0.0, min(
                width / 2.0,
                (height - top_reserve - bottom_reserve) / 2.0,
            ))
            return (
                (width / 2.0 - radius) / max(1e-12, width),
                top_reserve / max(1e-12, height),
                2.0 * radius / max(1e-12, width),
                2.0 * radius / max(1e-12, height),
            )
        source = getattr(chart, '_source_chart_xml', None)
        if not source:
            return (
                50.75 / width, 37.09 / height,
                (width - 61.75) / width, (height - 98.59) / height,
            )
        try:
            root = ET.fromstring(source)
            ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
            layout = root.find(
                './/c:plotArea/c:layout/c:manualLayout', ns
            )
            if (
                layout is None
                or any(
                    layout.find(f'c:{name}', ns) is None
                    for name in ('x', 'y', 'w', 'h')
                )
            ):
                has_category_title = bool(PdfLayoutEngine._chart_axis_title(
                    chart, ('.//c:dateAx', './/c:catAx')
                ))
                has_value_title = bool(PdfLayoutEngine._chart_axis_title(
                    chart, ('.//c:valAx',)
                ))
                if chart.type == ChartType.BAR:
                    bottom_reserve = 83.08
                elif chart.type == ChartType.AREA:
                    bottom_reserve = 83.08
                elif chart.type == ChartType.STOCK:
                    # Stock charts keep date labels horizontal, so they use
                    # the compact category-axis band rather than the rotated
                    # date-label reserve used by ordinary line charts.
                    bottom_reserve = 83.08
                else:
                    bottom_reserve = (
                        98.59 if PdfLayoutEngine._chart_has_date_axis(chart)
                        else 83.08
                    )
                is_horizontal_bar = (
                    chart.type == ChartType.BAR
                    and getattr(chart, 'bar_direction', 'col') == 'bar'
                )
                value_axis_at_top = (
                    is_horizontal_bar
                    and PdfLayoutEngine._chart_value_axis_position(chart) == 't'
                )
                if (
                    value_axis_at_top
                ):
                    top_reserve = 49.21
                elif (
                    is_horizontal_bar
                    and PdfLayoutEngine._chart_value_axis_log_base(chart)
                ):
                    top_reserve = 36.90
                else:
                    top_reserve = 37.09
                stock_axes_hidden = (
                    chart.type == ChartType.STOCK
                    and PdfLayoutEngine._chart_stock_features(chart)[3]
                )
                if stock_axes_hidden:
                    # Deleted stock axes leave only a narrow frame inset;
                    # titles and right legends overlay the expanded plot.
                    top_reserve = 10.75
                if is_horizontal_bar:
                    # Horizontal value labels extend past the last gridline,
                    # while category labels need less room than value labels.
                    left_reserve = 41.29
                    right_reserve = (
                        21.26
                        if PdfLayoutEngine._chart_grouping(chart)
                        == 'percentStacked'
                        else 28.97
                    )
                elif PdfLayoutEngine._chart_grouping(chart) == 'percentStacked':
                    left_reserve = (
                        35.33 if chart.type == ChartType.AREA else 35.35
                    )
                    right_reserve = None
                else:
                    left_reserve = 50.75
                    right_reserve = None
                if (
                    chart.type == ChartType.STOCK
                    and not stock_axes_hidden
                    and PdfLayoutEngine._chart_stock_has_general_axes(chart)
                ):
                    has_volume = PdfLayoutEngine._chart_stock_features(chart)[2]
                    left_reserve = 40.5 if has_volume else 29.0
                    if has_volume:
                        right_reserve = 29.0
                if has_category_title:
                    bottom_reserve += 22.0
                if has_value_title:
                    left_reserve += 22.0
                if right_reserve is None:
                    if chart.type == ChartType.AREA:
                        right_reserve = (
                            22.89
                            if PdfLayoutEngine._chart_has_date_axis(chart)
                            else 28.83
                        )
                    else:
                        right_reserve = (
                            10.0
                            if has_category_title or has_value_title
                            else 11.0
                        )
                bottom_plot_reserve = bottom_reserve - 37.09
                if stock_axes_hidden:
                    bottom_plot_reserve = 10.75
                elif (
                    chart.type == ChartType.STOCK
                    and PdfLayoutEngine._chart_stock_has_general_axes(chart)
                ):
                    bottom_plot_reserve = 72.0
                if value_axis_at_top:
                    # Moving the value labels above the plot frees the lower
                    # area between the bars and the bottom legend.
                    bottom_plot_reserve = 33.88
                return (
                    left_reserve / width, top_reserve / height,
                    (width - left_reserve - right_reserve) / width,
                    (height - top_reserve - bottom_plot_reserve) / height,
                )
            values = []
            for name, fallback in zip(('x', 'y', 'w', 'h'), defaults):
                element = layout.find(f'c:{name}', ns)
                values.append(float(element.get('val')) if element is not None else fallback)
            return tuple(values)
        except (ET.ParseError, TypeError, ValueError):
            return defaults

    @staticmethod
    def _chart_title(chart, automatic_title=None):
        if chart.title:
            return str(chart.title)
        source = getattr(chart, '_source_chart_xml', None)
        is_doughnut = False
        if source:
            try:
                root = ET.fromstring(source)
                ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
                is_doughnut = root.find('.//c:doughnutChart', ns) is not None
                value = root.find('.//c:title/c:tx//c:v', ns)
                if value is not None and value.text:
                    return value.text
                title = root.find('.//c:chart/c:title', ns)
                if (
                    title is not None
                    and automatic_title
                    and chart.type == ChartType.SCATTER
                    and (
                        (
                            getattr(chart, '_is_bubble', False)
                            and len(chart.n_series) == 1
                        )
                        or not chart.show_legend
                    )
                ):
                    return str(automatic_title)
            except ET.ParseError:
                pass
        if (
            chart.type == ChartType.PIE
            and automatic_title
            and not is_doughnut
            and not getattr(chart, '_auto_title_deleted', False)
        ):
            return str(automatic_title)
        return 'Chart Title'

    @staticmethod
    def _date_axis_layout(categories, reverse=False):
        parsed = []
        try:
            for label in categories:
                date = datetime.strptime(str(label), '%b-%y')
                parsed.append(date.year * 12 + date.month - 1)
        except ValueError:
            positions = tuple(
                (index + 0.5) / max(1, len(categories))
                for index in range(len(categories))
            )
            return positions, tuple(zip(categories, positions))
        if not parsed:
            return (), ()
        first = min(parsed)
        last = max(parsed)
        month_count = last - first + 1
        positions = tuple(
            (month - first + 0.5) / month_count for month in parsed
        )
        labels = []
        for offset in range(month_count):
            month = first + offset
            label = datetime(month // 12, month % 12 + 1, 1).strftime('%b-%y')
            labels.append((label, (offset + 0.5) / month_count))
        if reverse:
            positions = tuple(1.0 - position for position in positions)
            labels = [
                (label, 1.0 - position)
                for label, position in reversed(labels)
            ]
        return positions, tuple(labels)

    @staticmethod
    def _nice_chart_step(value):
        if value <= 0:
            return 1.0
        magnitude = 10.0 ** math.floor(math.log10(value))
        normalized = value / magnitude
        for candidate in (1.0, 2.0, 5.0, 10.0):
            if normalized <= candidate:
                return candidate * magnitude
        return 10.0 * magnitude

    @classmethod
    def _materialize_shapes(
        cls, worksheet, page_width, page_height, left, top, scale,
        row_heights, column_widths, row_chunk, column_chunk,
    ):
        if not worksheet.shapes:
            return ()
        default_column_width = RowHeightCalculator.column_width_to_points(
            worksheet, 1
        )
        default_row_height = worksheet.get_effective_default_row_height()
        source_column_width = (
            float(worksheet.properties.format.base_col_width)
            * worksheet.measure_text(
                "0", style=worksheet.get_effective_style(row=1, column=1)
            ).mdw_points
        )
        source_row_height = float(
            worksheet.properties.format.default_row_height
        )
        horizontal_ratio = default_column_width / source_column_width
        vertical_ratio = default_row_height / source_row_height
        column_origin = sum(
            RowHeightCalculator.column_width_to_points(worksheet, column)
            for column in range(1, column_chunk[0])
        )
        if (
            not worksheet.cells.get_all_cells()
            and not getattr(worksheet.page_setup, '_source_present', False)
        ):
            bounds = worksheet.get_content_bounds(include_drawings=True)
            if bounds is not None:
                column_origin += sum(
                    RowHeightCalculator.column_width_to_points(worksheet, column)
                    for column in range(1, max(1, bounds.min_col - 1))
                )
        row_origin = sum(
            row_heights.get(row, default_row_height)
            for row in range(1, row_chunk[0])
        )
        clip_right = left + sum(
            RowHeightCalculator.column_width_to_points(worksheet, column)
            for column in range(column_chunk[0], column_chunk[1] + 1)
        ) * scale
        clip_bottom = top + sum(
            row_heights.get(row, default_row_height)
            for row in range(row_chunk[0], row_chunk[1] + 1)
        ) * scale
        workbook = getattr(worksheet, '_workbook', None)
        map_chart_sheet = any(
            chart.type == ChartType.MAP for chart in worksheet.charts
        )
        shape_boxes = {}

        def grouped_shape_box(shape):
            # Group children have already been transformed into drawing-
            # absolute EMUs by the XML loader. Use the same calibrated drawing
            # canvas mapping as grouped pictures in chartEx worksheets.
            wide_map_table = RowHeightCalculator.column_width_to_points(
                worksheet, 3
            ) > 0.0
            base_x = 34.6 if wide_map_table else 42.0
            page_origins = (
                (0.0, 482.76, 1013.88)
                if wide_map_table else (0.0, 491.03, 996.58)
            )
            page_index = 0 if column_chunk[0] <= 1 else (
                1 if column_chunk[0] <= 12 else 2
            )
            return (
                (
                    shape._transform_x / 12700.0
                    + base_x - page_origins[page_index]
                ) * scale,
                (shape._transform_y / 12700.0 + 93.5) * scale,
                shape._transform_width / 12700.0 * 0.9592 * scale,
                shape._transform_height / 12700.0 * 0.994 * scale,
            )

        for source_shape in worksheet.shapes:
            if source_shape._transform_x is None:
                continue
            if getattr(source_shape, '_transform_is_group_absolute', False):
                shape_boxes[source_shape._shape_id] = grouped_shape_box(
                    source_shape
                )
            else:
                shape_boxes[source_shape._shape_id] = (
                    left + (
                        source_shape._transform_x / 12700.0 * horizontal_ratio
                        - column_origin
                    ) * scale,
                    top + (
                        source_shape._transform_y / 12700.0 * vertical_ratio
                        - row_origin
                    ) * scale,
                    source_shape._transform_width / 12700.0 * horizontal_ratio * scale,
                    source_shape._transform_height / 12700.0 * vertical_ratio * scale,
                )
        layouts = []
        for shape in worksheet.shapes:
            if shape.is_hidden or shape._transform_x is None:
                continue
            if getattr(shape, '_transform_is_group_absolute', False):
                x, y, width, height = grouped_shape_box(shape)
            elif map_chart_sheet and shape.hyperlink:
                if not (
                    column_chunk[0]
                    <= shape._upper_left_column
                    <= column_chunk[1]
                ):
                    continue
                # Excel positions chart-navigation shapes from the chartEx
                # drawing canvas rather than its fallback xfrm coordinates.
                x = left + 92.73 * scale
                y = top + 675.24 * scale
                width = 69.48 * scale
                height = 24.06 * scale
            else:
                shape_left = shape._transform_x / 12700.0 * horizontal_ratio
                shape_top = shape._transform_y / 12700.0 * vertical_ratio
                shape_width = (
                    shape._transform_width / 12700.0 * horizontal_ratio
                )
                shape_height = (
                    shape._transform_height / 12700.0 * vertical_ratio
                )
                x = left + (shape_left - column_origin) * scale
                y = top + (shape_top - row_origin) * scale
                width = shape_width * scale
                height = shape_height * scale
            if (
                x >= clip_right or y >= clip_bottom
                or x + width <= left or y + height <= top
            ):
                continue
            resolution = (
                workbook.font_strategy.resolve(shape.font)
                if workbook is not None and shape.text else None
            )
            connector_start = cls._shape_connection_point(
                shape_boxes.get(shape._start_connection_id),
                shape._start_connection_index,
            )
            connector_end = cls._shape_connection_point(
                shape_boxes.get(shape._end_connection_id),
                shape._end_connection_index,
            )
            layouts.append(PdfShapeLayout(
                name=shape.name,
                preset_geometry=shape._preset_geometry,
                is_connector=shape._is_connector,
                rotation=shape._rotation,
                flip_horizontal=shape._flip_horizontal,
                flip_vertical=shape._flip_vertical,
                tail_end=shape._tail_end,
                x=x,
                y=y,
                width=width,
                height=height,
                fill_visible=int(shape.fill.fill_type) != 0,
                fill_color=shape.fill.fore_color,
                line_visible=shape.line.is_visible,
                line_color=shape.line.color,
                line_width=(shape.line.weight / 12700.0) * scale,
                text=shape.text,
                text_wrapped=shape.is_text_wrapped,
                text_direction=shape._text_direction,
                text_horizontal_alignment=shape.text_horizontal_alignment,
                text_vertical_alignment=shape.text_vertical_alignment,
                font=shape.font.copy(),
                font_resolution=resolution,
                connector_start=connector_start,
                connector_end=connector_end,
            ))
        return tuple(layouts)

    @staticmethod
    def _shape_connection_point(box, index):
        if box is None or index is None:
            return None
        x, y, width, height = box
        points = (
            (x + width / 2.0, y),
            (x, y + height / 2.0),
            (x + width / 2.0, y + height),
            (x + width, y + height / 2.0),
        )
        return points[int(index) % len(points)]

    @classmethod
    def _line_runs(cls, worksheet, lines, style, scale):
        workbook = getattr(worksheet, '_workbook', None)
        result = []
        text_width = 0.0
        run_font = copy.copy(style.font)
        source_size = float(run_font.size)
        vertical_alignment = getattr(
            run_font, 'vertical_alignment', 'baseline'
        )
        baseline_shift = 0.0
        if vertical_alignment in ('superscript', 'subscript'):
            run_font.size = source_size * 2.0 / 3.0
        if vertical_alignment == 'superscript':
            baseline_shift = -source_size * 0.34
        elif vertical_alignment == 'subscript':
            baseline_shift = source_size * 0.18

        run_style = copy.copy(style)
        run_style.font = run_font
        for line in lines:
            measured = worksheet.measure_text(line, style=run_style)
            text_width = max(text_width, measured.width_points)
            line_result = []
            for run in measured.runs:
                fallback_font = copy.copy(run_font)
                fallback_font.name = run.family
                resolution = (
                    workbook.font_strategy.resolve(fallback_font)
                    if workbook is not None
                    else measured.font_resolution._replace(
                        requested_family=run.family,
                        resolved_family=run.family,
                    )
                )
                line_result.append(PdfTextRunLayout(
                    text=run.text,
                    width=run.width_points * scale,
                    font_resolution=resolution,
                    font_size=float(run_font.size),
                    color=run_font.color,
                    underline=bool(run_font.underline),
                    underline_type=cls._font_underline_type(run_font),
                    strikethrough=bool(run_font.strikethrough),
                    baseline_shift=baseline_shift,
                ))
            result.append(tuple(line_result))
        return tuple(result), text_width

    @staticmethod
    def _rich_line_runs(worksheet, rich_runs, style, scale):
        workbook = getattr(worksheet, '_workbook', None)
        line_runs = []
        total_width = 0.0
        line_height = 0.0
        for source_run in rich_runs:
            run_font = copy.copy(style.font)
            for source_name, target_name in (
                ('font_name', 'name'),
                ('font_size', 'size'),
                ('font_color', 'color'),
                ('bold', 'bold'),
                ('italic', 'italic'),
                ('underline', 'underline'),
                ('underline_type', 'underline_type'),
                ('strikethrough', 'strikethrough'),
            ):
                value = getattr(source_run, source_name)
                if value is not None:
                    setattr(run_font, target_name, value)

            source_size = float(run_font.size)
            vertical_alignment = source_run.vertical_alignment
            if vertical_alignment in ('superscript', 'subscript'):
                run_font.size = source_size * 2.0 / 3.0
            baseline_shift = 0.0
            if vertical_alignment == 'superscript':
                baseline_shift = -source_size * 0.5
            elif vertical_alignment == 'subscript':
                baseline_shift = source_size * 0.2

            run_style = copy.copy(style)
            run_style.font = run_font
            measured = worksheet.measure_text(source_run.text, style=run_style)
            italic_overhang = 0.0
            synthetic_italic = False
            if run_font.italic:
                normal_font = copy.copy(run_font)
                normal_font.italic = False
                normal_resolution = (
                    workbook.font_strategy.resolve(normal_font)
                    if workbook is not None else measured.font_resolution
                )
                synthetic_italic = (
                    measured.font_resolution.font_path
                    == normal_resolution.font_path
                )
            if synthetic_italic:
                # Excel reserves the visible overhang created by synthetic
                # italic shear before positioning the following rich-text run.
                italic_overhang = float(run_font.size) * 0.1725
            total_width += measured.width_points + italic_overhang
            line_height = max(line_height, measured.line_height_points)
            for run_index, measured_run in enumerate(measured.runs):
                fallback_font = copy.copy(run_font)
                fallback_font.name = measured_run.family
                resolution = (
                    workbook.font_strategy.resolve(fallback_font)
                    if workbook is not None
                    else measured.font_resolution._replace(
                        requested_family=measured_run.family,
                        resolved_family=measured_run.family,
                    )
                )
                line_runs.append(PdfTextRunLayout(
                    text=measured_run.text,
                    width=measured_run.width_points * scale,
                    font_resolution=resolution,
                    font_size=float(run_font.size),
                    color=run_font.color,
                    underline=bool(run_font.underline),
                    underline_type=PdfLayoutEngine._font_underline_type(
                        run_font
                    ),
                    strikethrough=bool(run_font.strikethrough),
                    baseline_shift=baseline_shift,
                    advance_adjustment=(
                        italic_overhang
                        if run_index == len(measured.runs) - 1 else 0.0
                    ),
                    synthetic_italic=synthetic_italic,
                ))
        return (tuple(line_runs),), total_width, line_height

    @staticmethod
    def _font_underline_type(font):
        if not bool(getattr(font, 'underline', False)):
            return 'none'
        underline_type = getattr(font, 'underline_type', None)
        return underline_type if underline_type not in (None, 'none') else 'single'

    @staticmethod
    def _hyperlink_map(worksheet):
        result = {}
        for hyperlink in worksheet.hyperlinks:
            target = hyperlink.address
            if not target:
                continue
            range_ref = str(hyperlink.range or '').replace('$', '')
            start_ref, separator, end_ref = range_ref.partition(':')
            end_ref = end_ref if separator else start_ref
            try:
                min_row, min_column = worksheet.cells.coordinate_from_string(
                    start_ref
                )
                max_row, max_column = worksheet.cells.coordinate_from_string(
                    end_ref
                )
            except ValueError:
                continue
            for row in range(min_row, max_row + 1):
                for column in range(min_column, max_column + 1):
                    result[(row, column)] = target
        return result

    @classmethod
    def _text_clip_width(
        cls, worksheet, row, column, merge_columns, page_columns,
        column_widths, scale, width, measured_width, indent_width,
        max_clip_width, style, all_cells, covered, merges,
    ):
        alignment = style.alignment
        if (
            alignment.wrap_text
            or alignment.shrink_to_fit
            or int(alignment.text_rotation or 0)
            or alignment.horizontal not in ('general', 'left')
        ):
            return width

        required_width = measured_width + indent_width + 2.0 * scale
        if required_width <= width:
            return width

        clip_width = width
        last_column = merge_columns[-1]
        try:
            start = page_columns.index(last_column) + 1
        except ValueError:
            return width

        previous = last_column
        blocked = False
        for next_column in page_columns[start:]:
            if next_column != previous + 1:
                blocked = True
                break
            if (row, next_column) in covered or (row, next_column) in merges:
                blocked = True
                break
            ref = worksheet.cells.coordinate_to_string(row, next_column)
            neighbor = all_cells.get(ref)
            if neighbor is not None and neighbor.get_display_text() != '':
                blocked = True
                break
            clip_width += column_widths[next_column] * scale
            if clip_width >= required_width:
                return clip_width
            previous = next_column
        if (
            not blocked
            and clip_width < required_width
            and not worksheet.print_area
        ):
            clip_width = min(max_clip_width, required_width)
        return clip_width

    @classmethod
    def _resolve_scale(
        cls, worksheet, area, row_heights, column_widths,
        usable_width, usable_height, options,
    ):
        setup = worksheet.page_setup
        candidates = []
        total_width = sum(column_widths[col] for col in range(area.min_col, area.max_col + 1))
        total_height = sum(row_heights[row] for row in range(area.min_row, area.max_row + 1))
        if setup.fit_to_width is not None and int(setup.fit_to_width) > 0 and total_width:
            candidates.append(usable_width * int(setup.fit_to_width) / total_width)
        if setup.fit_to_height is not None and int(setup.fit_to_height) > 0 and total_height:
            candidates.append(usable_height * int(setup.fit_to_height) / total_height)
        if candidates:
            scale = min(1.0, min(candidates))
        else:
            scale = float(setup.scale or 100) / 100.0
        return max(float(options.min_scale), scale)

    @staticmethod
    def _split_axis(start, end, sizes, capacity, title_range, manual_breaks):
        chunks = []
        current_start = start
        current_size = 0.0
        title_size = 0.0
        if title_range:
            title_size = sum(sizes.get(index, 0.0) for index in range(title_range[0], title_range[1] + 1))
        index = start
        while index <= end:
            available = capacity - (title_size if chunks else 0.0)
            available = max(0.0, available)
            forced_break = index > current_start and (index - 1) in manual_breaks
            next_size = sizes.get(index, 0.0)
            if index > current_start and (forced_break or current_size + next_size > available):
                chunks.append((current_start, index - 1))
                current_start = index
                current_size = 0.0
                continue
            current_size += next_size
            index += 1
        chunks.append((current_start, end))
        return chunks

    @staticmethod
    def _repeat_indexes(title_range, chunk, chunk_index):
        if not title_range or chunk_index == 0:
            return ()
        return tuple(
            index for index in range(title_range[0], title_range[1] + 1)
            if not (chunk[0] <= index <= chunk[1])
        )

    @staticmethod
    def _positions(indexes, sizes, offset, scale):
        result = {}
        position = offset
        for index in indexes:
            result[index] = position
            position += sizes.get(index, 0.0) * scale
        return result

    @staticmethod
    def _union_titles(area, rows, columns):
        return RenderBounds(
            min(area.min_row, rows[0] if rows else area.min_row),
            min(area.min_col, columns[0] if columns else area.min_col),
            max(area.max_row, rows[1] if rows else area.max_row),
            max(area.max_col, columns[1] if columns else area.max_col),
        )

    @classmethod
    def _page_size(cls, worksheet, options):
        if options.page_size is not None:
            return options.page_size
        width, height = cls.PAPER_SIZES.get(
            worksheet.page_setup.paper_size, cls.PAPER_SIZES[1]
        )
        if worksheet.page_setup.orientation == 'landscape' and height > width:
            width, height = height, width
        elif worksheet.page_setup.orientation == 'portrait' and width > height:
            width, height = height, width
        return float(width), float(height)

    @staticmethod
    def _merge_map(worksheet):
        covered = set()
        origins = {}
        for ref in worksheet.merged_cells:
            start, _, end = ref.partition(':')
            end = end or start
            min_row, min_col = worksheet.cells.coordinate_from_string(start)
            max_row, max_col = worksheet.cells.coordinate_from_string(end)
            origins[(min_row, min_col)] = (min_row, min_col, max_row, max_col)
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    if (row, column) != (min_row, min_col):
                        covered.add((row, column))
        return covered, origins
