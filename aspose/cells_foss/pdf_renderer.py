"""Skia implementation of the prepared PDF page renderer."""
import json
import math
import os

from .pdf_renderer_3d_area import PdfRenderer3DAreaMixin
from .pdf_renderer_3d_bar import PdfRenderer3DBarMixin
from .pdf_renderer_3d_line_column import PdfRenderer3DLineColumnMixin
from .pdf_renderer_drawing import PdfRendererDrawingMixin
from .pdf_renderer_pie_bar import PdfRendererPieBarMixin
from .pdf_renderer_special_charts import PdfRendererSpecialChartsMixin
from .pdf_renderer_stock_surface import PdfRendererStockSurfaceMixin


class SkiaPdfRenderer(PdfRendererStockSurfaceMixin, PdfRendererSpecialChartsMixin, PdfRendererPieBarMixin, PdfRenderer3DBarMixin, PdfRenderer3DAreaMixin, PdfRenderer3DLineColumnMixin, PdfRendererDrawingMixin):
    """Draws materialized page layouts without owning layout or pagination."""

    _SYNTHETIC_ITALIC_SKEW = -1.0 / 3.0
    _MAP_FEATURES = None

    _DENSITY_PATTERN_RATIOS = {
        'gray0625': 0.03125,
        'gray125': 0.0625,
        'mediumGray': 0.25,
        'darkGray': 0.5,
    }
    _BITMAP_PATTERN_MASKS = {
        'gray0625': frozenset(((0, 0), (4, 4))),
        'gray125': frozenset(((0, 0), (4, 2), (0, 4), (4, 6))),
        'mediumGray': frozenset(
            (x, y)
            for y in range(8)
            for x in range(8)
            if x % 4 == (0 if y % 2 == 0 else 2)
        ),
        'darkGray': frozenset(
            (x, y)
            for y in range(8)
            for x in range(8)
            if (x + y) % 2 == 0
        ),
        'darkHorizontal': frozenset(
            (x, y) for y in range(4) for x in range(8)
        ),
        'darkVertical': frozenset(
            (x, y) for y in range(8) for x in range(4)
        ),
        'darkDown': frozenset(
            (x, y) for y in range(8) for x in range(8)
            if (x - y) % 8 < 4
        ),
        'darkUp': frozenset(
            (x, y) for y, row in enumerate((
                '..####..', '.####...', '####....', '###....#',
                '##....##', '#....###', '....####', '...####.',
            )) for x, pixel in enumerate(row) if pixel == '#'
        ),
        'darkGrid': frozenset(
            (x, y) for y in range(8) for x in range(8)
            if (y < 4 and x < 4) or (y >= 4 and x >= 4)
        ),
        'darkTrellis': frozenset(
            (x, y) for y, row in enumerate((
                '########', '.######.', '..####..', '.######.',
                '########', '###..###', '##....##', '###..###',
            )) for x, pixel in enumerate(row) if pixel == '#'
        ),
        'lightHorizontal': frozenset(
            (x, y) for y in range(2) for x in range(8)
        ),
        'lightVertical': frozenset(
            (x, y) for y in range(8) for x in range(2)
        ),
        'lightDown': frozenset(
            (x, y) for y, row in enumerate((
                '##......', '.##.....', '..##....', '...##...',
                '....##..', '.....##.', '......##', '#......#',
            )) for x, pixel in enumerate(row) if pixel == '#'
        ),
        'lightUp': frozenset(
            (x, y) for y, row in enumerate((
                '...##...', '..##....', '.##.....', '##......',
                '#......#', '......##', '.....##.', '....##..',
            )) for x, pixel in enumerate(row) if pixel == '#'
        ),
        'lightGrid': frozenset(
            (x, y) for y in range(8) for x in range(8)
            if y < 2 or x < 2
        ),
        'lightTrellis': frozenset(
            (x, y)
            for y, row in enumerate((
                '##.##...',
                '.###....',
                '.###....',
                '##.##...',
                '#...##.#',
                '.....###',
                '.....###',
                '#...##.#',
            ))
            for x, pixel in enumerate(row)
            if pixel == '#'
        ),
    }
    _CHART_PATTERN_ROWS = {
        'ltUpDiag': (
            '#...#...',
            '...#...#',
            '..#...#.',
            '.#...#..',
            '#...#...',
            '...#...#',
            '..#...#.',
            '.#...#..',
        ),
        'weave': (
            '##......##......',
            '##......##......',
            '..##..##..##....',
            '..##..##..##....',
            '....##......##..',
            '....##......##..',
            '..##......##..##',
            '..##......##..##',
            '##......##......',
            '##......##......',
            '......##..##....',
            '......##..##....',
            '....##......##..',
            '....##......##..',
            '..##..##......##',
            '..##..##......##',
        ),
    }

    def __init__(self):
        try:
            import skia
        except ImportError as exc:  # pragma: no cover - packaging failure path
            raise RuntimeError(
                "PDF export requires skia-python; install the package dependencies"
            ) from exc
        self._skia = skia
        self._typefaces = {}
        self._pattern_images = {}
        self._picture_images = {}

    def render(self, pages, file_path):
        if not pages:
            raise ValueError("PDF export has no printable worksheets")

        stream = self._skia.FILEWStream(str(file_path))
        document = self._skia.PDF.MakeDocument(stream)
        if document is None:
            raise RuntimeError(f"Skia could not create PDF document: {file_path}")
        try:
            for page in pages:
                canvas = document.beginPage(page.page_width, page.page_height)
                self._draw_page(canvas, page)
                document.endPage()
        finally:
            document.close()
            stream.flush()

    def _draw_page(self, canvas, page):
        for cell in page.cells:
            self._draw_fill(canvas, cell)
        for cell in page.cells:
            self._draw_data_bar(canvas, cell)
        for cell in page.cells:
            self._draw_icon_set(canvas, cell)
        if page.render_gridlines:
            for cell in page.cells:
                self._draw_gridline(canvas, cell)
        for cell in page.cells:
            self._draw_cell_borders(canvas, cell)
        for cell in page.cells:
            self._draw_text(canvas, page, cell)
        if page.sparklines or page.pictures or page.shapes or page.charts:
            content_rect = self._skia.Rect.MakeXYWH(
                page.content_left,
                page.content_top,
                page.content_width,
                page.content_height,
            )
            canvas.save()
            canvas.clipRect(content_rect)
            for chart in page.charts:
                self._draw_chart_glow(canvas, page, chart)
            for sparkline in page.sparklines:
                self._draw_sparkline(canvas, sparkline)
            for picture in page.pictures:
                self._draw_picture(canvas, picture)
            for shape in page.shapes:
                self._draw_shape(canvas, shape)
            for chart in page.charts:
                self._draw_chart(canvas, chart)
            canvas.restore()

    def _draw_sparkline(self, canvas, sparkline):
        if sparkline.kind == 'line':
            self._draw_line_sparkline(canvas, sparkline)
        elif sparkline.kind == 'column':
            self._draw_column_sparkline(canvas, sparkline)

    def _draw_line_sparkline(self, canvas, sparkline):
        points = self._sparkline_line_points(sparkline)
        paint = self._paint(
            'FF' + sparkline.color_series,
            stroke_width=sparkline.line_width,
        )
        paint.setStrokeCap(self._skia.Paint.kRound_Cap)
        path = None
        for point in points:
            if point is None:
                if path is not None:
                    canvas.drawPath(path, paint)
                    path = None
                continue
            if path is None:
                path = self._skia.Path()
                path.moveTo(*point)
            else:
                path.lineTo(*point)
        if path is not None:
            canvas.drawPath(path, paint)

    def _draw_column_sparkline(self, canvas, sparkline):
        rectangles = self._sparkline_column_rectangles(sparkline)
        for rectangle in rectangles:
            if rectangle is None:
                continue
            x, y, width, height, negative = rectangle
            color = (
                sparkline.color_negative
                if negative else sparkline.color_series
            )
            if height <= 1e-9:
                canvas.drawLine(
                    x, y, x + width, y,
                    self._paint('55' + color, stroke_width=0.3),
                )
            else:
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(x, y, width, height),
                    self._paint('FF' + color, fill=True),
                )

    @staticmethod
    def _sparkline_line_points(sparkline):
        count = len(sparkline.values)
        if count == 0:
            return ()
        numeric = tuple(
            float(value) for value in sparkline.values if value is not None
        )
        if not numeric:
            return tuple(None for _ in sparkline.values)
        minimum = min(numeric)
        maximum = max(numeric)
        span = maximum - minimum
        left = sparkline.x + 7.0
        right = sparkline.x + sparkline.width - 7.0
        top = sparkline.y + 2.25
        bottom = sparkline.y + sparkline.height - 2.25
        points = []
        for index, value in enumerate(sparkline.values):
            if value is None:
                points.append(None)
                continue
            ratio = (
                (float(value) - minimum) / span
                if span > 0.0 else 0.5
            )
            x = (
                (left + right) / 2.0
                if count == 1 else
                left + (right - left) * index / (count - 1)
            )
            points.append((x, bottom - (bottom - top) * ratio))
        if sparkline.empty_cells == 'connected':
            return tuple(point for point in points if point is not None)
        return tuple(points)

    @staticmethod
    def _sparkline_column_rectangles(sparkline):
        count = len(sparkline.values)
        if count == 0:
            return ()
        numeric = tuple(
            float(value) for value in sparkline.values if value is not None
        )
        if not numeric:
            return tuple(None for _ in sparkline.values)
        minimum = min(numeric)
        maximum = max(numeric)
        span = maximum - minimum
        baseline_value = (
            0.0
            if minimum < 0.0 < maximum
            else minimum if minimum >= 0.0 else maximum
        )
        left = sparkline.x + 3.4
        available_width = max(0.0, sparkline.width - 6.8)
        gap_ratio = 0.35
        bar_width = available_width / max(
            1.0, count + gap_ratio * max(0, count - 1)
        )
        step = bar_width * (1.0 + gap_ratio)
        top = sparkline.y + 2.25
        bottom = sparkline.y + sparkline.height - 2.5

        def value_y(value):
            ratio = (
                (float(value) - minimum) / span
                if span > 0.0 else 0.5
            )
            return bottom - (bottom - top) * ratio

        baseline_y = value_y(baseline_value)
        rectangles = []
        for index, value in enumerate(sparkline.values):
            if value is None:
                rectangles.append(None)
                continue
            point_y = value_y(value)
            y = min(point_y, baseline_y)
            rectangles.append((
                left + step * index,
                y,
                bar_width,
                abs(point_y - baseline_y),
                float(value) < 0.0,
            ))
        return tuple(rectangles)

    def _draw_chart_glow(self, canvas, page, chart):
        if (
            not chart.chart_glow_color
            or chart.chart_glow_opacity <= 0.0
            or chart.chart_glow_radius <= 0.0
        ):
            return
        alpha = round(255.0 * min(1.0, chart.chart_glow_opacity * 2.0))
        paint = self._paint(
            f'{alpha:02X}' + chart.chart_glow_color,
            fill=True,
        )
        paint.setMaskFilter(self._skia.MaskFilter.MakeBlur(
            self._skia.BlurStyle.kNormal_BlurStyle,
            chart.chart_glow_radius / 3.0,
        ))
        rect = self._skia.Rect.MakeXYWH(
            chart.x, chart.y, chart.width, chart.height
        )
        canvas.save()
        if chart.x + chart.width <= page.content_left:
            canvas.clipRect(self._skia.Rect.MakeLTRB(
                chart.x + chart.width,
                0.0,
                page.page_width,
                page.page_height,
            ))
        canvas.clipRect(
            rect, self._skia.ClipOp.kDifference, doAntiAlias=True
        )
        canvas.drawRect(
            rect, paint,
        )
        canvas.restore()

    def _draw_picture(self, canvas, picture):
        cached = self._picture_images.get(picture.image_bytes)
        if cached is None:
            data = self._skia.Data.MakeWithoutCopy(picture.image_bytes)
            image = self._skia.Image.MakeFromEncoded(data)
            if image is None:
                raise RuntimeError(
                    f"Skia could not decode worksheet picture: {picture.name}"
                )
            cached = (image, data)
            self._picture_images[picture.image_bytes] = cached
        image = cached[0]
        rect = self._skia.Rect.MakeXYWH(
            picture.x, picture.y, picture.width, picture.height
        )
        canvas.drawImageRect(image, rect)

    def _draw_chart_top_bevel(self, canvas, chart, outer):
        width = min(
            max(0.0, getattr(chart, 'chart_bevel_top_width', 0.0)),
            chart.width / 2.0,
            chart.height / 2.0,
        )
        if width <= 0.0:
            return

        left = outer.left()
        top = outer.top()
        right = outer.right()
        bottom = outer.bottom()

        def gradient_paint(start, end, colors, positions):
            paint = self._skia.Paint(AntiAlias=True)
            paint.setShader(self._skia.GradientShader.MakeLinear(
                (self._skia.Point(*start), self._skia.Point(*end)),
                colors,
                positions,
                self._skia.TileMode.kClamp,
            ))
            return paint

        canvas.save()
        canvas.clipRect(
            outer, self._skia.ClipOp.kDifference, doAntiAlias=True
        )
        shadow = self._paint('68000000', fill=True)
        shadow.setMaskFilter(self._skia.MaskFilter.MakeBlur(
            self._skia.BlurStyle.kNormal_BlurStyle, 0.7,
        ))
        canvas.drawRect(
            self._skia.Rect.MakeXYWH(
                right - 0.2, top + width, 0.9, chart.height - width,
            ),
            shadow,
        )
        canvas.restore()

        right_bevel = self._polygon_path((
            (right - width, top + width),
            (right, top),
            (right, bottom),
            (right - width, bottom - width),
        ))
        canvas.drawPath(
            right_bevel,
            gradient_paint(
                (right - width, top), (right, top),
                (
                    self._skia.ColorSetARGB(0, 255, 255, 255),
                    self._skia.ColorSetARGB(255, 89, 89, 89),
                ),
                (0.0, 1.0),
            ),
        )

        bottom_bevel = self._polygon_path((
            (left, bottom),
            (right, bottom),
            (right - width, bottom - width),
            (left + width, bottom - width),
        ))
        canvas.drawPath(
            bottom_bevel,
            gradient_paint(
                (left, bottom - width), (left, bottom),
                (
                    self._skia.ColorSetARGB(0, 255, 255, 255),
                    self._skia.ColorSetARGB(255, 166, 166, 166),
                    self._skia.ColorSetARGB(255, 166, 166, 166),
                    self._skia.ColorSetARGB(255, 217, 217, 217),
                    self._skia.ColorSetARGB(255, 217, 217, 217),
                ),
                (0.0, 0.55, 0.72, 0.84, 1.0),
            ),
        )

    def _draw_chart_soft_edge_background(self, canvas, chart, outer):
        radius = min(
            max(0.0, chart.chart_soft_edge_radius),
            chart.width / 2.0,
            chart.height / 2.0,
        )
        underlay_color = chart.chart_glow_color or 'FFFFFF'
        corner_radius = radius * 0.65
        canvas.drawRoundRect(
            outer, corner_radius, corner_radius,
            self._paint('FF' + underlay_color, fill=True),
        )

        inset = radius * 0.45
        inner = self._skia.Rect.MakeLTRB(
            outer.left() + inset,
            outer.top() + inset,
            outer.right() - inset,
            outer.bottom() - inset,
        )
        background = self._chart_background_paint(chart)
        background.setMaskFilter(self._skia.MaskFilter.MakeBlur(
            self._skia.BlurStyle.kNormal_BlurStyle,
            max(0.1, radius / 3.0),
        ))
        canvas.save()
        canvas.clipRect(outer)
        canvas.drawRoundRect(
            inner, corner_radius, corner_radius, background
        )
        canvas.restore()

    def _draw_chart(self, canvas, chart):
        outer = self._skia.Rect.MakeXYWH(
            chart.x, chart.y, chart.width, chart.height
        )
        compact_triangle_radar = (
            chart.chart_kind == 'radar'
            and chart.radar_style in ('standard', 'filled')
            and len(chart.categories) == 3
        )
        if chart.chart_soft_edge_radius > 0.0:
            self._draw_chart_soft_edge_background(canvas, chart, outer)
            if chart.chart_border_color and chart.chart_border_width > 0:
                canvas.drawRect(
                    outer,
                    self._paint(
                        'FF' + chart.chart_border_color,
                        stroke_width=chart.chart_border_width,
                    ),
                )
        elif compact_triangle_radar or chart.rounded_corners:
            canvas.drawRoundRect(
                outer, 10.0, 10.0,
                self._chart_background_paint(chart),
            )
            canvas.drawRoundRect(
                outer, 10.0, 10.0,
                self._paint(
                    'FF' + (
                        '898989' if (
                            chart.surface_axes_hidden
                            or chart.scatter_vary_colors
                        )
                        else chart.chart_border_color or '898989'
                    ),
                    stroke_width=(
                        1.0 if chart.scatter_vary_colors
                        else chart.chart_border_width or 1.0
                    ),
                ),
            )
        else:
            canvas.drawRect(outer, self._chart_background_paint(chart))
            if chart.chart_border_color and chart.chart_border_width > 0:
                canvas.drawRect(
                    outer,
                    self._paint(
                        'FF' + chart.chart_border_color,
                        stroke_width=chart.chart_border_width,
                    ),
                )
            self._draw_chart_top_bevel(canvas, chart, outer)

        axis_font = self._font(chart.axis_font_resolution, 9.0)
        axis_title_font = self._font(chart.axis_font_resolution, 10.0)
        title_font = self._font(
            chart.title_font_resolution,
            getattr(chart.title_font_resolution, 'size_points', 14.0),
        )
        value_text_paint = self._chart_text_paint(
            chart.value_axis_text_color, chart.value_axis_fill_image
        )
        data_label_paint = self._chart_text_paint(
            chart.data_label_color, None
        )
        category_text_paint = self._chart_text_paint(
            chart.category_axis_text_color, chart.category_axis_fill_image
        )
        legend_text_paint = self._chart_text_paint(
            chart.legend_text_color,
            chart.legend_fill_image,
            chart.legend_gradient_stops,
            chart.legend_gradient_angle,
            (
                chart.x, chart.y + chart.height - 30.0,
                chart.width, 30.0,
            ),
        )
        title_paint = self._chart_text_paint(
            chart.title_color,
            chart.title_fill_image,
            chart.title_gradient_stops,
            chart.title_gradient_angle,
            (chart.x, chart.y, chart.width, 30.0),
        )
        grid_paint = self._paint(
            'FF' + (
                '000000' if chart.chart_kind == 'stock'
                and chart.stock_axes_hidden
                else chart.value_major_gridline_color
            ),
            stroke_width=0.75,
        )
        axis_metrics = axis_font.getMetrics()

        if chart.is_3d:
            if chart.chart_kind == 'column':
                self._draw_3d_column_chart(
                    canvas, chart, axis_font, title_font,
                    value_text_paint, category_text_paint,
                    legend_text_paint, title_paint, grid_paint,
                )
            elif chart.chart_kind == 'bar':
                self._draw_3d_bar_chart(
                    canvas, chart, axis_font, axis_title_font, title_font,
                    value_text_paint, category_text_paint,
                    legend_text_paint, title_paint, grid_paint,
                )
            elif chart.chart_kind == 'pie':
                self._draw_3d_pie_chart(
                    canvas, chart, axis_font, title_font,
                    legend_text_paint, title_paint,
                )
            elif chart.chart_kind == 'area':
                self._draw_3d_area_chart(
                    canvas, chart, axis_font, title_font,
                    value_text_paint, category_text_paint,
                    legend_text_paint, title_paint, grid_paint,
                )
            elif chart.chart_kind == 'surface':
                self._draw_3d_surface_chart(
                    canvas, chart, axis_font, title_font,
                    value_text_paint, category_text_paint,
                    legend_text_paint, title_paint,
                )
            else:
                self._draw_3d_line_chart(
                    canvas, chart, axis_font, title_font,
                    value_text_paint, category_text_paint,
                    legend_text_paint, title_paint, grid_paint,
                )
            return

        if chart.chart_kind == 'bar':
            self._draw_2d_bar_chart(
                canvas, chart, axis_font, axis_title_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return
        if chart.chart_kind == 'pie':
            self._draw_2d_pie_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        if chart.chart_kind == 'radar':
            self._draw_2d_radar_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint,
            )
            return
        if chart.chart_kind == 'box_whisker':
            self._draw_2d_box_whisker_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return
        if chart.chart_kind == 'funnel':
            self._draw_2d_funnel_chart(
                canvas, chart, axis_font, title_font,
                category_text_paint, legend_text_paint, title_paint,
                grid_paint,
            )
            return
        if chart.chart_kind == 'histogram':
            self._draw_2d_histogram_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return
        if chart.chart_kind == 'map':
            self._draw_2d_map_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        if chart.chart_kind == 'sunburst':
            self._draw_2d_sunburst_chart(
                canvas, chart, axis_font, title_font, title_paint
            )
            return
        if chart.chart_kind == 'treemap':
            self._draw_2d_treemap_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        if chart.chart_kind == 'waterfall':
            self._draw_2d_waterfall_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return
        if chart.chart_kind in ('scatter', 'bubble'):
            self._draw_2d_scatter_chart(
                canvas, chart, axis_font, axis_title_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return
        if chart.chart_kind == 'surface':
            self._draw_2d_surface_chart(
                canvas, chart, axis_font, title_font,
                category_text_paint, legend_text_paint, title_paint,
            )
            return

        if chart.chart_kind == 'stock':
            plot_fill = (
                '000000'
                if chart.stock_has_volume and chart.stock_axes_hidden
                else 'FFFFFF'
            )
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    chart.plot_x, chart.plot_y,
                    chart.plot_width, chart.plot_height,
                ),
                self._paint('FF' + plot_fill, fill=True),
            )

        axis_ticks = self._chart_value_axis_ticks(chart)
        tick_count = max(1, len(axis_ticks) - 1)
        for value, ratio in axis_ticks:
            y = chart.plot_y + chart.plot_height * (1.0 - ratio)
            canvas.drawLine(
                chart.plot_x, y,
                chart.plot_x + chart.plot_width, y,
                grid_paint,
            )
            if not chart.stock_axes_hidden:
                label = self._chart_value_label(
                    value, chart.axis_number_format
                )
                label_width = axis_font.measureText(label)
                label_gap = 8.8 if '%' in chart.axis_number_format else 10.0
                baseline = (
                    y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                    - 0.6
                )
                canvas.drawString(
                    label, chart.plot_x - label_gap - label_width, baseline,
                    axis_font, value_text_paint,
                )

        if chart.chart_kind == 'stock' and chart.stock_has_volume:
            self._draw_stock_price_axis(
                canvas, chart, axis_font, axis_metrics, value_text_paint
            )

        if chart.show_minor_horizontal_gridlines:
            minor_paint = self._paint(
                'FF' + chart.minor_gridline_color, stroke_width=0.75
            )
            for major_index in range(tick_count):
                for minor_index in range(1, 5):
                    ratio = (
                        major_index + minor_index / 5.0
                    ) / max(1, tick_count)
                    y = chart.plot_y + chart.plot_height * (1.0 - ratio)
                    canvas.drawLine(
                        chart.plot_x, y,
                        chart.plot_x + chart.plot_width, y,
                        minor_paint,
                    )

        if chart.value_axis_color and not chart.stock_axes_hidden:
            value_axis_paint = self._paint(
                'FF' + chart.value_axis_color, stroke_width=0.75
            )
            canvas.drawLine(
                chart.plot_x, chart.plot_y,
                chart.plot_x, chart.plot_y + chart.plot_height,
                value_axis_paint,
            )
            for x1, y1, x2, y2 in self._chart_value_axis_tick_segments(
                chart, horizontal=False
            ):
                canvas.drawLine(x1, y1, x2, y2, value_axis_paint)

        if chart.show_vertical_gridlines and chart.axis_categories:
            category_count = len(chart.axis_categories)
            for index in range(category_count + 1):
                x = chart.plot_x + chart.plot_width * index / category_count
                canvas.drawLine(
                    x, chart.plot_y,
                    x, chart.plot_y + chart.plot_height + 2.8,
                    grid_paint,
                )

        if (
            chart.plot_border_color
            or chart.chart_kind == 'stock' and chart.stock_axes_hidden
        ):
            plot_rect = self._skia.Rect.MakeXYWH(
                chart.plot_x, chart.plot_y,
                chart.plot_width, chart.plot_height,
            )
            canvas.drawRect(
                plot_rect,
                self._paint(
                    'FF' + (chart.plot_border_color or '000000'),
                    stroke_width=0.75,
                ),
            )

        if chart.chart_kind in ('column', 'area') and chart.axis_categories:
            category_count = len(chart.axis_categories)
            canvas.drawLine(
                chart.plot_x, chart.plot_y + chart.plot_height,
                chart.plot_x + chart.plot_width,
                chart.plot_y + chart.plot_height,
                grid_paint,
            )
            tick_positions = (
                tuple(position for _, position in chart.axis_categories)
                if chart.chart_kind == 'area'
                else tuple(
                    index / category_count
                    for index in range(category_count + 1)
                )
            )
            for position in tick_positions:
                x = chart.plot_x + chart.plot_width * position
                canvas.drawLine(
                    x, chart.plot_y + chart.plot_height,
                    x, chart.plot_y + chart.plot_height + 2.8,
                    grid_paint,
                )

        if chart.axis_categories and not chart.stock_axes_hidden:
            category_y = chart.plot_y + chart.plot_height + 5.5 - axis_metrics.fAscent
            for label, position in chart.axis_categories:
                x = chart.plot_x + chart.plot_width * position
                if (
                    not chart.show_vertical_gridlines
                    and chart.chart_kind not in ('column', 'area')
                ):
                    canvas.drawLine(
                        x, chart.plot_y + chart.plot_height,
                        x, chart.plot_y + chart.plot_height + 2.8,
                        grid_paint,
                    )
                label_width = axis_font.measureText(label)
                if chart.category_rotation:
                    canvas.save()
                    canvas.translate(x, chart.plot_y + chart.plot_height + 8.0)
                    canvas.rotate(chart.category_rotation)
                    canvas.drawString(
                        label, -label_width, 0.0,
                        axis_font, category_text_paint
                    )
                    canvas.restore()
                else:
                    canvas.drawString(
                        label, x - label_width / 2.0, category_y,
                        axis_font, category_text_paint,
                    )

        if chart.chart_kind == 'column':
            self._draw_column_chart_series(
                canvas, chart, axis_font, data_label_paint
            )
        elif chart.chart_kind == 'area' and chart.category_positions:
            point_count = len(chart.category_positions)
            area_value_height = chart.plot_height
            if chart.grouping == 'stacked' and len(chart.series) > 3:
                # Excel's compact multi-series stack extends the data bands
                # slightly beyond the grid's nominal value-axis height.
                area_value_height += 0.2
            area_baseline = chart.plot_y + chart.plot_height
            cumulative_values = [0.0] * point_count
            category_totals = [
                sum(
                    float(series.values[index])
                    for series in chart.series
                    if index < len(series.values)
                )
                for index in range(point_count)
            ]
            for series in chart.series:
                values = [
                    float(series.values[index])
                    if index < len(series.values) else 0.0
                    for index in range(point_count)
                ]
                if not values:
                    continue
                if chart.grouping in ('stacked', 'percentStacked'):
                    lower_values = list(cumulative_values)
                    upper_values = [
                        lower_values[index] + values[index]
                        for index in range(point_count)
                    ]
                    cumulative_values = upper_values
                    if chart.grouping == 'percentStacked':
                        lower_values = [
                            value / category_totals[index]
                            if category_totals[index] else 0.0
                            for index, value in enumerate(lower_values)
                        ]
                        upper_values = [
                            value / category_totals[index]
                            if category_totals[index] else 0.0
                            for index, value in enumerate(upper_values)
                        ]
                else:
                    lower_values = [0.0] * point_count
                    upper_values = values

                upper_points = []
                lower_points = []
                for position, lower_value, upper_value in zip(
                    chart.category_positions, lower_values, upper_values
                ):
                    x = chart.plot_x + chart.plot_width * position
                    upper_ratio = self._chart_value_ratio(
                        chart, upper_value
                    )
                    lower_ratio = self._chart_value_ratio(
                        chart, lower_value
                    )
                    upper_points.append((
                        x,
                        area_baseline - area_value_height * upper_ratio,
                    ))
                    lower_points.append((
                        x,
                        area_baseline - area_value_height * lower_ratio,
                    ))
                path = self._skia.Path()
                path.moveTo(*upper_points[0])
                for point in upper_points[1:]:
                    path.lineTo(*point)
                for point in reversed(lower_points):
                    path.lineTo(*point)
                path.close()
                canvas.drawPath(
                    path, self._paint('FF' + series.color, fill=True)
                )
        elif chart.chart_kind == 'stock' and chart.category_positions:
            if chart.stock_has_high_low_lines:
                self._draw_stock_chart_series(canvas, chart)
            else:
                self._draw_stock_combo_series(canvas, chart)
        elif chart.category_positions:
            stacked_values = [0.0] * len(chart.category_positions)
            category_totals = [
                sum(
                    float(series.values[index])
                    for series in chart.series
                    if index < len(series.values)
                )
                for index in range(len(chart.category_positions))
            ]
            for series_index, series in enumerate(chart.series):
                path = self._skia.Path()
                values = series.values
                if chart.grouping in ('stacked', 'percentStacked'):
                    cumulative_values = tuple(
                        stacked_values[index] + float(value)
                        for index, value in enumerate(series.values)
                    )
                    stacked_values[:len(cumulative_values)] = cumulative_values
                    if chart.grouping == 'percentStacked':
                        values = tuple(
                            value / category_totals[index]
                            if category_totals[index] else 0.0
                            for index, value in enumerate(cumulative_values)
                        )
                    else:
                        values = cumulative_values
                points = sorted(zip(chart.category_positions, values))
                rendered_points = []
                for index, (position, value) in enumerate(points):
                    x = chart.plot_x + chart.plot_width * position
                    ratio = self._chart_value_ratio(chart, value)
                    y = chart.plot_y + chart.plot_height * (1.0 - ratio)
                    rendered_points.append((x, y))
                    if index == 0:
                        path.moveTo(x, y)
                    else:
                        path.lineTo(x, y)
                series_paint = self._paint(
                    'FF' + series.color,
                    stroke_width=self._chart_series_line_width(
                        chart, series_index
                    ),
                )
                if series.line_cap == 'round':
                    series_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
                canvas.drawPath(path, series_paint)
                for x, y in rendered_points:
                    self._draw_chart_marker(canvas, series, x, y)

        self._draw_chart_axis_titles(
            canvas, chart, axis_title_font,
            value_text_paint, category_text_paint,
        )
        self._draw_chart_data_table(
            canvas, chart, axis_font, legend_text_paint, grid_paint
        )

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        if chart.title_layout_x is None:
            title_x = chart.x + (chart.width - title_width) / 2.0
        else:
            title_x = chart.x + chart.width * chart.title_layout_x + 3.0
        if chart.title_layout_y is None:
            title_box_y = chart.y + 5.99
            title_y = chart.y + 7.5 - title_metrics.fAscent
        else:
            title_box_y = chart.y + chart.height * chart.title_layout_y
            title_y = title_box_y + 1.54 - title_metrics.fAscent
        if chart.title_border_color:
            title_box = self._skia.Rect.MakeXYWH(
                title_x - 3.0, title_box_y,
                title_width + 6.0, 19.9,
            )
            canvas.drawRect(
                title_box,
                self._paint(
                    'FF' + chart.title_border_color, stroke_width=0.75
                ),
            )
        canvas.drawString(
            chart.title, title_x, title_y, title_font, title_paint
        )

        if chart.show_legend and chart.series:
            if chart.chart_kind == 'stock':
                self._draw_stock_chart_legend(
                    canvas, chart, axis_font, axis_metrics,
                    legend_text_paint,
                )
                return
            if chart.chart_kind == 'area':
                legend_font = self._font(chart.axis_font_resolution, 9.0)
                legend_font.setScaleX(0.96)
                compact_legend = len(chart.series) <= 3
                item_widths = [
                    (17.0 if compact_legend else 16.0)
                    + legend_font.measureText(series.name)
                    for series in chart.series
                ]
                legend_x = (
                    chart.x
                    + (chart.width - sum(item_widths)) / 2.0
                    + (6.02 if compact_legend else 5.9)
                )
                legend_square_y = chart.y + chart.height - 16.91
                legend_text_y = chart.y + chart.height - 20.43
                legend_metrics = legend_font.getMetrics()
                legend_baseline = legend_text_y - legend_metrics.fAscent
                for series, item_width in zip(chart.series, item_widths):
                    canvas.drawRect(
                        self._skia.Rect.MakeXYWH(
                            legend_x, legend_square_y, 4.94, 4.94
                        ),
                        self._paint('FF' + series.color, fill=True),
                    )
                    canvas.drawString(
                        series.name, legend_x + 7.07, legend_baseline,
                        legend_font, legend_text_paint,
                    )
                    legend_x += item_width
                return
            if chart.chart_kind == 'column':
                item_gap = 14.5 if chart.legend_layout_x is not None else 7.0
                text_offset = (
                    7.5 if chart.legend_layout_x is not None else 8.0
                )
                item_widths = [
                    9.0 + axis_font.measureText(series.name) + item_gap
                    for series in chart.series
                ]
                if chart.legend_layout_x is None:
                    legend_x = (
                        chart.x
                        + (chart.width - sum(item_widths)) / 2.0
                        + 7.0
                    )
                else:
                    legend_x = (
                        chart.x + chart.width * chart.legend_layout_x + 18.5
                    )
                if chart.legend_layout_y is None:
                    legend_y = chart.y + chart.height - (
                        20.5 if chart.show_data_table else 16.9
                    )
                else:
                    legend_y = (
                        chart.y + chart.height * chart.legend_layout_y + 4.2
                    )
                legend_baseline = legend_y - axis_metrics.fAscent
                for series, item_width in zip(chart.series, item_widths):
                    canvas.drawRect(
                        self._skia.Rect.MakeXYWH(
                            legend_x, legend_y + 1.5, 5.0, 5.0
                        ),
                        self._paint('FF' + series.color, fill=True),
                    )
                    canvas.drawString(
                        series.name, legend_x + text_offset, legend_baseline,
                        axis_font, legend_text_paint,
                    )
                    legend_x += item_width
                return
            repeated_stacked_names = (
                chart.grouping == 'stacked'
                and len({series.name for series in chart.series}) == 1
            )
            if repeated_stacked_names:
                item_gap = 8.0
                legend_offset = 7.5
            else:
                item_gap = 7.0 if len(chart.series) > 3 else 9.0
                legend_offset = 7.0
            item_widths = [
                24.0 + axis_font.measureText(series.name) + item_gap
                for series in chart.series
            ]
            legend_x = (
                chart.x + (chart.width - sum(item_widths)) / 2.0
                + legend_offset
            )
            legend_y = chart.y + chart.height - 20.6
            legend_baseline = legend_y - axis_metrics.fAscent
            for series, item_width in zip(chart.series, item_widths):
                legend_paint = self._paint(
                    'FF' + series.color, stroke_width=series.line_width
                )
                if series.line_cap == 'round':
                    legend_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
                canvas.drawLine(
                    legend_x, legend_y + 6.16,
                    legend_x + 19.2, legend_y + 6.16,
                    legend_paint,
                )
                self._draw_chart_marker(
                    canvas, series,
                    legend_x + 9.6, legend_y + 6.16,
                )
                canvas.drawString(
                    series.name, legend_x + 22.0, legend_baseline,
                    axis_font, legend_text_paint,
                )
                legend_x += item_width
