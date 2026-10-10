"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRendererPieBarMixin:
    def _draw_3d_pie_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        slices = self._chart_pie_slices(chart)
        inset_x = min(7.9, chart.plot_width * 0.05)
        base_ellipse_width = max(0.0, chart.plot_width - 2.0 * inset_x)
        yaw_scale = 1.0 - 0.065 * abs(
            math.sin(math.radians(chart.rotation_y))
        )
        ellipse_width = base_ellipse_width * yaw_scale
        tilt = self._pie_3d_tilt(chart.rotation_x, chart.rotation_y)
        ellipse_height = min(chart.plot_height, ellipse_width * tilt)
        ellipse_y = chart.plot_y + min(1.1, chart.plot_height * 0.01)
        depth = max(0.0, chart.plot_height - ellipse_height - 1.1)
        oval = self._skia.Rect.MakeXYWH(
            chart.plot_x + (chart.plot_width - ellipse_width) / 2.0,
            ellipse_y,
            ellipse_width, ellipse_height,
        )
        lower_oval = self._skia.Rect.MakeXYWH(
            oval.left(), oval.top() + depth,
            oval.width(), oval.height(),
        )

        projected_slices = tuple(
            (
                index,
                self._pie_3d_angle(
                    start_angle, chart.perspective, chart.rotation_y
                ),
                self._pie_3d_angle(
                    start_angle + sweep_angle,
                    chart.perspective,
                    chart.rotation_y,
                ) - self._pie_3d_angle(
                    start_angle, chart.perspective, chart.rotation_y
                ),
                fill,
                border,
                width,
            )
            for index, start_angle, sweep_angle, fill, border, width in slices
        )

        for _, start_angle, sweep_angle, fill, _, _ in projected_slices:
            for visible_start, visible_sweep in self._pie_visible_arc_segments(
                start_angle, sweep_angle
            ):
                start_point = self._ellipse_point(oval, visible_start)
                end_point = self._ellipse_point(
                    oval, visible_start + visible_sweep
                )
                path = self._skia.Path()
                path.moveTo(*start_point)
                path.arcTo(oval, visible_start, visible_sweep, False)
                path.lineTo(end_point[0], end_point[1] + depth)
                path.arcTo(
                    lower_oval,
                    visible_start + visible_sweep,
                    -visible_sweep,
                    False,
                )
                path.close()
                canvas.drawPath(
                    path,
                    self._chart_pie_side_paint(
                        fill, oval.left(), oval.right()
                    ),
                )

        center_x = oval.centerX()
        center_y = oval.centerY()
        for _, start_angle, sweep_angle, fill, border, width in projected_slices:
            path = self._skia.Path()
            path.moveTo(center_x, center_y)
            path.arcTo(oval, start_angle, sweep_angle, False)
            path.close()
            top_fill = self._scale_chart_color(fill, 0.92)
            canvas.drawPath(path, self._paint('FF' + top_fill, fill=True))
            if border and width > 0.0:
                canvas.drawPath(
                    path,
                    self._paint('FF' + border, stroke_width=width),
                )

        self._draw_pie_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint, slices,
        )

    @staticmethod
    def _ellipse_point(oval, angle):
        radians = math.radians(angle)
        return (
            oval.centerX() + oval.width() * 0.5 * math.cos(radians),
            oval.centerY() + oval.height() * 0.5 * math.sin(radians),
        )

    @staticmethod
    def _pie_3d_angle(angle, perspective, rotation_y=0.0):
        oriented = angle + rotation_y
        return (
            oriented
            - 0.9 * perspective * math.cos(math.radians(oriented))
        )

    @staticmethod
    def _pie_3d_tilt(rotation_x, rotation_y=0.0):
        yaw = abs(math.sin(math.radians(rotation_y)))
        tilt_scale = 0.74 + 0.27 * yaw
        return max(
            0.12,
            min(
                0.65,
                math.sin(math.radians(abs(rotation_x))) * tilt_scale,
            ),
        )

    @staticmethod
    def _pie_visible_arc_segments(start_angle, sweep_angle):
        end_angle = start_angle + sweep_angle
        segments = []
        first_turn = math.floor(start_angle / 360.0) - 1
        last_turn = math.ceil(end_angle / 360.0) + 1
        for turn in range(first_turn, last_turn + 1):
            visible_start = 360.0 * turn
            visible_end = visible_start + 180.0
            segment_start = max(start_angle, visible_start)
            segment_end = min(end_angle, visible_end)
            if segment_end > segment_start:
                segments.append((segment_start, segment_end - segment_start))
        return tuple(segments)

    def _chart_pie_side_paint(self, color, left, right):
        colors = []
        for factor in (0.92, 0.68, 0.40):
            shaded = self._scale_chart_color(color, factor)
            alpha, red, green, blue = self._argb_channels('FF' + shaded)
            colors.append(self._skia.ColorSetARGB(alpha, red, green, blue))
        paint = self._skia.Paint(AntiAlias=True)
        paint.setShader(self._skia.GradientShader.MakeLinear(
            (
                self._skia.Point(left, 0.0),
                self._skia.Point(max(left + 0.01, right), 0.0),
            ),
            colors,
            (0.0, 0.55, 1.0),
            self._skia.TileMode.kClamp,
        ))
        return paint

    def _draw_pie_title_and_legend(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint, slices,
    ):
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
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    title_x - 3.0, title_box_y,
                    title_width + 6.0, 19.9,
                ),
                self._paint(
                    'FF' + chart.title_border_color, stroke_width=0.75
                ),
            )
        canvas.drawString(
            chart.title,
            title_x,
            title_y,
            title_font,
            title_paint,
        )

        if not chart.show_legend or not slices:
            return
        labels = tuple(str(value) for value in chart.categories)
        compact_labels = all(
            len(label) <= 2 and label.isdigit() for label in labels
        )
        item_padding = 12.36 if compact_labels else 15.03
        item_widths = tuple(
            item_padding + axis_font.measureText(label) for label in labels
        )
        legend_x = (
            chart.x + (chart.width - sum(item_widths)) / 2.0
            + (4.26 if compact_labels else 5.97)
        )
        legend_y = chart.y + chart.height - 16.91
        axis_metrics = axis_font.getMetrics()
        legend_baseline = (
            chart.y + chart.height - 20.43 - axis_metrics.fAscent
        )
        for label, item_width, point_style in zip(
            labels, item_widths, chart.pie_point_styles
        ):
            fill, border, width = point_style
            rect = self._skia.Rect.MakeXYWH(
                legend_x, legend_y, 4.94, 4.94
            )
            canvas.drawRect(rect, self._paint('FF' + fill, fill=True))
            if border and width > 0.0:
                canvas.drawRect(
                    rect,
                    self._paint('FF' + border, stroke_width=width),
                )
            canvas.drawString(
                label, legend_x + 7.1, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += item_width

    def _draw_2d_pie_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        if chart.pie_variant == 'doughnut':
            self._draw_doughnut_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        if chart.pie_variant == 'pie_of_pie':
            self._draw_pie_of_pie_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        if chart.pie_variant == 'bar_of_pie':
            self._draw_bar_of_pie_chart(
                canvas, chart, axis_font, title_font,
                legend_text_paint, title_paint,
            )
            return
        slices = self._chart_pie_slices(chart)
        oval = self._skia.Rect.MakeXYWH(
            chart.plot_x, chart.plot_y,
            chart.plot_width, chart.plot_height,
        )
        center_x = chart.plot_x + chart.plot_width / 2.0
        center_y = chart.plot_y + chart.plot_height / 2.0

        for _, start_angle, sweep_angle, fill, border, width in slices:
            path = self._skia.Path()
            path.moveTo(center_x, center_y)
            path.arcTo(oval, start_angle, sweep_angle, False)
            path.close()
            canvas.drawPath(path, self._paint('FF' + fill, fill=True))
            if border and width > 0.0:
                canvas.drawPath(
                    path,
                    self._paint('FF' + border, stroke_width=width),
                )

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font,
            title_paint,
        )

        if not chart.show_legend or not slices:
            return
        labels = tuple(str(value) for value in chart.categories)
        compact_labels = all(
            len(label) <= 2 and label.isdigit() for label in labels
        )
        item_padding = 12.36 if compact_labels else 15.03
        item_widths = tuple(
            item_padding + axis_font.measureText(label) for label in labels
        )
        legend_x = (
            chart.x + (chart.width - sum(item_widths)) / 2.0
            + (4.26 if compact_labels else 5.97)
        )
        legend_y = chart.y + chart.height - 16.91
        axis_metrics = axis_font.getMetrics()
        legend_baseline = (
            chart.y + chart.height - 20.43 - axis_metrics.fAscent
        )
        for label, item_width, point_style in zip(
            labels, item_widths, chart.pie_point_styles
        ):
            fill, border, width = point_style
            rect = self._skia.Rect.MakeXYWH(
                legend_x, legend_y, 4.94, 4.94
            )
            canvas.drawRect(rect, self._paint('FF' + fill, fill=True))
            if border and width > 0.0:
                canvas.drawRect(
                    rect,
                    self._paint('FF' + border, stroke_width=width),
                )
            canvas.drawString(
                label, legend_x + 7.1, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += item_width

    def _draw_doughnut_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        rings = self._doughnut_ring_geometry(chart)
        for series_index, (outer_rect, inner_rect) in enumerate(rings):
            slices = self._chart_doughnut_slices(chart, series_index)
            if not slices:
                continue
            outer = self._skia.Rect.MakeXYWH(*outer_rect)
            inner = self._skia.Rect.MakeXYWH(*inner_rect)
            for _, start_angle, sweep_angle, fill, border, width in slices:
                path = self._skia.Path()
                path.arcTo(outer, start_angle, sweep_angle, True)
                path.arcTo(
                    inner, start_angle + sweep_angle,
                    -sweep_angle, False,
                )
                path.close()
                canvas.drawPath(path, self._paint('FF' + fill, fill=True))
                if border and width > 0.0:
                    canvas.drawPath(
                        path,
                        self._paint('FF' + border, stroke_width=width),
                    )

        self._draw_pie_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
            self._chart_doughnut_slices(chart, 0),
        )

    @staticmethod
    def _doughnut_ring_geometry(chart):
        series_count = len(chart.series)
        if series_count <= 0:
            return ()
        diameter = min(chart.plot_width, chart.plot_height)
        outer_radius = diameter / 2.0
        hole_radius = outer_radius * chart.pie_hole_size / 100.0
        ring_width = (outer_radius - hole_radius) / series_count
        center_x = chart.plot_x + chart.plot_width / 2.0
        center_y = chart.plot_y + chart.plot_height / 2.0
        rings = []
        for series_index in range(series_count):
            inner_radius = hole_radius + ring_width * series_index
            series_outer_radius = inner_radius + ring_width
            rings.append((
                (
                    center_x - series_outer_radius,
                    center_y - series_outer_radius,
                    series_outer_radius * 2.0,
                    series_outer_radius * 2.0,
                ),
                (
                    center_x - inner_radius,
                    center_y - inner_radius,
                    inner_radius * 2.0,
                    inner_radius * 2.0,
                ),
            ))
        return tuple(rings)

    @staticmethod
    def _chart_doughnut_slices(chart, series_index):
        if not 0 <= series_index < len(chart.series):
            return ()
        values = tuple(
            max(0.0, float(value))
            for value in chart.series[series_index].values
        )
        total = sum(values)
        if total <= 0.0:
            return ()
        styles = (
            chart.pie_series_point_styles[series_index]
            if series_index < len(chart.pie_series_point_styles)
            else chart.pie_point_styles
        )
        angle = -90.0 + float(chart.pie_first_slice_angle)
        slices = []
        for index, value in enumerate(values):
            sweep = 360.0 * value / total
            fill, border, width = (
                styles[index]
                if index < len(styles)
                else ('156082', None, 0.0)
            )
            slices.append((
                index, angle, sweep, fill, border, width,
            ))
            angle += sweep
        return tuple(slices)

    def _draw_pie_of_pie_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        primary_slices, secondary_slices = (
            self._chart_pie_of_pie_slices(chart)
        )
        if not primary_slices or not secondary_slices:
            return
        main_rect, secondary_rect = self._pie_of_pie_geometry(chart)
        main_oval = self._skia.Rect.MakeXYWH(*main_rect)
        secondary_oval = self._skia.Rect.MakeXYWH(*secondary_rect)

        if chart.plot_border_color:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    chart.plot_x, chart.plot_y,
                    chart.plot_width, chart.plot_height,
                ),
                self._paint(
                    'FF' + chart.plot_border_color, stroke_width=0.75
                ),
            )

        other_slice = primary_slices[-1]
        boundaries = sorted((
            self._ellipse_point(main_oval, other_slice[1]),
            self._ellipse_point(
                main_oval, other_slice[1] + other_slice[2]
            ),
        ), key=lambda point: point[1])
        connector_paint = self._paint(
            'FF' + chart.pie_connector_color,
            stroke_width=chart.pie_connector_width,
        )
        canvas.drawLine(
            *boundaries[0], secondary_oval.centerX(), secondary_oval.top(),
            connector_paint,
        )
        canvas.drawLine(
            *boundaries[1], secondary_oval.centerX(), secondary_oval.bottom(),
            connector_paint,
        )
        self._draw_pie_paths(canvas, main_oval, primary_slices)
        self._draw_pie_paths(canvas, secondary_oval, secondary_slices)
        self._draw_pie_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
            self._chart_pie_slices(chart),
        )

    def _draw_bar_of_pie_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        primary_slices, secondary_slices = (
            self._chart_pie_of_pie_slices(chart)
        )
        if not primary_slices or not secondary_slices:
            return
        main_rect, bar_rect = self._bar_of_pie_geometry(chart)
        main_oval = self._skia.Rect.MakeXYWH(*main_rect)
        bar_x, bar_y, bar_width, bar_height = bar_rect

        if chart.plot_border_color:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    chart.plot_x, chart.plot_y,
                    chart.plot_width, chart.plot_height,
                ),
                self._paint(
                    'FF' + chart.plot_border_color, stroke_width=0.75
                ),
            )
        other_slice = primary_slices[-1]
        boundaries = sorted((
            self._ellipse_point(main_oval, other_slice[1]),
            self._ellipse_point(
                main_oval, other_slice[1] + other_slice[2]
            ),
        ), key=lambda point: point[1])
        connector_paint = self._paint(
            'FF' + chart.pie_connector_color,
            stroke_width=chart.pie_connector_width,
        )
        canvas.drawLine(
            *boundaries[0], bar_x, bar_y, connector_paint
        )
        canvas.drawLine(
            *boundaries[1], bar_x, bar_y + bar_height, connector_paint
        )
        self._draw_pie_paths(canvas, main_oval, primary_slices)

        y = bar_y
        total_sweep = sum(item[2] for item in secondary_slices)
        for _, _, sweep, fill, border, width in secondary_slices:
            height = (
                bar_height * sweep / total_sweep
                if total_sweep > 0.0 else 0.0
            )
            rect = self._skia.Rect.MakeXYWH(
                bar_x, y, bar_width, height
            )
            canvas.drawRect(rect, self._paint('FF' + fill, fill=True))
            if border and width > 0.0:
                canvas.drawRect(
                    rect,
                    self._paint('FF' + border, stroke_width=width),
                )
            y += height
        self._draw_pie_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
            self._chart_pie_slices(chart),
        )

    def _draw_pie_paths(self, canvas, oval, slices):
        center_x = oval.centerX()
        center_y = oval.centerY()
        for _, start_angle, sweep_angle, fill, border, width in slices:
            path = self._skia.Path()
            path.moveTo(center_x, center_y)
            path.arcTo(oval, start_angle, sweep_angle, False)
            path.close()
            canvas.drawPath(path, self._paint('FF' + fill, fill=True))
            if border and width > 0.0:
                canvas.drawPath(
                    path,
                    self._paint('FF' + border, stroke_width=width),
                )

    @staticmethod
    def _pie_of_pie_geometry(chart):
        secondary_ratio = chart.pie_secondary_size / 100.0
        gap_ratio = chart.gap_width / 200.0
        main_size = min(
            chart.plot_height,
            chart.plot_width / max(
                1e-12, 1.0 + secondary_ratio + gap_ratio
            ),
        )
        secondary_size = main_size * secondary_ratio
        gap = main_size * gap_ratio
        used_width = main_size + gap + secondary_size
        main_x = chart.plot_x + (chart.plot_width - used_width) / 2.0
        main_y = chart.plot_y + (chart.plot_height - main_size) / 2.0
        secondary_x = main_x + main_size + gap
        secondary_y = (
            chart.plot_y + (chart.plot_height - secondary_size) / 2.0
        )
        return (
            (main_x, main_y, main_size, main_size),
            (secondary_x, secondary_y, secondary_size, secondary_size),
        )

    @staticmethod
    def _bar_of_pie_geometry(chart):
        secondary_ratio = chart.pie_secondary_size / 100.0
        bar_width_ratio = secondary_ratio / 2.0
        gap_ratio = bar_width_ratio * chart.gap_width / 150.0
        main_size = min(
            chart.plot_height,
            chart.plot_width / max(
                1e-12, 1.0 + gap_ratio + bar_width_ratio
            ),
        )
        bar_height = main_size * secondary_ratio
        bar_width = main_size * bar_width_ratio
        gap = main_size * gap_ratio
        used_width = main_size + gap + bar_width
        main_x = chart.plot_x + (chart.plot_width - used_width) / 2.0
        main_y = chart.plot_y + (chart.plot_height - main_size) / 2.0
        bar_x = main_x + main_size + gap
        bar_y = chart.plot_y + (chart.plot_height - bar_height) / 2.0
        return (
            (main_x, main_y, main_size, main_size),
            (bar_x, bar_y, bar_width, bar_height),
        )

    @staticmethod
    def _chart_pie_of_pie_slices(chart):
        if not chart.series:
            return (), ()
        values = tuple(
            max(0.0, float(value)) for value in chart.series[0].values
        )
        total = sum(values)
        secondary_indices = tuple(chart.pie_secondary_indices)
        if total <= 0.0 or not secondary_indices:
            return (), ()
        secondary_set = set(secondary_indices)
        secondary_total = sum(values[index] for index in secondary_indices)
        other_sweep = 360.0 * secondary_total / total
        angle = float(chart.pie_first_slice_angle) + other_sweep / 2.0
        primary_slices = []
        for index, value in enumerate(values):
            if index in secondary_set:
                continue
            sweep = 360.0 * value / total
            fill, border, width = chart.pie_point_styles[index]
            primary_slices.append((
                index, angle, sweep, fill, border, width,
            ))
            angle += sweep
        fill, border, width = chart.pie_other_style
        primary_slices.append((
            -1, angle, other_sweep, fill, border, width,
        ))

        secondary_slices = []
        angle = float(chart.pie_first_slice_angle) + other_sweep / 2.0
        for index in secondary_indices:
            sweep = 360.0 * values[index] / secondary_total
            fill, border, width = chart.pie_point_styles[index]
            secondary_slices.append((
                index, angle, sweep, fill, border, width,
            ))
            angle += sweep
        return tuple(primary_slices), tuple(secondary_slices)

    @staticmethod
    def _chart_pie_slices(chart):
        if not chart.series:
            return ()
        values = tuple(
            max(0.0, float(value)) for value in chart.series[0].values
        )
        total = sum(values)
        if total <= 0.0:
            return ()
        styles = chart.pie_point_styles
        angle = -90.0 + float(chart.pie_first_slice_angle)
        slices = []
        for index, value in enumerate(values):
            sweep = 360.0 * value / total
            fill, border, width = (
                styles[index]
                if index < len(styles)
                else ('156082', None, 0.0)
            )
            slices.append((
                index, angle, sweep, fill, border, width,
            ))
            angle += sweep
        return tuple(slices)

    def _draw_2d_bar_chart(
        self, canvas, chart, axis_font, axis_title_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        axis_ticks = self._chart_value_axis_ticks(chart)
        plot_bottom = chart.plot_y + chart.plot_height
        value_axis_at_top = chart.value_axis_position == 't'
        value_axis_y = chart.plot_y if value_axis_at_top else plot_bottom
        value_grid_paint = self._paint(
            'FF' + chart.value_major_gridline_color,
            stroke_width=0.75,
        )
        category_grid_paint = self._paint(
            'FF' + chart.category_major_gridline_color,
            stroke_width=0.75,
        )

        for index, (value, ratio) in enumerate(axis_ticks):
            x = chart.plot_x + chart.plot_width * ratio
            if index > 0:
                canvas.drawLine(
                    x, chart.plot_y, x, plot_bottom, value_grid_paint
                )
            display_value = value / max(
                1e-12, chart.value_axis_display_unit
            )
            label = self._chart_value_label(
                display_value, chart.axis_number_format
            )
            label_width = axis_font.measureText(label)
            if value_axis_at_top:
                baseline = chart.plot_y - 6.66 - axis_metrics.fDescent
                label_x = x - label_width / 2.0
            else:
                baseline = plot_bottom + 5.62 - axis_metrics.fAscent
                label_x = x - label_width / 2.0
                if chart.grouping != 'percentStacked':
                    label_x -= 0.8
            canvas.drawString(
                label, label_x, baseline,
                axis_font, value_text_paint,
            )

        rectangles = self._chart_bar_rectangles(chart)
        for series_index, _, x, y, width, height in rectangles:
            if width <= 0.0 or height <= 0.0:
                continue
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(x, y, width, height),
                self._paint(
                    'FF' + chart.series[series_index].color, fill=True
                ),
            )

        if chart.value_axis_color:
            value_axis_paint = self._paint(
                'FF' + chart.value_axis_color, stroke_width=0.75
            )
            canvas.drawLine(
                chart.plot_x, value_axis_y,
                chart.plot_x + chart.plot_width, value_axis_y,
                value_axis_paint,
            )
            for x1, y1, x2, y2 in self._chart_value_axis_tick_segments(
                chart, horizontal=True
            ):
                canvas.drawLine(x1, y1, x2, y2, value_axis_paint)

        category_count = max(
            len(chart.axis_categories), len(chart.category_positions)
        )
        canvas.drawLine(
            chart.plot_x, chart.plot_y,
            chart.plot_x, plot_bottom,
            grid_paint,
        )
        for index in range(category_count + 1):
            y = chart.plot_y + chart.plot_height * index / max(
                1, category_count
            )
            if chart.show_vertical_gridlines:
                canvas.drawLine(
                    chart.plot_x, y,
                    chart.plot_x + chart.plot_width, y,
                    category_grid_paint,
                )
            canvas.drawLine(
                chart.plot_x - 2.8, y,
                chart.plot_x, y,
                grid_paint,
            )

        for label, position in chart.axis_categories:
            y = chart.plot_y + chart.plot_height * (1.0 - position)
            label_width = axis_font.measureText(label)
            baseline = y - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0 - 0.5
            canvas.drawString(
                label, chart.plot_x - 8.25 - label_width, baseline,
                axis_font, category_text_paint,
            )

        if chart.value_axis_display_unit_label:
            label = chart.value_axis_display_unit_label
            label_width = axis_title_font.measureText(label)
            label_metrics = axis_title_font.getMetrics()
            if value_axis_at_top:
                label_baseline = (
                    chart.plot_y - 19.67 - label_metrics.fDescent
                )
            else:
                label_baseline = (
                    plot_bottom + 19.67 - label_metrics.fAscent
                )
            canvas.drawString(
                label,
                chart.plot_x + chart.plot_width - label_width - 2.62,
                label_baseline,
                axis_title_font,
                value_text_paint,
            )

        if chart.plot_border_color:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    chart.plot_x, chart.plot_y,
                    chart.plot_width, chart.plot_height,
                ),
                self._paint(
                    'FF' + chart.plot_border_color, stroke_width=0.75
                ),
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

        if not chart.show_legend or not chart.series:
            return
        legend_series = self._chart_bar_legend_series(chart)
        item_widths = [
            16.33 + axis_font.measureText(series.name)
            for series in legend_series
        ]
        legend_x = (
            chart.x + (chart.width - sum(item_widths)) / 2.0 + 7.165
        )
        legend_y = chart.y + chart.height - 16.91
        legend_baseline = (
            chart.y + chart.height - 20.43 - axis_metrics.fAscent
        )
        for series, item_width in zip(legend_series, item_widths):
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    legend_x, legend_y, 5.0, 5.0
                ),
                self._paint('FF' + series.color, fill=True),
            )
            canvas.drawString(
                series.name, legend_x + 7.1, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += item_width

    @staticmethod
    def _chart_bar_legend_series(chart):
        if chart.grouping in ('stacked', 'percentStacked'):
            return tuple(chart.series)
        return tuple(reversed(chart.series))

    @staticmethod
    def _chart_tick_span(origin, tick_mark, inside_direction, length):
        inside = origin + inside_direction * length
        outside = origin - inside_direction * length
        if tick_mark == 'in':
            return inside, origin
        if tick_mark == 'out':
            return origin, outside
        if tick_mark == 'cross':
            return outside, inside
        return None

    @classmethod
    def _chart_value_axis_tick_segments(cls, chart, horizontal):
        axis_ticks = cls._chart_value_axis_ticks(chart)
        segments = []

        def append_ticks(mark, ratios, length):
            if mark == 'none':
                return
            for ratio in ratios:
                if horizontal:
                    position = chart.plot_x + chart.plot_width * ratio
                    if chart.value_axis_position == 't':
                        origin = chart.plot_y
                        inside_direction = 1.0
                    else:
                        origin = chart.plot_y + chart.plot_height
                        inside_direction = -1.0
                    span = cls._chart_tick_span(
                        origin, mark, inside_direction, length
                    )
                    segments.append((
                        position, span[0], position, span[1]
                    ))
                else:
                    position = chart.plot_y + chart.plot_height * (
                        1.0 - ratio
                    )
                    origin = chart.plot_x
                    span = cls._chart_tick_span(
                        origin, mark, 1.0, length
                    )
                    segments.append((
                        span[0], position, span[1], position
                    ))

        if chart.value_axis_log_base:
            import math

            base = chart.value_axis_log_base
            minimum_exponent = round(math.log(chart.value_min, base))
            maximum_exponent = round(math.log(chart.value_max, base))
            factors = range(1, max(2, int(round(base))))
            minor_ratios = tuple(
                cls._chart_value_ratio(
                    chart, (base ** exponent) * factor
                )
                for exponent in range(minimum_exponent, maximum_exponent)
                for factor in factors
            ) + (1.0,)
        else:
            interval_count = max(1, len(axis_ticks) - 1) * 5
            minor_ratios = tuple(
                index / interval_count
                for index in range(interval_count + 1)
            )
        append_ticks(
            chart.value_minor_tick_mark, minor_ratios, 2.11
        )
        append_ticks(
            chart.value_major_tick_mark,
            tuple(ratio for _, ratio in axis_ticks),
            2.82,
        )
        return tuple(segments)

    @staticmethod
    def _chart_value_ratio(chart, value):
        if chart.value_axis_log_base:
            import math

            base = chart.value_axis_log_base
            minimum = max(1e-300, float(chart.value_min))
            maximum = max(minimum * base, float(chart.value_max))
            numeric_value = max(minimum, float(value))
            ratio = (
                math.log(numeric_value, base) - math.log(minimum, base)
            ) / (
                math.log(maximum, base) - math.log(minimum, base)
            )
        else:
            span = max(1e-12, chart.value_max - chart.value_min)
            ratio = (float(value) - chart.value_min) / span
        return max(0.0, min(1.0, ratio))

    @classmethod
    def _chart_value_axis_ticks(cls, chart):
        if chart.value_axis_log_base:
            import math

            base = chart.value_axis_log_base
            minimum_exponent = round(math.log(chart.value_min, base))
            maximum_exponent = round(math.log(chart.value_max, base))
            return tuple(
                (base ** exponent, cls._chart_value_ratio(
                    chart, base ** exponent
                ))
                for exponent in range(
                    minimum_exponent, maximum_exponent + 1
                )
            )
        tick_count = max(1, int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        )))
        return tuple(
            (
                chart.value_min + chart.value_step * index,
                index / tick_count,
            )
            for index in range(tick_count + 1)
        )

    @staticmethod
    def _draw_chart_axis_titles(
        canvas, chart, axis_title_font,
        value_text_paint, category_text_paint,
    ):
        metrics = axis_title_font.getMetrics()
        if chart.category_axis_title:
            text_width = axis_title_font.measureText(chart.category_axis_title)
            x = chart.plot_x + (chart.plot_width - text_width) / 2.0
            y = (
                chart.plot_y + chart.plot_height
                + 20.0 - metrics.fAscent
            )
            canvas.drawString(
                chart.category_axis_title, x, y,
                axis_title_font, category_text_paint,
            )
        if chart.value_axis_title:
            text_width = axis_title_font.measureText(chart.value_axis_title)
            canvas.save()
            canvas.translate(
                chart.x + 23.0,
                chart.plot_y + chart.plot_height / 2.0,
            )
            canvas.rotate(-90.0)
            canvas.drawString(
                chart.value_axis_title, -text_width / 2.0, 0.0,
                axis_title_font, value_text_paint,
            )
            canvas.restore()

    def _draw_column_chart_series(
        self, canvas, chart, axis_font, data_label_paint
    ):
        rectangles = self._chart_column_rectangles(chart)
        for series_index, _, x, y, width, height in rectangles:
            if height <= 0.0:
                continue
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(x, y, width, height),
                self._paint(
                    'FF' + chart.series[series_index].color, fill=True
                ),
            )
        self._draw_column_chart_error_bars(canvas, chart, rectangles)
        if not chart.show_data_labels:
            return
        metrics = axis_font.getMetrics()
        for text, center_x, center_y, _, _ in (
            self._chart_column_data_label_anchors(chart, rectangles)
        ):
            text_width = axis_font.measureText(text)
            baseline = center_y - (
                metrics.fAscent + metrics.fDescent
            ) / 2.0
            canvas.drawString(
                text, center_x - text_width / 2.0, baseline,
                axis_font, data_label_paint,
            )

    def _draw_column_chart_error_bars(self, canvas, chart, rectangles):
        import math
        import statistics

        if not any(series.error_bars for series in chart.series):
            return
        totals = [
            sum(
                float(series.values[index])
                for series in chart.series
                if index < len(series.values)
            )
            for index in range(len(chart.category_positions))
        ]
        cumulative = [0.0] * len(chart.category_positions)
        for series_index, series in enumerate(chart.series):
            cumulative = [
                cumulative[index] + float(value)
                for index, value in enumerate(series.values)
            ]
            for error in series.error_bars:
                val_type, bar_type, no_end_cap, width, color, value = error
                if val_type == 'stdErr':
                    error_values = (
                        cumulative
                        if chart.grouping in ('stacked', 'percentStacked')
                        else [float(item) for item in series.values]
                    )
                    rendered = [
                        error_values[index] / totals[index]
                        if chart.grouping == 'percentStacked' and totals[index]
                        else error_values[index]
                        for index in range(len(error_values))
                    ]
                    amount = (
                        statistics.stdev(rendered) / math.sqrt(len(rendered))
                        if len(rendered) > 1 else 0.0
                    )
                elif val_type == 'percentage':
                    amount = abs(value) / 100.0
                elif val_type == 'fixedVal':
                    amount = abs(value)
                else:
                    continue
                delta = chart.plot_height * amount / max(
                    1e-12, chart.value_max - chart.value_min
                )
                paint = self._paint('FF' + color, stroke_width=width)
                for rect_series, _, x, y, rect_width, _ in rectangles:
                    if rect_series != series_index:
                        continue
                    center_x = x + rect_width / 2.0
                    top = max(chart.plot_y, y - delta)
                    bottom = min(
                        chart.plot_y + chart.plot_height, y + delta
                    )
                    if bar_type == 'plus':
                        bottom = y
                    elif bar_type == 'minus':
                        top = y
                    canvas.drawLine(center_x, top, center_x, bottom, paint)
                    if not no_end_cap:
                        if top > chart.plot_y or bar_type != 'both':
                            canvas.drawLine(
                                center_x - 2.25, top,
                                center_x + 2.25, top, paint,
                            )
                        if bottom < chart.plot_y + chart.plot_height:
                            canvas.drawLine(
                                center_x - 2.25, bottom,
                                center_x + 2.25, bottom, paint,
                            )

    def _draw_chart_data_table(
        self, canvas, chart, axis_font, text_paint, grid_paint
    ):
        if not chart.show_data_table or not chart.categories:
            return
        table_left = chart.x + 11.0
        data_left = chart.plot_x
        table_right = chart.plot_x + chart.plot_width
        table_top = chart.plot_y + chart.plot_height + 23.5
        row_height = 14.4
        rows = tuple(reversed(chart.series))
        column_width = chart.plot_width / len(chart.categories)
        table_bottom = table_top + row_height * (len(rows) + 1)
        for row in range(len(rows) + 2):
            y = table_top + row * row_height
            start_x = data_left if row == 0 else table_left
            canvas.drawLine(start_x, y, table_right, y, grid_paint)
        canvas.drawLine(
            table_left, table_top + row_height,
            table_left, table_bottom, grid_paint,
        )
        for column in range(len(chart.categories) + 1):
            x = data_left + column * column_width
            canvas.drawLine(x, table_top, x, table_bottom, grid_paint)
        metrics = axis_font.getMetrics()

        def baseline(row):
            center = table_top + (row + 0.5) * row_height
            return center - (metrics.fAscent + metrics.fDescent) / 2.0

        for column, label in enumerate(chart.categories):
            text_width = axis_font.measureText(label)
            canvas.drawString(
                label,
                data_left + (column + 0.5) * column_width - text_width / 2.0,
                baseline(0), axis_font, text_paint,
            )
        for row, series in enumerate(rows, start=1):
            text_x = table_left + 6.0
            if chart.data_table_show_keys:
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        table_left + 3.2,
                        table_top + row * row_height + 4.7,
                        5.0, 5.0,
                    ),
                    self._paint('FF' + series.color, fill=True),
                )
                text_x = table_left + 10.2
            canvas.drawString(
                series.name, text_x, baseline(row), axis_font, text_paint
            )
            for column, value in enumerate(series.display_values):
                text = str(value)
                text_width = axis_font.measureText(text)
                canvas.drawString(
                    text,
                    data_left + (column + 0.5) * column_width
                    - text_width / 2.0,
                    baseline(row), axis_font, text_paint,
                )

    @classmethod
    def _chart_column_data_label_anchors(cls, chart, rectangles=None):
        if not chart.show_data_labels:
            return ()
        if rectangles is None:
            rectangles = cls._chart_column_rectangles(chart)
        anchors = []
        for series_index, category_index, x, y, width, height in rectangles:
            value = chart.series[series_index].values[category_index]
            anchors.append((
                cls._chart_value_label(value, chart.axis_number_format),
                x + width / 2.0,
                y + height / 2.0,
                series_index,
                category_index,
            ))
        return tuple(anchors)

    @staticmethod
    def _chart_column_rectangles(chart):
        category_positions = chart.category_positions
        series_count = len(chart.series)
        if not category_positions or series_count == 0:
            return ()

        slot_count = max(len(chart.axis_categories), len(category_positions))
        slot_width = chart.plot_width / max(1, slot_count)
        is_stacked = chart.grouping in ('stacked', 'percentStacked')
        overlap_ratio = max(-1.0, min(1.0, chart.overlap / 100.0))
        series_step_ratio = 0.0 if is_stacked else 1.0 - overlap_ratio
        gap_ratio = max(0.0, chart.gap_width / 100.0)
        width_ratio = (
            1.0 + (series_count - 1) * series_step_ratio + gap_ratio
        )
        bar_width = slot_width / width_ratio
        series_step = bar_width * series_step_ratio
        cluster_width = bar_width + (series_count - 1) * series_step

        plot_bottom = chart.plot_y + chart.plot_height
        positive_stacks = [0.0] * len(category_positions)
        negative_stacks = [0.0] * len(category_positions)
        category_totals = [
            sum(
                float(series.values[index])
                for series in chart.series
                if index < len(series.values)
            )
            for index in range(len(category_positions))
        ]

        def value_y(value):
            ratio = PdfRendererPieBarMixin._chart_value_ratio(chart, value)
            return chart.plot_y + chart.plot_height * (1.0 - ratio)

        rectangles = []
        for series_index, series in enumerate(chart.series):
            for category_index, (position, value) in enumerate(zip(
                category_positions, series.values
            )):
                numeric_value = float(value)
                if is_stacked:
                    stacks = (
                        positive_stacks
                        if numeric_value >= 0.0 else negative_stacks
                    )
                    start_value = stacks[category_index]
                    end_value = start_value + numeric_value
                    stacks[category_index] = end_value
                    if chart.grouping == 'percentStacked':
                        total = category_totals[category_index]
                        if total:
                            start_value /= total
                            end_value /= total
                        else:
                            start_value = end_value = 0.0
                else:
                    start_value = 0.0
                    end_value = numeric_value
                start_y = value_y(start_value)
                end_y = value_y(end_value)
                top = max(chart.plot_y, min(start_y, end_y))
                bottom = min(plot_bottom, max(start_y, end_y))
                center_x = chart.plot_x + chart.plot_width * position
                x = (
                    center_x - cluster_width / 2.0
                    + series_index * series_step
                )
                rectangles.append((
                    series_index,
                    category_index,
                    x,
                    top,
                    bar_width,
                    max(0.0, bottom - top),
                ))
        return tuple(rectangles)

    @staticmethod
    def _chart_bar_rectangles(chart):
        category_positions = chart.category_positions
        series_count = len(chart.series)
        if not category_positions or series_count == 0:
            return ()

        slot_count = max(len(chart.axis_categories), len(category_positions))
        slot_height = chart.plot_height / max(1, slot_count)
        is_stacked = chart.grouping in ('stacked', 'percentStacked')
        overlap_ratio = max(-1.0, min(1.0, chart.overlap / 100.0))
        series_step_ratio = 0.0 if is_stacked else 1.0 - overlap_ratio
        gap_ratio = max(0.0, chart.gap_width / 100.0)
        height_ratio = (
            1.0 + (series_count - 1) * series_step_ratio + gap_ratio
        )
        bar_height = slot_height / height_ratio
        series_step = bar_height * series_step_ratio
        cluster_height = bar_height + (series_count - 1) * series_step

        plot_right = chart.plot_x + chart.plot_width
        positive_stacks = [0.0] * len(category_positions)
        negative_stacks = [0.0] * len(category_positions)
        category_totals = [
            sum(
                float(series.values[index])
                for series in chart.series
                if index < len(series.values)
            )
            for index in range(len(category_positions))
        ]

        def value_x(value):
            ratio = PdfRendererPieBarMixin._chart_value_ratio(chart, value)
            return chart.plot_x + chart.plot_width * ratio

        rectangles = []
        for series_index, series in enumerate(chart.series):
            for category_index, (position, value) in enumerate(zip(
                category_positions, series.values
            )):
                numeric_value = float(value)
                if is_stacked:
                    stacks = (
                        positive_stacks
                        if numeric_value >= 0.0 else negative_stacks
                    )
                    start_value = stacks[category_index]
                    end_value = start_value + numeric_value
                    stacks[category_index] = end_value
                    if chart.grouping == 'percentStacked':
                        total = category_totals[category_index]
                        if total:
                            start_value /= total
                            end_value /= total
                        else:
                            start_value = end_value = 0.0
                else:
                    start_value = 0.0
                    end_value = numeric_value
                start_x = value_x(start_value)
                end_x = value_x(end_value)
                left = max(chart.plot_x, min(start_x, end_x))
                right = min(plot_right, max(start_x, end_x))
                center_y = (
                    chart.plot_y
                    + chart.plot_height * (1.0 - position)
                )
                y = (
                    center_y - cluster_height / 2.0
                    + (series_count - 1 - series_index) * series_step
                )
                rectangles.append((
                    series_index,
                    category_index,
                    left,
                    y,
                    max(0.0, right - left),
                    bar_height,
                ))
        return tuple(rectangles)

    def _chart_background_paint(self, chart):
        import math

        if chart.background_pattern_type in self._CHART_PATTERN_ROWS:
            return self._chart_pattern_paint(
                chart.background_pattern_type,
                chart.background_pattern_foreground,
                chart.background_pattern_background,
            )
        if chart.background_gradient_stops:
            angle = math.radians(chart.background_gradient_angle)
            direction_x = math.cos(angle)
            direction_y = math.sin(angle)
            center_x = chart.x + chart.width / 2.0
            center_y = chart.y + chart.height / 2.0
            half_length = (
                abs(chart.width * direction_x)
                + abs(chart.height * direction_y)
            ) / 2.0
            start = self._skia.Point(
                center_x - direction_x * half_length,
                center_y - direction_y * half_length,
            )
            end = self._skia.Point(
                center_x + direction_x * half_length,
                center_y + direction_y * half_length,
            )
            colors = []
            positions = []
            stops = sorted(
                chart.background_gradient_stops,
                key=lambda item: item[0],
            )
            for position, color in stops:
                alpha, red, green, blue = self._argb_channels(color)
                colors.append(self._skia.ColorSetARGB(alpha, red, green, blue))
                positions.append(max(0.0, min(1.0, position)))
            paint = self._skia.Paint(AntiAlias=True)
            paint.setShader(self._skia.GradientShader.MakeLinear(
                (start, end), colors, positions,
                self._skia.TileMode.kClamp,
            ))
            return paint
        color = chart.background_fill_color or 'FFFFFF'
        return self._paint('FF' + color, fill=True)

    def _chart_pattern_paint(
        self, pattern_type, foreground_color, background_color,
    ):
        key = ('chart', pattern_type, foreground_color, background_color)
        cached = self._pattern_images.get(key)
        if cached is None:
            rows = self._CHART_PATTERN_ROWS[pattern_type]
            foreground = self._rgba_channels('FF' + foreground_color)
            background = self._rgba_channels('FF' + background_color)
            pixels = bytes(
                channel
                for row in rows
                for pixel in row
                for channel in (
                    foreground if pixel == '#' else background
                )
            )
            size = len(rows)
            info = self._skia.ImageInfo.Make(
                size,
                size,
                self._skia.ColorType.kRGBA_8888_ColorType,
                self._skia.AlphaType.kUnpremul_AlphaType,
            )
            image = self._skia.Image.MakeRasterData(
                info, pixels, size * 4
            )
            if image is None:
                raise RuntimeError('Skia could not create a chart fill pattern')
            cached = (image, pixels)
            self._pattern_images[key] = cached
        pattern_scale = 0.35 if pattern_type == 'ltUpDiag' else 0.25
        shader = cached[0].makeShader(
            self._skia.TileMode.kRepeat,
            self._skia.TileMode.kRepeat,
            self._skia.SamplingOptions(self._skia.FilterMode.kNearest),
            self._skia.Matrix.Scale(pattern_scale, pattern_scale),
        )
        paint = self._paint('FFFFFFFF', fill=True)
        paint.setShader(shader)
        return paint

