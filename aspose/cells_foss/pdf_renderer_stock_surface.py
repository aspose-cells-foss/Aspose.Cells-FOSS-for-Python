"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRendererStockSurfaceMixin:
    def _draw_stock_chart_series(self, canvas, chart):
        visible_series = tuple(
            series for series in chart.series if not series.hidden
        )
        if not visible_series:
            return
        price_offset = 1 if chart.stock_has_volume else 0
        price_series = visible_series[price_offset:]
        if chart.stock_has_volume:
            self._draw_stock_volume_columns(
                canvas, chart, visible_series[0]
            )
        if len(price_series) < 3:
            return
        point_count = min(
            len(chart.category_positions),
            *(len(series.values) for series in price_series),
        )
        high_low_paint = self._paint('FF404040', stroke_width=0.75)
        open_series = price_series[0] if chart.stock_has_up_down_bars else None
        high_series = price_series[1] if open_series is not None else price_series[0]
        low_series = price_series[2] if open_series is not None else price_series[1]
        close_series = price_series[-1]
        close_paint = self._paint(
            'FF' + close_series.color, stroke_width=0.75
        )
        close_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
        slot_width = chart.plot_width / max(1, point_count)
        bar_width = slot_width / max(1.0, 1.0 + chart.gap_width / 100.0)
        bar_border = self._paint('FF595959', stroke_width=0.75)
        for index in range(point_count):
            x = (
                chart.plot_x
                + chart.plot_width * chart.category_positions[index]
            )
            high_y = self._stock_price_y(
                chart, high_series.values[index]
            )
            low_y = self._stock_price_y(chart, low_series.values[index])
            canvas.drawLine(x, high_y, x, low_y, high_low_paint)

            close_y = self._stock_price_y(
                chart, close_series.values[index]
            )
            if open_series is not None:
                open_value = float(open_series.values[index])
                close_value = float(close_series.values[index])
                open_y = self._stock_price_y(chart, open_value)
                top = min(open_y, close_y)
                rect = self._skia.Rect.MakeXYWH(
                    x - bar_width / 2.0,
                    top,
                    bar_width,
                    max(0.75, abs(close_y - open_y)),
                )
                canvas.drawRect(
                    rect,
                    self._paint(
                        'FFFFFFFF' if close_value >= open_value
                        else 'FF404040',
                        fill=True,
                    ),
                )
                canvas.drawRect(rect, bar_border)
            canvas.drawLine(
                x, close_y, x + 1.44, close_y, close_paint
            )

    def _draw_stock_volume_columns(self, canvas, chart, volume_series):
        point_count = min(
            len(chart.category_positions), len(volume_series.values)
        )
        slot_width = chart.plot_width / max(1, point_count)
        bar_width = slot_width / max(1.0, 1.0 + chart.gap_width / 100.0)
        baseline = chart.plot_y + chart.plot_height
        paint = self._paint('FF' + volume_series.color, fill=True)
        for position, value in zip(
            chart.category_positions[:point_count],
            volume_series.values[:point_count],
        ):
            x = chart.plot_x + chart.plot_width * position
            y = chart.plot_y + chart.plot_height * (
                1.0 - self._chart_value_ratio(chart, value)
            )
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    x - bar_width / 2.0, y, bar_width, baseline - y
                ),
                paint,
            )

    @staticmethod
    def _stock_price_ratio(chart, value):
        span = chart.stock_price_max - chart.stock_price_min
        if span <= 0.0:
            return 0.0
        return max(0.0, min(
            1.0, (float(value) - chart.stock_price_min) / span
        ))

    def _stock_price_y(self, chart, value):
        ratio = (
            self._stock_price_ratio(chart, value)
            if chart.stock_has_volume
            else self._chart_value_ratio(chart, value)
        )
        return chart.plot_y + chart.plot_height * (1.0 - ratio)

    def _draw_stock_price_axis(
        self, canvas, chart, axis_font, axis_metrics, text_paint
    ):
        if chart.stock_axes_hidden or chart.stock_price_step <= 0.0:
            return
        count = int(round(
            (chart.stock_price_max - chart.stock_price_min)
            / chart.stock_price_step
        ))
        axis_paint = self._paint('FFBFBFBF', stroke_width=0.75)
        right = chart.plot_x + chart.plot_width
        canvas.drawLine(
            right, chart.plot_y, right, chart.plot_y + chart.plot_height,
            axis_paint,
        )
        for index in range(count + 1):
            value = chart.stock_price_min + index * chart.stock_price_step
            ratio = index / max(1, count)
            y = chart.plot_y + chart.plot_height * (1.0 - ratio)
            baseline = (
                y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                - 0.6
            )
            canvas.drawString(
                str(int(round(value))), right + 5.0, baseline,
                axis_font, text_paint,
            )

    def _draw_stock_combo_series(self, canvas, chart):
        visible_series = tuple(
            series for series in chart.series if not series.hidden
        )
        if not visible_series:
            return
        price_offset = 1 if chart.stock_has_volume else 0
        price_series = visible_series[price_offset:]
        point_count = min(
            len(chart.category_positions),
            *(len(series.values) for series in visible_series),
        )
        slot_width = chart.plot_width / max(1, point_count)
        bar_width = slot_width / max(1.0, 1.0 + chart.gap_width / 100.0)

        if chart.stock_has_volume:
            volume = visible_series[0]
            baseline = chart.plot_y + chart.plot_height
            for position, value in zip(
                chart.category_positions[:point_count],
                volume.values[:point_count],
            ):
                x = chart.plot_x + chart.plot_width * position
                y = chart.plot_y + chart.plot_height * (
                    1.0 - self._chart_value_ratio(chart, value)
                )
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        x - bar_width / 2.0, y, bar_width, baseline - y
                    ),
                    self._paint('FF' + volume.color, fill=True),
                )

        if chart.stock_has_up_down_bars and len(price_series) >= 2:
            open_series = price_series[0]
            close_series = price_series[-1]
            border = self._paint('FF000000', stroke_width=0.75)
            for index, position in enumerate(
                chart.category_positions[:point_count]
            ):
                open_value = float(open_series.values[index])
                close_value = float(close_series.values[index])
                open_y = chart.plot_y + chart.plot_height * (
                    1.0 - self._chart_value_ratio(chart, open_value)
                )
                close_y = chart.plot_y + chart.plot_height * (
                    1.0 - self._chart_value_ratio(chart, close_value)
                )
                x = chart.plot_x + chart.plot_width * position
                top = min(open_y, close_y)
                height = max(0.75, abs(close_y - open_y))
                rect = self._skia.Rect.MakeXYWH(
                    x - bar_width / 2.0, top, bar_width, height
                )
                canvas.drawRect(
                    rect,
                    self._paint(
                        'FFFFFFFF' if close_value >= open_value
                        else 'FF3F3F3F',
                        fill=True,
                    ),
                )
                canvas.drawRect(rect, border)

        for series in price_series:
            path = self._skia.Path()
            rendered_points = []
            for index, (position, value) in enumerate(zip(
                chart.category_positions[:point_count],
                series.values[:point_count],
            )):
                x = chart.plot_x + chart.plot_width * position
                y = chart.plot_y + chart.plot_height * (
                    1.0 - self._chart_value_ratio(chart, value)
                )
                rendered_points.append((x, y))
                if index == 0:
                    path.moveTo(x, y)
                else:
                    path.lineTo(x, y)
            paint = self._paint(
                'FF' + series.color, stroke_width=series.line_width
            )
            canvas.drawPath(path, paint)
            for x, y in rendered_points:
                self._draw_chart_marker(canvas, series, x, y)

    def _draw_stock_chart_legend(
        self, canvas, chart, axis_font, axis_metrics, legend_text_paint
    ):
        visible_series = tuple(
            series for series in chart.series if not series.hidden
        )
        if not visible_series:
            return
        if chart.legend_position == 'r' and not chart.stock_has_high_low_lines:
            legend_x = chart.x + chart.width - 57.0
            legend_y = chart.y + 73.0
            for index, series in enumerate(visible_series):
                y = legend_y + index * 18.0
                if chart.stock_has_volume and index == 0:
                    canvas.drawRect(
                        self._skia.Rect.MakeXYWH(legend_x + 7.0, y, 6.0, 6.0),
                        self._paint('FF' + series.color, fill=True),
                    )
                else:
                    canvas.drawLine(
                        legend_x, y + 3.0, legend_x + 19.2, y + 3.0,
                        self._paint(
                            'FF' + series.color,
                            stroke_width=series.line_width,
                        ),
                    )
                    self._draw_chart_marker(
                        canvas, series, legend_x + 9.6, y + 3.0
                    )
                canvas.drawString(
                    series.name, legend_x + 22.0,
                    y - axis_metrics.fAscent,
                    axis_font, legend_text_paint,
                )
            return
        item_widths = [
            15.5 + axis_font.measureText(series.name)
            for series in visible_series
        ]
        legend_x = (
            chart.x + (chart.width - sum(item_widths)) / 2.0 - 2.2
        )
        legend_y = chart.y + chart.height - 20.6
        legend_baseline = legend_y - axis_metrics.fAscent
        for index, (series, item_width) in enumerate(zip(
            visible_series, item_widths
        )):
            if chart.stock_has_volume and index == 0:
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        legend_x + 7.0, legend_y + 3.2, 5.9, 5.9
                    ),
                    self._paint('FF' + series.color, fill=True),
                )
            elif index == len(visible_series) - 1:
                marker_paint = self._paint(
                    'FF' + series.color, stroke_width=0.75
                )
                marker_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
                canvas.drawLine(
                    legend_x + 8.9, legend_y + 6.16,
                    legend_x + 10.3, legend_y + 6.16,
                    marker_paint,
                )
            canvas.drawString(
                series.name, legend_x + 15.5, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += item_width

    def _draw_3d_surface_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint,
    ):
        visible = tuple(series for series in chart.series if not series.hidden)
        point_count = min(
            (len(series.values) for series in visible), default=0
        )
        if len(visible) < 2 or point_count < 2:
            return

        axes_hidden = bool(chart.surface_axes_hidden)
        actual_minimum = min(
            (float(value) for series in visible for value in series.values),
            default=chart.value_min,
        )
        maximum_value = max(
            (float(value) for series in visible for value in series.values),
            default=chart.value_max,
        )
        if axes_hidden:
            band_minimum = actual_minimum
            band_maximum = maximum_value
            band_step = max(
                (14.5 - actual_minimum) / 2.0, 1e-12
            )
            band_bounds = (
                band_minimum,
                min(band_minimum + band_step, band_maximum),
                min(band_minimum + 2.0 * band_step, band_maximum),
                band_maximum,
            )
            band_count = 3
        else:
            band_minimum = chart.value_min
            band_maximum = maximum_value
            raw_step = max(
                (maximum_value - chart.value_min) / 5.0, 1e-12
            )
            magnitude = 10.0 ** math.floor(math.log10(raw_step))
            normalized = raw_step / magnitude
            band_step = next(
                factor for factor in (1.0, 2.0, 5.0, 10.0)
                if normalized <= factor
            ) * magnitude
            band_count = max(1, int(math.ceil(
                (maximum_value - chart.value_min) / band_step
            )))
            band_bounds = tuple(
                min(
                    band_minimum + index * band_step,
                    band_maximum,
                )
                for index in range(band_count + 1)
            )
        colors = chart.surface_band_colors or (
            '156082', 'E97132', '196B24', '0F9ED5', 'A02B93',
            '4EA72E', '7F8FA3', 'EF927E', '849E88', '7AB2DD',
            'BB7DB4', '89B883',
        )

        width_scale = chart.width / 377.4
        height_scale = chart.height / 219.1

        def project(category, depth, value):
            # Excel's default 20/15 perspective is close to a bilinear floor
            # projection with a depth-dependent vertical value vector. Both
            # filled and wireframe variants share this view geometry.
            category_squared = category * category
            ground_x = (
                52.0 + 204.0 * category
                + 44.0 * category_squared + 79.0 * depth
                - 50.0 * category * depth
                - 2.0 * category_squared * depth
            )
            ground_y = (
                139.0 + 31.0 * category - 34.0 * depth
                - 25.0 * category * depth
            )
            value_scale = 3.04 * (
                1.0 - 0.145 * depth
                - 0.07 * category * depth
            )
            # The value vector fans out across the foreground: it leans
            # slightly left at Jan and slightly right at Apr, then becomes
            # upright at the back wall.
            value_skew = (
                0.13 - 0.50 * category
            ) * (1.0 - depth)
            screen_x = (
                chart.x + (ground_x - value_skew * value) * width_scale
            )
            screen_y = (
                chart.y + (ground_y - value * value_scale) * height_scale
            )
            if axes_hidden:
                screen_x += (
                    42.0 - 27.0 * depth - 18.0 * category * category
                )
                screen_y -= 33.0 * depth
                band_value = value
            else:
                band_value = value
            return screen_x, screen_y, band_value

        grid_paint = self._paint(
            'FF202020' if axes_hidden else 'FFD9D9D9',
            stroke_width=0.75,
        )
        if axes_hidden:
            frame_width_scale = chart.width / 425.7
            frame_height_scale = chart.height / 252.7

            def frame_point(local_x, local_y):
                return (
                    chart.x + local_x * frame_width_scale,
                    chart.y + local_y * frame_height_scale,
                )

            for depth in (0.0, 0.5, 1.0):
                first = frame_point(
                    24.5 + 76.5 * depth
                    + 12.0 * depth * (1.0 - depth),
                    191.5 - 48.0 * depth,
                )
                second = frame_point(
                    359.0 + 32.0 * depth,
                    235.5 - 62.0 * depth,
                )
                canvas.drawLine(
                    first[0], first[1], second[0], second[1], grid_paint
                )
            for depth in (0.0, 1.0 / 3.0, 2.0 / 3.0, 1.0):
                base = frame_point(
                    24.5 + 107.5 * depth
                    + 48.0 * depth * (1.0 - depth),
                    191.5 - 91.5 * depth,
                )
                top = frame_point(
                    16.5 + 112.5 * depth
                    + 48.0 * depth * (1.0 - depth),
                    49.0 - 34.0 * depth,
                )
                canvas.drawLine(
                    base[0], base[1], top[0], top[1], grid_paint
                )
        else:
            for index in range(band_count + 1):
                value = chart.value_min + index * band_step
                left_front = project(0.0, 0.0, value)
                left_back = project(0.0, 1.0, value)
                right_back = project(1.0, 1.0, value)
                path = self._skia.Path()
                path.moveTo(left_front[0], left_front[1])
                path.lineTo(left_back[0], left_back[1])
                path.lineTo(right_back[0], right_back[1])
                canvas.drawPath(path, grid_paint)
                label = self._chart_value_label(
                    value, chart.axis_number_format
                )
                label_width = axis_font.measureText(label)
                metrics = axis_font.getMetrics()
                baseline = left_front[1] - (
                    metrics.fAscent + metrics.fDescent
                ) / 2.0
                canvas.drawString(
                    label, left_front[0] - 8.0 - label_width, baseline,
                    axis_font, value_text_paint,
                )

        def intersect(first, second, threshold):
            delta = second[2] - first[2]
            ratio = 0.0 if abs(delta) < 1e-12 else (
                threshold - first[2]
            ) / delta
            return (
                first[0] + (second[0] - first[0]) * ratio,
                first[1] + (second[1] - first[1]) * ratio,
                threshold,
            )

        def clip_polygon(points, threshold, keep_above):
            result = []
            for first, second in zip(points, points[1:] + points[:1]):
                first_inside = (
                    first[2] >= threshold if keep_above
                    else first[2] <= threshold
                )
                second_inside = (
                    second[2] >= threshold if keep_above
                    else second[2] <= threshold
                )
                if first_inside:
                    result.append(first)
                if first_inside != second_inside:
                    result.append(intersect(first, second, threshold))
            return result

        projected_rows = tuple(
            tuple(
                project(
                    column / (point_count - 1),
                    row / (len(visible) - 1),
                    float(series.values[column]),
                )
                for column in range(point_count)
            )
            for row, series in enumerate(visible)
        )
        if chart.surface_wireframe:
            def draw_banded_segment(first, second):
                points = [first]
                low_value = min(first[2], second[2])
                high_value = max(first[2], second[2])
                for band in range(1, band_count + 1):
                    threshold = band_minimum + band * band_step
                    if low_value < threshold < high_value:
                        points.append(intersect(first, second, threshold))
                points.append(second)
                points.sort(
                    key=lambda point: (
                        (point[0] - first[0]) ** 2
                        + (point[1] - first[1]) ** 2
                    )
                )
                for start, end in zip(points, points[1:]):
                    midpoint = (start[2] + end[2]) / 2.0
                    band = int(math.floor(
                        (midpoint - band_minimum) / band_step
                    ))
                    band = max(0, min(band_count - 1, band))
                    canvas.drawLine(
                        start[0], start[1], end[0], end[1],
                        self._paint(
                            'FF' + colors[band % len(colors)],
                            stroke_width=0.8,
                        ),
                    )

            # A wireframe surface contains only the two orthogonal data-grid
            # directions. Drawing triangulation diagonals makes it look like a
            # filled surface with missing faces rather than Excel's wireframe.
            for row in reversed(range(len(visible))):
                for column in range(point_count - 1):
                    draw_banded_segment(
                        projected_rows[row][column],
                        projected_rows[row][column + 1],
                    )
            for column in range(point_count):
                for row in reversed(range(len(visible) - 1)):
                    draw_banded_segment(
                        projected_rows[row][column],
                        projected_rows[row + 1][column],
                    )

            def draw_contour_segment(triangle, threshold, color):
                intersections = []
                for first, second in zip(
                    triangle, triangle[1:] + triangle[:1]
                ):
                    first_delta = first[2] - threshold
                    second_delta = second[2] - threshold
                    if abs(first_delta) < 1e-9:
                        intersections.append(first)
                    if first_delta * second_delta < 0.0:
                        intersections.append(
                            intersect(first, second, threshold)
                        )
                unique = []
                for point in intersections:
                    if not any(
                        abs(point[0] - other[0]) < 1e-6
                        and abs(point[1] - other[1]) < 1e-6
                        for other in unique
                    ):
                        unique.append(point)
                if len(unique) < 2:
                    return
                first, second = max(
                    (
                        (first, second)
                        for index, first in enumerate(unique)
                        for second in unique[index + 1:]
                    ),
                    key=lambda pair: (
                        (pair[1][0] - pair[0][0]) ** 2
                        + (pair[1][1] - pair[0][1]) ** 2
                    ),
                )
                canvas.drawLine(
                    first[0], first[1], second[0], second[1],
                    self._paint('FF' + color, stroke_width=0.8),
                )

            # Excel overlays band-boundary contours on the orthogonal mesh.
            # These interpolated lines are especially visible through the
            # center of a 3-D wireframe surface.
            for row in reversed(range(len(visible) - 1)):
                for column in range(point_count - 1):
                    lower_left = projected_rows[row][column]
                    lower_right = projected_rows[row][column + 1]
                    upper_right = projected_rows[row + 1][column + 1]
                    upper_left = projected_rows[row + 1][column]
                    triangles = (
                        (lower_left, lower_right, upper_right),
                        (lower_left, upper_right, upper_left),
                    ) if column % 2 == 0 else (
                        (lower_left, lower_right, upper_left),
                        (lower_right, upper_right, upper_left),
                    )
                    minimum = min(point[2] for point in (
                        lower_left, lower_right, upper_right, upper_left
                    ))
                    maximum = max(point[2] for point in (
                        lower_left, lower_right, upper_right, upper_left
                    ))
                    for band in range(1, band_count + 1):
                        threshold = band_minimum + band * band_step
                        if not minimum < threshold < maximum:
                            continue
                        color = colors[(band - 1) % len(colors)]
                        for triangle in triangles:
                            draw_contour_segment(
                                triangle, threshold, color
                            )
        else:
            for row in reversed(range(len(visible) - 1)):
                for column in range(point_count - 1):
                    lower_left = projected_rows[row][column]
                    lower_right = projected_rows[row][column + 1]
                    upper_right = projected_rows[row + 1][column + 1]
                    upper_left = projected_rows[row + 1][column]
                    triangles = (
                        (lower_left, lower_right, upper_right),
                        (lower_left, upper_right, upper_left),
                    ) if column % 2 == 0 else (
                        (lower_left, lower_right, upper_left),
                        (lower_right, upper_right, upper_left),
                    )
                    for triangle_index, triangle in enumerate(triangles):
                        shade = (
                            0.86 + 0.10 * (column % 2)
                            + 0.02 * triangle_index
                        )
                        if axes_hidden:
                            shade *= 0.72
                        for band in range(band_count):
                            low = band_bounds[band]
                            high = band_bounds[band + 1]
                            polygon = clip_polygon(list(triangle), low, True)
                            if polygon:
                                polygon = clip_polygon(polygon, high, False)
                            if len(polygon) < 3:
                                continue
                            path = self._skia.Path()
                            path.moveTo(polygon[0][0], polygon[0][1])
                            for point in polygon[1:]:
                                path.lineTo(point[0], point[1])
                            path.close()
                            color = self._scale_chart_color(
                                colors[band % len(colors)], shade
                            )
                            canvas.drawPath(
                                path, self._paint('FF' + color, fill=True)
                            )
                            canvas.drawPath(
                                path,
                                self._paint(
                                    'FF' + color, stroke_width=0.24
                                ),
                            )

        axis_metrics = axis_font.getMetrics()
        if not axes_hidden:
            floor_left = project(0.0, 0.0, chart.value_min)
            floor_right = project(1.0, 0.0, chart.value_min)
            canvas.drawLine(
                floor_left[0], floor_left[1],
                floor_right[0], floor_right[1], grid_paint,
            )
            for index, label in enumerate(chart.categories[:point_count]):
                point = project(
                    index / (point_count - 1), 0.0, chart.value_min
                )
                text = str(label)
                canvas.drawString(
                    text, point[0] - axis_font.measureText(text) / 2.0,
                    point[1] + 6.0 - axis_metrics.fAscent,
                    axis_font, category_text_paint,
                )
            depth_axis_points = tuple(
                project(
                    1.0, row / (len(visible) - 1), chart.value_min
                )
                for row in range(len(visible))
            )
            depth_axis = self._skia.Path()
            depth_axis.moveTo(
                depth_axis_points[0][0], depth_axis_points[0][1]
            )
            for point in depth_axis_points[1:]:
                depth_axis.lineTo(point[0], point[1])
            canvas.drawPath(depth_axis, grid_paint)
            for point in depth_axis_points:
                canvas.drawLine(
                    point[0], point[1], point[0] + 5.0, point[1],
                    grid_paint,
                )
            for row, series in enumerate(visible):
                point = depth_axis_points[row]
                canvas.drawString(
                    series.name, point[0] + 7.0,
                    point[1] - (
                        axis_metrics.fAscent + axis_metrics.fDescent
                    ) / 2.0,
                    axis_font, category_text_paint,
                )

        render_title_font = title_font
        if axes_hidden:
            render_title_font = self._font(
                chart.title_font_resolution, 20.0
            )
        title_width = render_title_font.measureText(chart.title)
        title_metrics = render_title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + (11.0 if axes_hidden else 7.5)
            - title_metrics.fAscent,
            render_title_font,
            (
                self._paint('FF000000', fill=True)
                if axes_hidden else title_paint
            ),
        )
        if chart.show_legend:
            if axes_hidden and chart.legend_position == 'r':
                legend_x = chart.x + chart.width - 49.0
                legend_y = chart.y + 24.0
                legend_minimum = 12.0
                legend_maximum = 23.1
                legend_count = 12
                for row, index in enumerate(reversed(range(legend_count))):
                    low = legend_minimum + index
                    high = min(low + 1.0, legend_maximum)
                    label = f'{low:g}-{high:g}'
                    item_y = legend_y + row * 18.2
                    canvas.drawRect(
                        self._skia.Rect.MakeXYWH(
                            legend_x, item_y, 5.5, 5.5
                        ),
                        self._paint(
                            'FF' + colors[index % len(colors)], fill=True
                        ),
                    )
                    canvas.drawString(
                        label, legend_x + 8.0,
                        item_y - axis_metrics.fAscent - 1.0,
                        axis_font, legend_text_paint,
                    )
                return
            labels = tuple(
                f'{self._chart_value_label(band_minimum + index * band_step, chart.axis_number_format)}-'
                f'{self._chart_value_label(min(band_minimum + (index + 1) * band_step, band_maximum), chart.axis_number_format)}'
                for index in range(band_count)
            )
            widths = tuple(
                14.0 + axis_font.measureText(label) for label in labels
            )
            legend_x = chart.x + (chart.width - sum(widths)) / 2.0 + 4.0
            legend_y = chart.y + chart.height - 18.5
            baseline = legend_y - axis_metrics.fAscent - 0.5
            for index, (label, item_width) in enumerate(zip(labels, widths)):
                color = colors[index % len(colors)]
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        legend_x, legend_y + 1.2, 6.0, 6.0
                    ),
                    self._paint(
                        'FF' + color,
                        fill=not chart.surface_wireframe,
                        stroke_width=(
                            0.75 if chart.surface_wireframe else 1.0
                        ),
                    ),
                )
                canvas.drawString(
                    label, legend_x + 8.5, baseline,
                    axis_font, legend_text_paint,
                )
                legend_x += item_width

    def _draw_2d_surface_chart(
        self, canvas, chart, axis_font, title_font,
        category_text_paint, legend_text_paint, title_paint,
    ):
        visible = tuple(series for series in chart.series if not series.hidden)
        point_count = min(
            (len(series.values) for series in visible), default=0
        )
        if len(visible) < 2 or point_count < 2:
            return

        plot_left = chart.plot_x
        plot_top = chart.plot_y
        plot_right = plot_left + chart.plot_width
        plot_bottom = plot_top + chart.plot_height
        x_step = chart.plot_width / (point_count - 1)
        y_step = chart.plot_height / (len(visible) - 1)
        rows = tuple(
            tuple(
                (
                    plot_left + column * x_step,
                    plot_bottom - row * y_step,
                    float(series.values[column]),
                )
                for column in range(point_count)
            )
            for row, series in enumerate(visible)
        )

        maximum_value = max(
            (float(value) for series in visible for value in series.values),
            default=chart.value_max,
        )
        raw_step = max(
            (maximum_value - chart.value_min) / 5.0, 1e-12
        )
        magnitude = 10.0 ** math.floor(math.log10(raw_step))
        normalized = raw_step / magnitude
        nice_factor = next(
            factor for factor in (1.0, 2.0, 5.0, 10.0)
            if normalized <= factor
        )
        band_step = nice_factor * magnitude
        band_count = max(1, int(math.ceil(
            (maximum_value - chart.value_min) / band_step
        )))
        default_colors = ('156082', 'E97132', '196B24', '0F9ED5')
        colors = chart.surface_band_colors or default_colors

        def intersect(first, second, threshold):
            delta = second[2] - first[2]
            ratio = 0.0 if abs(delta) < 1e-12 else (
                threshold - first[2]
            ) / delta
            return (
                first[0] + (second[0] - first[0]) * ratio,
                first[1] + (second[1] - first[1]) * ratio,
                threshold,
            )

        def band_index(value):
            return max(0, min(
                band_count - 1,
                int((value - chart.value_min) / band_step),
            ))

        def draw_colored_edge(first, second):
            points = [first]
            low_value, high_value = sorted((first[2], second[2]))
            for index in range(1, band_count):
                threshold = chart.value_min + index * band_step
                if low_value < threshold < high_value:
                    points.append(intersect(first, second, threshold))
            points.append(second)
            points.sort(key=lambda point: (
                (point[0] - first[0]) ** 2 + (point[1] - first[1]) ** 2
            ))
            for start, end in zip(points, points[1:]):
                color_index = band_index((start[2] + end[2]) / 2.0)
                canvas.drawLine(
                    start[0], start[1], end[0], end[1],
                    self._paint(
                        'FF' + colors[color_index % len(colors)],
                        stroke_width=0.75,
                    ),
                )

        def clip_polygon(points, threshold, keep_above):
            result = []
            for first, second in zip(points, points[1:] + points[:1]):
                first_inside = (
                    first[2] >= threshold if keep_above
                    else first[2] <= threshold
                )
                second_inside = (
                    second[2] >= threshold if keep_above
                    else second[2] <= threshold
                )
                if first_inside:
                    result.append(first)
                if first_inside != second_inside:
                    result.append(intersect(first, second, threshold))
            return result

        border_paint = self._paint('FFD9D9D9', stroke_width=0.75)
        canvas.drawRect(
            self._skia.Rect.MakeLTRB(
                plot_left, plot_top, plot_right, plot_bottom
            ),
            border_paint,
        )
        canvas.save()
        canvas.clipRect(self._skia.Rect.MakeLTRB(
            plot_left - 0.75, plot_top - 0.75,
            plot_right + 0.75, plot_bottom + 0.75,
        ))
        filled_paths = {}
        for row in range(len(rows) - 1):
            for column in range(point_count - 1):
                lower_left = rows[row][column]
                lower_right = rows[row][column + 1]
                upper_right = rows[row + 1][column + 1]
                upper_left = rows[row + 1][column]
                use_forward_diagonal = (
                    column % 2 == 0
                    and not (
                        column == point_count - 2 and row > 0
                    )
                )
                triangles = (
                    (
                        (lower_left, lower_right, upper_right),
                        (lower_left, upper_right, upper_left),
                    )
                    if use_forward_diagonal else
                    (
                        (lower_left, lower_right, upper_left),
                        (lower_right, upper_right, upper_left),
                    )
                )
                if not chart.surface_wireframe:
                    for band in range(band_count):
                        low = chart.value_min + band * band_step
                        high = low + band_step
                        for triangle in triangles:
                            polygon = clip_polygon(list(triangle), low, True)
                            if polygon:
                                polygon = clip_polygon(
                                    polygon, high, False
                                )
                            if len(polygon) < 3:
                                continue
                            path = self._skia.Path()
                            path.moveTo(polygon[0][0], polygon[0][1])
                            for point in polygon[1:]:
                                path.lineTo(point[0], point[1])
                            path.close()
                            key = (column, band)
                            existing = filled_paths.get(key)
                            filled_paths[key] = (
                                path if existing is None else
                                self._skia.Op(
                                    existing, path,
                                    self._skia.PathOp.kUnion_PathOp,
                                )
                            )
                    continue
                for index in range(1, band_count):
                    threshold = chart.value_min + index * band_step
                    paint = self._paint(
                        'FF' + colors[(index - 1) % len(colors)],
                        stroke_width=0.75,
                    )
                    for triangle in triangles:
                        crossings = []
                        edges = tuple(zip(
                            triangle, triangle[1:] + triangle[:1]
                        ))
                        for first, second in edges:
                            if ((first[2] < threshold <= second[2]) or
                                    (second[2] < threshold <= first[2])):
                                point = intersect(first, second, threshold)
                                if not any(
                                    abs(point[0] - prior[0]) < 1e-7
                                    and abs(point[1] - prior[1]) < 1e-7
                                    for prior in crossings
                                ):
                                    crossings.append(point)
                        if len(crossings) == 2:
                            canvas.drawLine(
                                crossings[0][0], crossings[0][1],
                                crossings[1][0], crossings[1][1], paint,
                            )
        if not chart.surface_wireframe:
            for (column, band), path in sorted(filled_paths.items()):
                shade = 0.675 if column % 2 == 0 else 0.77
                face_color = self._scale_chart_color(
                    colors[band % len(colors)], shade
                )
                canvas.drawPath(
                    path, self._paint('FF' + face_color, fill=True)
                )
        if chart.surface_wireframe:
            for row_index, row in enumerate(rows):
                if row_index == len(rows) - 1:
                    continue
                for first, second in zip(row, row[1:]):
                    draw_colored_edge(first, second)
            for column in range(point_count):
                if column == 0:
                    continue
                for row in range(len(rows) - 1):
                    draw_colored_edge(
                        rows[row][column], rows[row + 1][column]
                    )
        canvas.restore()

        metrics = axis_font.getMetrics()
        label_baseline = plot_bottom + 5.5 - metrics.fAscent
        categories = chart.categories[:point_count]
        for index, label in enumerate(categories):
            text = str(label)
            width = axis_font.measureText(text)
            x = plot_left + index * x_step
            canvas.drawString(
                text, x - width / 2.0, label_baseline,
                axis_font, category_text_paint,
            )
        for row, series in enumerate(visible):
            y = plot_bottom - row * y_step
            baseline = y - (metrics.fAscent + metrics.fDescent) / 2.0
            canvas.drawString(
                series.name, plot_right + 6.0, baseline,
                axis_font, category_text_paint,
            )

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font, title_paint,
        )

        if not chart.show_legend:
            return
        separator = ' -' if chart.surface_wireframe else '-'
        labels = tuple(
            f'{self._chart_value_label(chart.value_min + index * band_step, chart.axis_number_format)}'
            f'{separator}{self._chart_value_label(chart.value_min + (index + 1) * band_step, chart.axis_number_format)}'
            for index in range(band_count)
        )
        item_spacing = 11.5 if chart.surface_wireframe else 14.0
        widths = tuple(
            item_spacing + axis_font.measureText(label) for label in labels
        )
        legend_x = chart.x + (chart.width - sum(widths)) / 2.0
        if not chart.surface_wireframe:
            legend_x += 4.0
        legend_y = chart.y + chart.height - 18.5
        legend_baseline = legend_y - metrics.fAscent - 0.5
        for index, (label, width) in enumerate(zip(labels, widths)):
            color = colors[index % len(colors)]
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    legend_x, legend_y + 1.2, 6.0, 6.0
                ),
                self._paint(
                    'FF' + color,
                    fill=not chart.surface_wireframe,
                    stroke_width=0.75,
                ),
            )
            canvas.drawString(
                label, legend_x + 8.5, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += width

    def _draw_2d_histogram_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        visible = next(
            (series for series in chart.series if not series.hidden), None
        )
        if visible is None:
            return
        axis_metrics = axis_font.getMetrics()
        rectangles, labels, left_max = self._histogram_chart_geometry(chart)
        if not rectangles:
            return

        tick_step = 50000.0 if chart.histogram_variant == 'pareto' else 0.5
        tick_count = max(1, int(round(left_max / tick_step)))
        for index in range(tick_count + 1):
            value = tick_step * index
            y = chart.plot_y + chart.plot_height * (1.0 - value / left_max)
            canvas.drawLine(
                chart.plot_x, y,
                chart.plot_x + chart.plot_width, y,
                grid_paint,
            )
            label = f'{value:g}'
            label_width = axis_font.measureText(label)
            baseline = y - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0 - 0.6
            canvas.drawString(
                label, chart.plot_x - 8.0 - label_width, baseline,
                axis_font, value_text_paint,
            )

        bar_paint = self._paint('FF' + visible.color, fill=True)
        for x, y, width, height in rectangles:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(x, y, width, height), bar_paint
            )

        axis_paint = self._paint('FFD9D9D9', stroke_width=0.75)
        plot_bottom = chart.plot_y + chart.plot_height
        canvas.drawLine(
            chart.plot_x, chart.plot_y, chart.plot_x, plot_bottom, axis_paint
        )
        canvas.drawLine(
            chart.plot_x, plot_bottom,
            chart.plot_x + chart.plot_width, plot_bottom,
            axis_paint,
        )

        label_y = plot_bottom + 5.5 - axis_metrics.fAscent
        for label, (x, _, width, _) in zip(labels, rectangles):
            label_width = axis_font.measureText(label)
            canvas.drawString(
                label, x + (width - label_width) / 2.0, label_y,
                axis_font, category_text_paint,
            )

        if chart.histogram_variant == 'pareto':
            right_x = chart.plot_x + chart.plot_width
            canvas.drawLine(
                right_x, chart.plot_y, right_x, plot_bottom, axis_paint
            )
            for index in range(11):
                ratio = index / 10.0
                y = chart.plot_y + chart.plot_height * (1.0 - ratio)
                label = f'{index * 10}%'
                baseline = y - (
                    axis_metrics.fAscent + axis_metrics.fDescent
                ) / 2.0 - 0.6
                canvas.drawString(
                    label, right_x + 6.5, baseline,
                    axis_font, value_text_paint,
                )
            total = sum(float(value) for value in visible.values)
            cumulative = 0.0
            path = self._skia.Path()
            ordered_values = sorted(
                (float(value) for value in visible.values), reverse=True
            )
            for index, value in enumerate(ordered_values):
                cumulative += value
                px = chart.plot_x + chart.plot_width * (
                    index + 0.5
                ) / len(ordered_values)
                py = plot_bottom - chart.plot_height * (
                    cumulative / total if total else 0.0
                )
                if index == 0:
                    path.moveTo(px, py)
                else:
                    path.lineTo(px, py)
            canvas.drawPath(
                path, self._paint('FFE97132', stroke_width=2.25)
            )

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        title_lift = 2.05 if chart.histogram_variant == 'pareto' else 0.56
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent - title_lift,
            title_font, title_paint,
        )

        if chart.show_legend:
            legend_x = chart.x + 209.46
            legend_y = chart.y + 38.34
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(legend_x, legend_y, 7.08, 7.08),
                bar_paint,
            )
            canvas.drawString(
                visible.name, legend_x + 10.87,
                legend_y - axis_metrics.fAscent - 2.02,
                axis_font, legend_text_paint,
            )

    @staticmethod
    def _histogram_chart_geometry(chart):
        visible = next(
            (series for series in chart.series if not series.hidden), None
        )
        if visible is None:
            return (), (), 1.0

        def snap(value):
            return round(value / 0.12) * 0.12

        if chart.histogram_variant == 'pareto':
            ordered = sorted(
                zip(visible.values, chart.categories), reverse=True
            )
            values = tuple(float(value) for value, _ in ordered)
            labels = tuple(str(label) for _, label in ordered)
            left_max = max(50000.0, math.ceil(max(values) / 50000.0) * 50000.0)
            slot = chart.plot_width / len(values)
            left = chart.plot_x - 0.54
            baseline = chart.plot_y + chart.plot_height - 0.06
            rectangles = []
            for index, value in enumerate(values):
                if len(values) == 6:
                    offsets = (0.0, 49.8, 100.2, 150.0, 199.8, 249.6)
                    widths = (49.68, 50.28, 49.68, 49.68, 49.68, 49.68)
                    x = left + offsets[index]
                    width = widths[index]
                else:
                    x = left + slot * index
                    width = max(0.0, slot - 0.12)
                top = snap(
                    baseline - chart.plot_height * value / left_max + 0.12
                )
                rectangles.append((x, top, width, baseline - top))
            return tuple(rectangles), labels, left_max

        bins = chart.histogram_bins
        if not bins:
            return (), (), 1.0
        frequencies = tuple(float(item[1]) for item in bins)
        left_max = math.ceil(max(frequencies) * 2.0) / 2.0 + 0.5
        slot = chart.plot_width / len(bins)
        baseline = chart.plot_y + chart.plot_height - 0.06
        rectangles = []
        for index, frequency in enumerate(frequencies):
            if len(bins) == 2:
                left_offsets = (26.58, 242.58)
                right_offsets = (188.46, 403.86)
                left = chart.plot_x + left_offsets[index]
                right = chart.plot_x + right_offsets[index]
            else:
                center = chart.plot_x + slot * (index + 0.5)
                half_width = slot * 0.375
                left = snap(center - half_width)
                right = snap(center + half_width)
            top = snap(
                baseline - chart.plot_height * frequency / left_max
                + (0.12 if len(bins) == 2 and index == 0 else 0.0)
            )
            rectangles.append((left, top, right - left, baseline - top))
        return tuple(rectangles), tuple(item[0] for item in bins), left_max

    def _draw_2d_funnel_chart(
        self, canvas, chart, axis_font, title_font,
        category_text_paint, legend_text_paint, title_paint, grid_paint,
    ):
        visible_series = tuple(
            series for series in chart.series if not series.hidden
        )
        if not visible_series:
            return
        series = visible_series[0]
        rectangles = self._funnel_chart_rectangles(chart)
        if not rectangles:
            return

        axis_metrics = axis_font.getMetrics()
        bar_paint = self._paint('FF' + series.color, fill=True)
        for x, y, width, height in rectangles:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(x, y, width, height),
                bar_paint,
            )

        canvas.drawLine(
            chart.plot_x, chart.plot_y,
            chart.plot_x, chart.plot_y + chart.plot_height,
            grid_paint,
        )
        category_y_offset = -(
            axis_metrics.fAscent + axis_metrics.fDescent
        ) / 2.0 - 0.6
        for label, (_, y, _, height) in zip(chart.categories, rectangles):
            center_y = y + height / 2.0
            label_width = axis_font.measureText(label)
            canvas.drawString(
                label,
                chart.plot_x - 6.5 - label_width,
                center_y + category_y_offset + 0.58,
                axis_font,
                category_text_paint,
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

        if chart.show_legend:
            legend_x = (
                chart.x
                + (chart.width - 7.08 - 3.79
                   - axis_font.measureText(series.name)) / 2.0
                + 2.34
            )
            legend_y = chart.y + 38.34
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(legend_x, legend_y, 7.08, 7.08),
                bar_paint,
            )
            canvas.drawString(
                series.name,
                legend_x + 10.87,
                legend_y - axis_metrics.fAscent - 2.02,
                axis_font,
                legend_text_paint,
            )

    @staticmethod
    def _funnel_chart_rectangles(chart):
        visible_series = tuple(
            series for series in chart.series if not series.hidden
        )
        if not visible_series:
            return ()
        values = tuple(float(value) for value in visible_series[0].values)
        if not values:
            return ()
        maximum = max((abs(value) for value in values), default=0.0)
        if maximum <= 0.0:
            return ()

        count = len(values)
        bar_height = chart.plot_height * 0.12435530085959885
        center_top = chart.plot_y + 17.4
        center_step = (
            (chart.plot_height - 34.8) / max(1, count - 1)
            if count > 1 else 0.0
        )
        center_x = chart.plot_x + 0.06 + chart.plot_width / 2.0

        def snap(value):
            return round(value / 0.12) * 0.12

        rectangles = []
        for index, value in enumerate(values):
            width = chart.plot_width * abs(value) / maximum
            center_y = center_top + center_step * index
            left = snap(center_x - width / 2.0)
            right = snap(center_x + width / 2.0)
            top = snap(center_y - bar_height / 2.0)
            bottom = snap(center_y + bar_height / 2.0)
            rectangles.append((left, top, right - left, bottom - top))
        return tuple(rectangles)

