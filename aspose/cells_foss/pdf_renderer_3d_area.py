"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRenderer3DAreaMixin:
    def _draw_stacked_3d_area_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        dense_stack = len(chart.series) > 3 and len(chart.categories) <= 3
        high_perspective_dense = (
            dense_stack
            and chart.perspective >= 65.0
            and bool(
                chart.back_wall_fill_image
                or chart.side_wall_fill_image
            )
        )
        percent_stack = (
            chart.grouping == 'percentStacked'
            and len(chart.series) == 3
            and len(chart.categories) == 6
        )
        if percent_stack:
            image_x = chart.x + 35.28
            image_y = chart.y + 37.08
            pixel_x = 326.10 / 907.0
            pixel_y = 112.14 / 313.0
        elif dense_stack:
            image_x = chart.x + 57.60
            image_y = chart.y + 37.08
            pixel_x = 297.78 / 829.0
            pixel_y = 112.14 / 313.0
        else:
            image_x = chart.x + 50.64
            image_y = chart.y + 64.80
            pixel_x = 305.94 / 851.0
            pixel_y = 107.34 / 300.0

        def image_point(x, y):
            return image_x + x * pixel_x, image_y + y * pixel_y

        def spaced_width(text, font, spacing):
            return sum(font.measureText(char) for char in text) + (
                max(0, len(text) - 1) * spacing
            )

        def draw_spaced_text(text, x, baseline, font, paint, spacing):
            cursor = x
            for char in text:
                canvas.drawString(char, cursor, baseline, font, paint)
                cursor += font.measureText(char) + spacing

        def project(position, ratio):
            position = max(0.0, min(1.0, float(position)))
            ratio = max(0.0, min(1.0, float(ratio)))
            if percent_stack:
                baselines = (
                    (104.8, 209.4), (218.5, 223.1),
                    (345.2, 238.4), (474.1, 253.9),
                    (637.9, 273.6), (804.5, 293.7),
                )
                tops = (
                    (91.3, 16.6), (213.6, 24.2),
                    (353.7, 29.7), (480.0, 39.7),
                    (651.8, 47.6), (823.5, 58.3),
                )
                scaled = position * (len(baselines) - 1)
                left = min(len(baselines) - 1, int(scaled))
                right = min(len(baselines) - 1, left + 1)
                fraction = scaled - left
                base_x = baselines[left][0] + (
                    baselines[right][0] - baselines[left][0]
                ) * fraction
                base_y = baselines[left][1] + (
                    baselines[right][1] - baselines[left][1]
                ) * fraction
                top_x = tops[left][0] + (
                    tops[right][0] - tops[left][0]
                ) * fraction
                top_y = tops[left][1] + (
                    tops[right][1] - tops[left][1]
                ) * fraction
                return image_point(
                    base_x + (top_x - base_x) * ratio,
                    base_y + (top_y - base_y) * ratio,
                )
            if dense_stack:
                if high_perspective_dense:
                    baseline_x = (125.0, 380.0, 665.0)
                    depth_x = (-32.0, -2.0, 43.0)
                    scaled = position * (len(baseline_x) - 1)
                    left = min(len(baseline_x) - 1, int(scaled))
                    right = min(len(baseline_x) - 1, left + 1)
                    fraction = scaled - left
                    base_x = baseline_x[left] + (
                        baseline_x[right] - baseline_x[left]
                    ) * fraction
                    shift_x = depth_x[left] + (
                        depth_x[right] - depth_x[left]
                    ) * fraction
                    base_y = (184.0, 236.0, 278.0)[left] + (
                        (184.0, 236.0, 278.0)[right]
                        - (184.0, 236.0, 278.0)[left]
                    ) * fraction
                    return image_point(
                        base_x + shift_x * ratio,
                        base_y - (176.0 + 62.0 * position) * ratio,
                    )
                baseline_x = 75.0 + 715.0 * position - 10.0 * position ** 2
                baseline_y = (
                    184.0 + 114.0 * position - 20.0 * position ** 2
                )
                return image_point(
                    baseline_x + 10.0 * position * ratio,
                    baseline_y - (162.0 + 134.0 * position) * ratio,
                )
            base_x = 96.0 + 662.0 * position
            base_y = 202.0 + 86.0 * position
            value_height = 190.0 + 43.0 * position
            return image_point(
                base_x + (-14.0 + 33.0 * position) * ratio,
                base_y - value_height * ratio,
            )

        def value_axis_point(ratio):
            if percent_stack:
                return image_point(
                    88.0 - 16.7 * ratio,
                    217.0 - 198.0 * ratio,
                )
            if dense_stack:
                if high_perspective_dense:
                    return image_point(
                        60.0 - 40.0 * ratio,
                        181.0 - 139.0 * ratio,
                    )
                return image_point(
                    32.0 - 15.0 * ratio,
                    214.0 - 179.0 * ratio,
                )
            return image_point(
                78.0 - 17.0 * ratio,
                208.0 - 190.0 * ratio,
            )

        tick_count = max(1, int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        )))
        stacked_grid_paint = self._paint(
            'FF' + chart.value_major_gridline_color, stroke_width=0.35
        )
        if high_perspective_dense:
            side_wall = tuple(image_point(x, y) for x, y in (
                (60.0, 181.0), (20.0, 42.0),
                (200.0, 50.0), (220.0, 168.0), (120.0, 168.0),
            ))
            back_wall = tuple(image_point(x, y) for x, y in (
                (220.0, 168.0), (200.0, 50.0),
                (647.0, 41.0), (627.0, 162.0),
            ))
            if chart.back_wall_fill_image:
                canvas.drawPath(
                    self._polygon_path(back_wall),
                    self._chart_bitmap_paint(
                        chart.back_wall_fill_image, 0.36
                    ),
                )
            if chart.side_wall_fill_image:
                side_wall_path = self._polygon_path(side_wall)
                canvas.drawPath(
                    side_wall_path,
                    self._chart_bitmap_paint(
                        chart.side_wall_fill_image, 0.36
                    ),
                )
                canvas.drawPath(
                    side_wall_path, self._paint('5A000000', fill=True)
                )
        for index in range(tick_count + 1):
            ratio = index / tick_count
            near = value_axis_point(ratio)
            if percent_stack:
                corner = image_point(
                    109.0 + 24.0 * ratio,
                    209.0 - 206.0 * ratio,
                )
                far = image_point(
                    803.0 + 28.0 * ratio,
                    308.0 - 268.0 * ratio,
                )
            elif dense_stack:
                if high_perspective_dense:
                    corner = image_point(
                        220.0 - 20.0 * ratio,
                        168.0 - 118.0 * ratio,
                    )
                    far = image_point(
                        627.0 + 20.0 * ratio,
                        162.0 - 121.0 * ratio,
                    )
                else:
                    corner = image_point(
                        178.0 - 27.0 * ratio,
                        154.0 - 151.0 * ratio,
                    )
                    far = image_point(
                        795.0 + 16.0 * ratio,
                        221.0 - 182.0 * ratio,
                    )
            else:
                corner = image_point(130.0, 174.0 - 170.0 * ratio)
                far = image_point(
                    768.0 + 13.0 * ratio,
                    235.0 - 196.0 * ratio,
                )
            if (not dense_stack and not percent_stack) or index > 0:
                path = self._skia.Path()
                path.moveTo(*near)
                path.lineTo(*corner)
                path.lineTo(*far)
                canvas.drawPath(path, stacked_grid_paint)

            value = chart.value_min + chart.value_step * index
            label = self._chart_value_label(
                value, chart.axis_number_format
            )
            label_width = axis_font.measureText(label)
            baseline = near[1] - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            canvas.drawString(
                label, near[0] - 9.5 - label_width, baseline,
                axis_font, value_text_paint,
            )

        positions = tuple(chart.category_positions)
        baseline = tuple(project(position, 0.0) for position in positions)
        axis_bottom = value_axis_point(0.0)
        if percent_stack:
            shifted_baseline = tuple(image_point(x, y) for x, y in (
                (84.7, 217.8), (205.3, 233.7),
                (335.8, 249.5), (478.0, 267.1),
                (632.8, 287.6), (802.6, 308.5),
            ))
        elif dense_stack:
            dense_floor_y = (
                (193.0, 262.0, 306.0)
                if not high_perspective_dense else
                (190.0, 246.0, 287.0)
            )
            dense_floor_x = (
                (75.0, 430.0, 780.0)
                if not high_perspective_dense else
                (125.0, 380.0, 665.0)
            )
            shifted_baseline = tuple(
                image_point(
                    dense_floor_x[index],
                    dense_floor_y[index],
                )
                for index in range(len(positions))
            )
        else:
            shifted_baseline = tuple(
                (x - 4.2, y + 2.7) for x, y in baseline
            )
        floor = self._skia.Path()
        floor.moveTo(*axis_bottom)
        for point in baseline:
            floor.lineTo(*point)
        for point in reversed(shifted_baseline):
            floor.lineTo(*point)
        floor.close()
        if not dense_stack and not percent_stack:
            canvas.drawPath(floor, self._paint('FFC7C7C7', fill=True))

        category_paint = self._paint(
            'FFD9D9D9',
            stroke_width=0.75 if dense_stack or percent_stack else 0.4,
        )
        category_axis = self._skia.Path()
        if percent_stack:
            dense_category_axis_points = shifted_baseline
            category_axis.moveTo(*dense_category_axis_points[0])
            category_axis.lineTo(*baseline[0])
            category_axis.moveTo(*dense_category_axis_points[0])
            for point in dense_category_axis_points[1:]:
                category_axis.lineTo(*point)
        elif dense_stack:
            if high_perspective_dense:
                dense_category_axis_points = (
                    image_point(60.0, 181.0),
                    image_point(380.0, 236.0),
                    image_point(704.0, 288.0),
                )
            else:
                dense_category_axis_points = (
                    image_point(34.0, 214.4),
                    image_point(365.5, 256.0),
                    image_point(780.0, 308.1),
                )
            category_axis.moveTo(*dense_category_axis_points[0])
            category_axis.lineTo(*shifted_baseline[0])
            category_axis.moveTo(*dense_category_axis_points[0])
            category_axis.lineTo(*dense_category_axis_points[-1])
        else:
            dense_category_axis_points = ()
            category_axis.moveTo(*axis_bottom)
            for point in baseline:
                category_axis.lineTo(*point)
        canvas.drawPath(
            category_axis,
            category_paint,
        )

        totals = [
            sum(float(series.values[index]) for series in chart.series)
            for index in range(len(positions))
        ]
        dense_boundary_y = (
            (184.0, 170.0, 157.0, 126.0, 96.0, 63.0, 26.0),
            (236.0, 221.0, 215.0, 189.0, 167.0, 142.0, 100.0),
            (278.0, 265.0, 256.0, 239.0, 226.0, 215.0, 200.0),
        )
        high_perspective_boundary_y = (
            (168.0, 157.0, 149.0, 140.0, 117.0, 101.0, 83.0),
            (207.0, 199.0, 195.0, 180.0, 167.0, 152.0, 124.0),
            (250.0, 241.0, 234.0, 222.0, 213.0, 204.0, 195.0),
        )
        percent_boundaries = (
            (
                (104.8, 209.4), (218.5, 223.1), (345.2, 238.4),
                (474.1, 253.9), (637.9, 273.6), (804.5, 293.7),
            ),
            (
                (98.0, 112.0), (216.0, 122.0), (349.0, 148.0),
                (477.0, 151.0), (645.0, 161.0), (812.0, 203.0),
            ),
            (
                (94.0, 55.0), (214.0, 82.0), (352.0, 74.0),
                (479.0, 77.0), (650.0, 78.0), (821.0, 90.0),
            ),
            (
                (91.3, 16.6), (213.6, 24.2), (353.7, 29.7),
                (480.0, 39.7), (651.8, 47.6), (823.5, 58.3),
            ),
        )

        def dense_boundary_point(category_index, boundary_index, ratio):
            position = positions[category_index]
            if high_perspective_dense:
                baseline_x = (125.0, 380.0, 665.0)[category_index]
                depth_x = (-32.0, -2.0, 43.0)[category_index]
                boundary_y = high_perspective_boundary_y
                return image_point(
                    baseline_x + depth_x * ratio,
                    boundary_y[category_index][boundary_index],
                )
            baseline_x = (75.0, 430.0, 780.0)[category_index]
            return image_point(
                baseline_x + (-11.0 + 21.0 * position) * ratio,
                dense_boundary_y[category_index][boundary_index],
            )

        cumulative = [0.0] * len(positions)
        bands = []
        for series_index, series in enumerate(chart.series):
            lower = list(cumulative)
            for index, value in enumerate(series.values[:len(positions)]):
                cumulative[index] += float(value)
            upper = list(cumulative)
            if chart.grouping == 'percentStacked':
                lower_ratios = [
                    value / totals[index] if totals[index] else 0.0
                    for index, value in enumerate(lower)
                ]
                upper_ratios = [
                    value / totals[index] if totals[index] else 0.0
                    for index, value in enumerate(upper)
                ]
            else:
                span = max(1e-12, chart.value_max - chart.value_min)
                lower_ratios = [
                    (value - chart.value_min) / span for value in lower
                ]
                upper_ratios = [
                    (value - chart.value_min) / span for value in upper
                ]
            if percent_stack:
                lower_points = tuple(
                    image_point(*point)
                    for point in percent_boundaries[series_index]
                )
                upper_points = tuple(
                    image_point(*point)
                    for point in percent_boundaries[series_index + 1]
                )
            elif dense_stack:
                lower_points = tuple(
                    dense_boundary_point(index, series_index, ratio)
                    for index, ratio in enumerate(lower_ratios)
                )
                upper_points = tuple(
                    dense_boundary_point(index, series_index + 1, ratio)
                    for index, ratio in enumerate(upper_ratios)
                )
            else:
                lower_points = tuple(
                    project(position, ratio)
                    for position, ratio in zip(positions, lower_ratios)
                )
                upper_points = tuple(
                    project(position, ratio)
                    for position, ratio in zip(positions, upper_ratios)
                )
            bands.append((
                series, lower_points, upper_points,
                lower_ratios, upper_ratios,
            ))

        top_series, _, top_points, _, _ = bands[-1]
        def top_extrusion_x(position):
            return 4.4 - 3.25 * position

        extrusion_y = -3.5
        if percent_stack:
            shifted_top = tuple(
                image_point(
                    132.0 + (827.0 - 132.0) * position,
                    8.0 + (49.0 - 8.0) * position,
                )
                for position in positions
            )
        elif dense_stack:
            dense_top = (
                ((99.0, 18.0), (431.0, 82.0), (789.0, 190.0))
                if not high_perspective_dense else
                ((130.0, 79.0), (380.0, 109.0), (645.0, 166.0))
            )
            shifted_top = tuple(
                image_point(x, y) for x, y in dense_top
            )
        else:
            shifted_top = tuple(
                (
                    x + top_extrusion_x(position),
                    y + extrusion_y,
                )
                for (x, y), position in zip(top_points, positions)
            )
        top_path = self._skia.Path()
        top_path.moveTo(*shifted_top[0])
        for point in shifted_top[1:]:
            top_path.lineTo(*point)
        for point in reversed(top_points):
            top_path.lineTo(*point)
        top_path.close()
        if (
            top_series.pattern_type in self._CHART_PATTERN_ROWS
            and top_series.pattern_foreground
            and top_series.pattern_background
        ):
            top_paint = self._chart_pattern_paint(
                top_series.pattern_type,
                self._scale_chart_color(
                    top_series.pattern_foreground, 0.64
                ),
                self._scale_chart_color(
                    top_series.pattern_background, 0.78
                ),
            )
        else:
            top_paint = self._paint(
                'FF' + self._scale_chart_color(top_series.color, 0.78),
                fill=True,
            )
        canvas.drawPath(top_path, top_paint)

        if percent_stack:
            side_back_points = tuple(image_point(x, y) for x, y in (
                (806.0, 290.0),
                (816.0, 190.0),
                (823.0, 87.0),
                (827.0, 49.0),
            ))
            for series_index, (
                series, lower_points, upper_points,
                _, _,
            ) in enumerate(bands):
                side = self._skia.Path()
                side.moveTo(*upper_points[-1])
                side.lineTo(*side_back_points[series_index + 1])
                side.lineTo(*side_back_points[series_index])
                side.lineTo(*lower_points[-1])
                side.close()
                canvas.drawPath(
                    side,
                    self._paint(
                        'FF' + self._scale_chart_color(series.color, 0.62),
                        fill=True,
                    ),
                )

        for series, lower_points, upper_points, _, _ in bands:
            path = self._skia.Path()
            path.moveTo(*upper_points[0])
            for point in upper_points[1:]:
                path.lineTo(*point)
            for point in reversed(lower_points):
                path.lineTo(*point)
            path.close()
            if (
                series.pattern_type in self._CHART_PATTERN_ROWS
                and series.pattern_foreground
                and series.pattern_background
            ):
                paint = self._chart_pattern_paint(
                    series.pattern_type,
                    series.pattern_foreground,
                    series.pattern_background,
                )
            else:
                paint = self._paint('FF' + series.color, fill=True)
            canvas.drawPath(path, paint)

        if dense_stack and not high_perspective_dense:
            top_side_dx = shifted_top[-1][0] - top_points[-1][0]
            top_side_dy = shifted_top[-1][1] - top_points[-1][1]
            base_side_dx = 1.0 * pixel_x
            base_side_dy = -1.0 * pixel_y
            top_ratio = max(1e-12, bands[-1][4][-1])

            def side_offset(ratio):
                scale = max(0.0, min(1.0, ratio / top_ratio))
                return (
                    base_side_dx + (top_side_dx - base_side_dx) * scale,
                    base_side_dy + (top_side_dy - base_side_dy) * scale,
                )

            for (
                series, lower_points, upper_points,
                lower_ratios, upper_ratios,
            ) in bands:
                lower = lower_points[-1]
                upper = upper_points[-1]
                lower_dx, lower_dy = side_offset(lower_ratios[-1])
                upper_dx, upper_dy = side_offset(upper_ratios[-1])
                side = self._skia.Path()
                side.moveTo(*upper)
                side.lineTo(upper[0] + upper_dx, upper[1] + upper_dy)
                side.lineTo(lower[0] + lower_dx, lower[1] + lower_dy)
                side.lineTo(*lower)
                side.close()
                canvas.drawPath(
                    side,
                    self._paint(
                        'FF' + self._scale_chart_color(series.color, 0.62),
                        fill=True,
                    ),
                )

        if not dense_stack and not percent_stack:
            value_axis = self._skia.Path()
            value_axis.moveTo(*value_axis_point(0.0))
            value_axis.lineTo(*value_axis_point(1.0))
            for index in range(tick_count + 1):
                ratio = index / tick_count
                x, y = value_axis_point(ratio)
                value_axis.moveTo(x - 2.6, y)
                value_axis.lineTo(x + 2.6, y)
            canvas.drawPath(
                value_axis,
                self._paint('FFD9D9D9', stroke_width=0.4),
            )

        for category_index, (label, position) in enumerate(
            chart.axis_categories
        ):
            if percent_stack:
                label_centers = (
                    65.77, 109.13, 156.12,
                    207.17, 262.88, 323.94,
                )
                label_tops = (
                    122.18, 127.75, 133.70,
                    140.04, 147.21, 154.73,
                )
                x = chart.x + label_centers[category_index]
                y = chart.y + label_tops[category_index]
                tick_x, tick_y = dense_category_axis_points[category_index]
                canvas.drawLine(
                    tick_x, tick_y, tick_x, tick_y + 2.8, category_paint
                )
            elif dense_stack:
                if high_perspective_dense:
                    label_centers = (79.72, 154.78, 311.20)
                    label_tops = (111.24, 125.64, 155.47)
                else:
                    label_centers = (69.81, 188.98, 337.86)
                    label_tops = (121.70, 136.65, 155.37)
                x = chart.x + label_centers[category_index]
                y = chart.y + label_tops[category_index]
                tick_x, tick_y = dense_category_axis_points[category_index]
                canvas.drawLine(
                    tick_x, tick_y, tick_x, tick_y + 2.8, category_paint
                )
            else:
                x, y = project(position, 0.0)
                label_offset_y = 11.0
            if not dense_stack and not percent_stack:
                canvas.drawLine(x, y, x, y + 3.0, category_paint)
            if dense_stack or percent_stack:
                text = str(label)
                label_width = axis_font.measureText(text)
                canvas.drawString(
                    text, x - label_width / 2.0,
                    y - axis_metrics.fAscent,
                    axis_font, category_text_paint,
                )
            else:
                text = str(label).upper()
                label_width = spaced_width(text, axis_font, 1.2)
                draw_spaced_text(
                    text, x - label_width / 2.0,
                    y + label_offset_y - axis_metrics.fAscent,
                    axis_font, self._paint('FF808080', fill=True), 1.2,
                )

        title_metrics = title_font.getMetrics()
        if dense_stack or percent_stack:
            dense_title_font = self._font(
                chart.title_font_resolution,
                getattr(chart.title_font_resolution, 'size_points', 14.0),
            )
            dense_title_font.setScaleX(0.972)
            title_width = dense_title_font.measureText(chart.title)
            title_metrics = dense_title_font.getMetrics()
            canvas.drawString(
                chart.title,
                chart.x + (chart.width - title_width) / 2.0,
                chart.y + 7.53 - title_metrics.fAscent,
                dense_title_font, title_paint,
            )
        else:
            title_width = spaced_width(chart.title, title_font, 1.3)
            draw_spaced_text(
                chart.title,
                chart.x + (chart.width - title_width) / 2.0,
                chart.y + 6.75 - title_metrics.fAscent,
                title_font, title_paint, 1.3,
            )

        if not chart.show_legend:
            return
        legend_font = self._font(chart.axis_font_resolution, 9.0)
        legend_font.setScaleX(0.96 if dense_stack or percent_stack else 0.90)
        legend_metrics = legend_font.getMetrics()
        item_gap = 8.0 if percent_stack else 7.0
        item_widths = [
            9.0 + legend_font.measureText(series.name) + item_gap
            for series in chart.series
        ]
        legend_x = chart.x + (chart.width - sum(item_widths)) / 2.0
        if dense_stack:
            legend_x += 5.4
        elif percent_stack:
            legend_x += 5.1
        legend_y = (
            chart.y + chart.height - 20.5
            if dense_stack or percent_stack else chart.y + 40.0
        )
        legend_baseline = legend_y - legend_metrics.fAscent
        for series, item_width in zip(chart.series, item_widths):
            rect = self._skia.Rect.MakeXYWH(
                legend_x, legend_y + 3.4, 5.0, 5.0
            )
            if (
                series.pattern_type in self._CHART_PATTERN_ROWS
                and series.pattern_foreground
                and series.pattern_background
            ):
                legend_paint = self._chart_pattern_paint(
                    series.pattern_type,
                    series.pattern_foreground,
                    series.pattern_background,
                )
            else:
                legend_paint = self._paint('FF' + series.color, fill=True)
            canvas.drawRect(rect, legend_paint)
            canvas.drawString(
                series.name, legend_x + 8.0, legend_baseline,
                legend_font,
                legend_text_paint if dense_stack or percent_stack else
                self._paint('FF808080', fill=True),
            )
            legend_x += item_width

    def _draw_3d_area_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        import math

        if chart.grouping in ('stacked', 'percentStacked'):
            self._draw_stacked_3d_area_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return

        # Excel rasterizes the 3-D plot itself. These normalized projections
        # reproduce its default 15/20-degree camera while keeping PDF text and
        # the surrounding chart frame vector based.
        scale_x = chart.width / 384.36
        scale_y = chart.height / 195.36
        axis_metrics = axis_font.getMetrics()
        dense_depth = len(chart.series) > 3 and len(chart.categories) <= 3
        expanded_date_depth = (
            not dense_depth
            and len(chart.axis_categories) > len(chart.categories)
        )
        if dense_depth and chart.x < 0.0:
            # Excel's continuation image is transparent on the right-hand
            # page, leaving only the chart-area frame drawn by the caller.
            return

        def perspective(position):
            curvature = 0.42
            return (
                math.exp(curvature * position) - 1.0
            ) / (math.exp(curvature) - 1.0)

        def projected_horizontal(position):
            horizontal = perspective(position)
            if dense_depth:
                horizontal += 0.0295 * 4.0 * position * (1.0 - position)
            return horizontal

        def category_calibration(position, depth, rows):
            if not dense_depth:
                return 0.0
            row = rows[min(int(depth), len(rows) - 1)]
            if position <= 0.5:
                weight = max(0.0, position * 2.0)
                return row[0] + (row[1] - row[0]) * weight
            weight = min(1.0, (position - 0.5) * 2.0)
            return row[1] + (row[2] - row[1]) * weight

        category_x_adjustments = (
            (0.0, 0.23, -0.90),
            (0.0, 0.65, -0.50),
            (0.67, 1.10, -0.60),
            (1.34, 0.95, 0.00),
            (2.01, 0.80, 0.00),
            (2.68, 2.70, 0.65),
        )
        category_value_height_adjustments = (
            (0.0, -8.80, -23.60),
            (0.0, 11.50, -22.60),
            (6.55, 0.47, -16.00),
            (10.87, 5.20, -15.80),
            (9.30, 10.00, -13.90),
            (13.98, 1.70, -10.70),
        )
        category_surface_y_adjustments = (
            (0.0, 0.0, 0.0),
            (0.0, 0.0, 0.0),
            (0.0, 0.0, 0.0),
            (0.0, 0.0, 0.0),
            (0.0, 0.0, 0.0),
            (0.0, 0.0, 0.0),
        )

        def wall_point(ratio, section):
            if dense_depth:
                if section == 0:
                    x = 135.55
                    y = 176.15 - 90.85 * ratio - 3.35 * ratio * ratio
                elif section == 1:
                    x = 246.0
                    y = 126.0 - 87.0 * ratio
                else:
                    x = 361.0
                    # Excel gives the back-wall gridlines a subtle downward
                    # tilt toward the right rather than making them level.
                    y = 131.2 - 87.0 * ratio
                return chart.x + x, chart.y + y
            if expanded_date_depth:
                if section == 0:
                    x = 88.0 - 13.0 * ratio
                    y = 178.0 - 132.0 * ratio
                elif section == 1:
                    x = 257.0 - 6.0 * ratio
                    y = 117.0 - 114.0 * ratio
                else:
                    x = 789.0 + 23.0 * ratio
                    y = 155.0 - 123.0 * ratio
                return (
                    chart.x + 10.92 + x * (322.62 / 897.0),
                    chart.y + 37.08 + y * (96.66 / 270.0),
                )
            if section == 0:
                x = 65.69 - 5.2 * ratio
                y = 116.19 - 62.3 * ratio
            elif section == 1:
                x = 123.0 - 1.1 * ratio
                y = 75.8 - 37.3 * ratio
            else:
                x = 324.6 + 2.6 * ratio
                y = 76.9 - 29.1 * ratio
            return (
                chart.x + x * scale_x,
                chart.y + y * scale_y,
            )

        tick_count = int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        ))
        draw_minor_gridlines = (
            dense_depth and chart.show_minor_horizontal_gridlines
        )
        wall_divisions = (
            tick_count * 5 if draw_minor_gridlines else tick_count
        )
        minor_grid_paint = self._paint(
            'FF' + chart.minor_gridline_color, stroke_width=0.75
        )
        if not dense_depth:
            side_wall = (
                wall_point(0.0, 0), wall_point(0.0, 1),
                wall_point(1.0, 1), wall_point(1.0, 0),
            )
            back_wall = (
                wall_point(0.0, 1), wall_point(0.0, 2),
                wall_point(1.0, 2), wall_point(1.0, 1),
            )
            if chart.back_wall_fill_image:
                canvas.drawPath(
                    self._polygon_path(back_wall),
                    self._chart_bitmap_paint(
                        chart.back_wall_fill_image, 0.36
                    ),
                )
            if chart.side_wall_fill_image:
                side_wall_path = self._polygon_path(side_wall)
                canvas.drawPath(
                    side_wall_path,
                    self._chart_bitmap_paint(
                        chart.side_wall_fill_image, 0.36
                    ),
                )
                canvas.drawPath(
                    side_wall_path, self._paint('33000000', fill=True)
                )
        for index in range(wall_divisions + 1):
            ratio = index / max(1, wall_divisions)
            near, corner, far = (
                wall_point(ratio, section) for section in range(3)
            )
            path = self._skia.Path()
            path.moveTo(*near)
            path.lineTo(*corner)
            path.lineTo(*far)
            is_major_gridline = (
                not draw_minor_gridlines or index % 5 == 0
            )
            canvas.drawPath(
                path,
                grid_paint if is_major_gridline else minor_grid_paint,
            )

            if is_major_gridline:
                value = (
                    chart.value_min
                    + chart.value_step * (
                        index // 5 if draw_minor_gridlines else index
                    )
                )
                label = self._chart_value_label(
                    value, chart.axis_number_format
                )
                label_width = axis_font.measureText(label)
                baseline = near[1] - (
                    axis_metrics.fAscent + axis_metrics.fDescent
                ) / 2.0
                if not dense_depth and chart.value_axis_position == 'r':
                    if expanded_date_depth:
                        label_x = far[0] + 11.65 - 4.25 * ratio
                        baseline += -8.34 + 3.42 * ratio
                    else:
                        label_x = far[0] + 8.0
                else:
                    label_x = (
                        near[0]
                        - (3.13 if dense_depth else 10.0)
                        - label_width
                    )
                canvas.drawString(
                    label, label_x, baseline,
                    axis_font, value_text_paint,
                )

        def project(position, depth, ratio):
            horizontal = projected_horizontal(position)
            if dense_depth:
                depth_steps = (0.0, 0.265, 0.49, 0.685, 0.855, 1.0)
                if depth < len(depth_steps):
                    depth_ratio = depth_steps[int(depth)]
                else:
                    depth_ratio = 1.0 + 0.13 * (
                        depth - len(depth_steps) + 1.0
                    )
                # Leave room for the extruded right face before the depth
                # labels while keeping the category axis endpoint unchanged.
                base_x = 146.0 + 154.6 * horizontal
                base_y = 176.2 + 14.3 * horizontal
                value_x = (-1.5 + 6.0 * horizontal) * ratio
                value_height = (
                    98.0 + 20.0 * horizontal
                    - 42.0 * horizontal * depth_ratio
                    - 7.5 * horizontal * (1.0 - horizontal)
                    * depth_ratio
                    + category_calibration(
                        position, depth,
                        category_value_height_adjustments,
                    )
                )
                value_y = -value_height * ratio
                depth_x = (
                    (90.0 - 42.0 * horizontal) * depth_ratio
                    + 1.4 * horizontal * (1.0 - depth_ratio)
                    + (1.2 + 1.75 * depth_ratio) * horizontal
                    + category_calibration(
                        position, depth, category_x_adjustments
                    )
                )
                depth_y = -(
                    30.0 + 42.0 * horizontal
                ) * depth_ratio
                return (
                    chart.x + base_x + value_x + depth_x,
                    chart.y + base_y + value_y + depth_y,
                )
            if expanded_date_depth:
                base_x = 98.2 + 718.4 * horizontal
                base_y = 166.5 + 85.2 * horizontal
                left_front_weight = (
                    (1.0 - horizontal) ** 4 if depth == 0 else 0.0
                )
                right_front_weight = horizontal ** 4 if depth == 0 else 0.0
                # Excel keeps the exposed front-left edge almost vertical.
                # Pull its floor point toward the side wall, then offset the
                # value vector so the already-calibrated top point stays put.
                base_x -= 10.2 * left_front_weight
                base_x += 2.5 * right_front_weight
                value_x = (
                    -40.0 + 54.0 * horizontal
                    + 35.0 * left_front_weight
                ) * ratio
                front_height = 0.0
                if depth == 0:
                    front_height = (
                        7.0
                        + 10.0 * (1.0 - left_front_weight)
                        * (1.0 - 0.4 * horizontal)
                    )
                value_y = -(
                    111.0 + 47.0 * horizontal + front_height
                ) * ratio
                depth_fraction = min(1.0, max(
                    0.0,
                    depth / max(1.0, float(len(chart.series) - 1)),
                ))
                middle_depth = (
                    4.0 * depth_fraction * (1.0 - depth_fraction)
                )
                back_depth = (
                    max(0.0, 2.0 * depth_fraction - 1.0)
                    if depth <= float(len(chart.series) - 1)
                    else 0.0
                )
                # The depth axis converges sharply toward the right edge in
                # Excel's reversed date-axis projection. A nearly uniform
                # offset makes the rear series protrude at both endpoints.
                depth_x = (
                    (56.5 - 53.5 * horizontal) * depth
                    + (2.0 - 10.0 * horizontal) * back_depth
                )
                depth_y = (
                    -(23.3 + 8.3 * horizontal) * depth
                    - 7.0 * middle_depth
                    - 13.0 * (
                        1.0 - (1.0 - horizontal) ** 4
                    ) * back_depth
                )
                return (
                    chart.x + 10.92
                    + (base_x + value_x + depth_x) * (322.62 / 897.0),
                    chart.y + 37.08
                    + (base_y + value_y + depth_y) * (96.66 / 270.0),
                )
            base_x = 71.9 + 241.6 * horizontal
            base_y = 113.7 + 34.1 * horizontal
            value_x = (-3.2 + 10.2 * horizontal) * ratio
            value_y = -(63.0 + 19.0 * horizontal) * ratio
            depth_x = (13.0 - 7.0 * horizontal) * depth
            depth_y = -(10.5 + 3.5 * horizontal) * depth
            return (
                chart.x + (base_x + value_x + depth_x) * scale_x,
                chart.y + (base_y + value_y + depth_y) * scale_y,
            )

        def project_baseline(position, depth):
            if not dense_depth:
                return project(position, depth, 0.0)
            horizontal = projected_horizontal(position)
            depth_steps = (0.0, 0.265, 0.49, 0.685, 0.855, 1.0)
            if depth < len(depth_steps):
                depth_ratio = depth_steps[int(depth)]
            else:
                depth_ratio = 1.0 + 0.13 * (
                    depth - len(depth_steps) + 1.0
                )
            return (
                chart.x
                + 146.0 + 154.6 * horizontal
                + (90.0 - 42.0 * horizontal) * depth_ratio
                + 1.4 * horizontal * (1.0 - depth_ratio)
                + (1.2 + 1.75 * depth_ratio) * horizontal
                + category_calibration(
                    position, depth, category_x_adjustments
                ),
                chart.y
                + 169.17 + 21.35 * horizontal
                - 71.4 * depth_ratio,
            )

        def category_axis_point(position):
            horizontal = projected_horizontal(position)
            if dense_depth:
                return (
                    chart.x + 139.69 + 159.07 * horizontal,
                    chart.y + 176.17 + 21.35 * horizontal,
                )
            if expanded_date_depth:
                axis_start = chart.axis_categories[0][1]
                axis_end = chart.axis_categories[-1][1]
                axis_span = max(1e-12, axis_end - axis_start)
                axis_ratio = (position - axis_start) / axis_span
                bow = 4.0 * axis_ratio * (1.0 - axis_ratio)
                return (
                    chart.x + 38.086 + 269.869 * horizontal - 3.0 * bow,
                    chart.y + 98.597 + 33.612 * horizontal - 0.4 * bow,
                )
            return (
                chart.x + (65.69 + 247.18 * horizontal) * scale_x,
                chart.y + (116.19 + 31.29 * horizontal) * scale_y,
            )

        if not dense_depth and chart.show_vertical_gridlines:
            far_depth = max(3.4, float(len(chart.series)))
            for _, position in chart.axis_categories:
                front = category_axis_point(position)
                back = project(position, far_depth, 0.0)
                canvas.drawLine(*front, *back, grid_paint)

        for series_index in reversed(range(len(chart.series))):
            series = chart.series[series_index]
            top_points = []
            base_points = []
            series_points = sorted(zip(
                chart.category_positions, series.values
            ))
            for point_index, (position, value) in enumerate(series_points):
                ratio = self._chart_value_ratio(chart, float(value))
                top_point = project(position, series_index, ratio)
                base_point = project_baseline(position, series_index)
                if (
                    expanded_date_depth
                    and series_index == len(chart.series) - 1
                    and point_index == 0
                ):
                    # The front sheets occlude the lower part of the rear
                    # sheet at the left boundary in Excel's projection. Its
                    # exposed start edge remains vertical in screen space.
                    base_point = (
                        top_point[0],
                        min(base_point[1], top_point[1] + 10.0 * (96.66 / 270.0)),
                    )
                top_points.append(top_point)
                base_points.append(base_point)
            if not top_points:
                continue
            path = self._skia.Path()
            path.moveTo(*top_points[0])
            for point in top_points[1:]:
                path.lineTo(*point)
            if (
                expanded_date_depth
                and series_index == len(chart.series) - 1
                and len(base_points) > 1
            ):
                for point in reversed(base_points[1:]):
                    path.lineTo(*point)
                path.lineTo(
                    top_points[0][0] + 78.0 * (322.62 / 897.0),
                    top_points[0][1] - 3.0 * (96.66 / 270.0),
                )
                path.lineTo(*base_points[0])
            else:
                for point in reversed(base_points):
                    path.lineTo(*point)
            path.close()

            if dense_depth:
                depth_steps = (0.0, 0.265, 0.49, 0.685, 0.855, 1.0)
                depth_ratio = depth_steps[min(
                    series_index, len(depth_steps) - 1
                )]
                extrusion_x = 4.55 - 1.75 * depth_ratio
            elif expanded_date_depth:
                extrusion_x = 0.8 * scale_x
            else:
                extrusion_x = 4.2 * scale_x
            if dense_depth:
                extrusion_y = -4.5
            elif expanded_date_depth:
                # Keep the sheet thickness parallel to the floor's depth
                # edge. This makes the area plane read as perpendicular to
                # both the side wall and the floor instead of leaning upward.
                floor_near = wall_point(0.0, 0)
                floor_corner = wall_point(0.0, 1)
                floor_depth_slope = (
                    (floor_corner[1] - floor_near[1])
                    / max(1e-12, floor_corner[0] - floor_near[0])
                )
                extrusion_y = extrusion_x * floor_depth_slope
            else:
                extrusion_y = -3.2 * scale_y
            if dense_depth:
                shifted_top = tuple(
                    (
                        x + extrusion_x
                        + 4.05 * (1.0 - projected_horizontal(position)),
                        y - 3.4
                        - 1.1 * projected_horizontal(position)
                        + category_calibration(
                            position, series_index,
                            category_surface_y_adjustments,
                        ),
                    )
                    for (x, y), (position, _) in zip(
                        top_points, series_points
                    )
                )
            elif expanded_date_depth:
                shifted_top = tuple(
                    (
                        x + (
                            6.0 - 5.2 * projected_horizontal(position)
                        ) * scale_x,
                        y + (
                            6.0 - 5.2 * projected_horizontal(position)
                        ) * scale_x * floor_depth_slope,
                    )
                    for (x, y), (position, _) in zip(
                        top_points, series_points
                    )
                )
            else:
                shifted_top = tuple(
                    (x + extrusion_x, y + extrusion_y)
                    for x, y in top_points
                )
            shifted_base_end = (
                base_points[-1][0] + extrusion_x,
                base_points[-1][1] + extrusion_y,
            )

            side_path = self._skia.Path()
            side_path.moveTo(*top_points[-1])
            side_path.lineTo(*base_points[-1])
            side_path.lineTo(*shifted_base_end)
            side_path.lineTo(*shifted_top[-1])
            side_path.close()
            canvas.drawPath(
                side_path,
                self._paint(
                    'FF' + self._scale_chart_color(series.color, 0.64),
                    fill=True,
                ),
            )

            top_path = self._skia.Path()
            top_path.moveTo(*shifted_top[0])
            for point in shifted_top[1:]:
                top_path.lineTo(*point)
            for point in reversed(top_points):
                top_path.lineTo(*point)
            top_path.close()
            top_color = self._scale_chart_color(series.color, 0.80)
            canvas.drawPath(
                top_path,
                self._paint('FF' + top_color, fill=True),
            )
            if expanded_date_depth and series_index > 0:
                canvas.drawPath(
                    top_path,
                    self._paint('FF' + top_color, stroke_width=0.28),
                )

            canvas.drawPath(
                path, self._paint('FF' + series.color, fill=True)
            )
            if expanded_date_depth and series_index > 0:
                # Seal sub-pixel seams where antialiasing would otherwise
                # blend the opaque series edge with the textured wall below.
                canvas.drawPath(
                    path,
                    self._paint(
                        'FF' + series.color, stroke_width=0.28
                    ),
                )

        if chart.value_axis_color:
            axis_paint = self._paint(
                'FF' + chart.value_axis_color, stroke_width=0.75
            )
            if dense_depth:
                def axis_point(ratio):
                    return (
                        chart.x + 139.62 - 4.07 * ratio,
                        wall_point(ratio, 0)[1],
                    )
            else:
                def axis_point(ratio):
                    return wall_point(ratio, 0)
            bottom = axis_point(0.0)
            top = axis_point(1.0)
            axis_path = self._skia.Path()
            axis_path.moveTo(*bottom)
            axis_path.lineTo(*top)
            if chart.value_major_tick_mark == 'cross':
                for index in range(tick_count + 1):
                    ratio = index / max(1, tick_count)
                    x, y = axis_point(ratio)
                    axis_path.moveTo(x - 2.8, y)
                    axis_path.lineTo(x + 2.8, y)
            canvas.drawPath(axis_path, axis_paint)

        category_axis_paint = self._paint(
            'FF' + chart.value_major_gridline_color,
            stroke_width=0.4 if dense_depth else 0.75,
        )
        if chart.axis_categories:
            category_axis_start = category_axis_point(
                chart.axis_categories[0][1]
            )
            category_axis_end = category_axis_point(
                chart.axis_categories[-1][1]
            )
            category_axis_path = self._skia.Path()
            category_axis_path.moveTo(*category_axis_start)
            category_axis_path.lineTo(*category_axis_end)
            canvas.drawPath(category_axis_path, category_axis_paint)

        for label, position in chart.axis_categories:
            x, y = category_axis_point(position)
            canvas.drawLine(x, y, x, y + 2.8, category_axis_paint)
            label_width = axis_font.measureText(label)
            dense_label_x = 0.7 * (1.0 - position) if dense_depth else 0.0
            dense_label_y = (
                0.87 + 0.04 * position if dense_depth else 0.0
            )
            if chart.category_rotation:
                canvas.save()
                canvas.translate(x + dense_label_x, y + 7.1 + dense_label_y)
                canvas.rotate(chart.category_rotation)
                canvas.drawString(
                    label, -label_width, 0.0,
                    axis_font, category_text_paint,
                )
                canvas.restore()
            else:
                canvas.drawString(
                    label, x - label_width / 2.0 + dense_label_x,
                    y + 7.1 + dense_label_y - axis_metrics.fAscent,
                    axis_font, category_text_paint,
                )

        if dense_depth:
            label_x_offsets = (
                314.35, 327.0, 337.85, 347.23, 355.46, 362.18,
            )
            label_top_offsets = (
                181.75, 163.63, 148.1, 134.64, 122.88, 112.47,
            )
        elif expanded_date_depth:
            label_x_offsets = (300.36, 302.11, 303.48)
            label_top_offsets = (118.10, 103.70, 92.02)
        else:
            label_x_offsets = (321.7, 325.27, 328.18)
            label_top_offsets = (133.61, 119.16, 107.09)
        for series_index, series in enumerate(chart.series):
            label_x_offset = (
                label_x_offsets[series_index]
                if series_index < len(label_x_offsets)
                else label_x_offsets[-1] + 3.0 * (
                    series_index - len(label_x_offsets) + 1
                )
            )
            label_top_offset = (
                label_top_offsets[series_index]
                if series_index < len(label_top_offsets)
                else label_top_offsets[-1] - 12.0 * (
                    series_index - len(label_top_offsets) + 1
                )
            )
            label_x = chart.x + label_x_offset * (
                1.0 if dense_depth else scale_x
            )
            label_top = chart.y + label_top_offset * (
                1.0 if dense_depth else scale_y
            )
            if dense_depth:
                label_width = axis_font.measureText(series.name)
                canvas.drawRect(
                    self._skia.Rect.MakeLTRB(
                        label_x - 5.0,
                        label_top - 1.0,
                        label_x + label_width + 2.0,
                        label_top - axis_metrics.fAscent
                        + axis_metrics.fDescent + 1.0,
                    ),
                    self._paint('FFFFFFFF', fill=True),
                )
            canvas.drawString(
                series.name, label_x, label_top - axis_metrics.fAscent,
                axis_font, category_text_paint,
            )

        self._draw_3d_chart_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
            legend_offset_x=-1.9 if expanded_date_depth else -1.1,
            legend_font_scale_x=0.96,
            legend_item_gap=8.0 if expanded_date_depth else 7.0,
            title_font_scale_x=0.972 if expanded_date_depth else 1.0,
        )

