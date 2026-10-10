"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRenderer3DLineColumnMixin:
    def _draw_3d_line_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        import math

        if chart.perspective >= 45.0 and len(chart.series) > 3:
            self._draw_perspective_3d_line_chart(
                canvas, chart, axis_font, title_font,
                value_text_paint, category_text_paint,
                legend_text_paint, title_paint, grid_paint,
            )
            return

        left_bottom = (chart.x + 74.2, chart.y + 102.88)
        right_bottom = (chart.x + 308.25, chart.y + 132.24)
        left_top = (chart.x + 69.2, chart.y + 52.3)
        right_top = (chart.x + 318.6, chart.y + 47.4)
        tick_count = int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        ))
        axis_metrics = axis_font.getMetrics()

        for index in range(tick_count + 1):
            ratio = index / max(1, tick_count)
            wall_points = self._chart_3d_wall_points(chart, ratio)
            (x1, y1), (corner_x, corner_y), (x2, y2) = wall_points
            wall_path = self._skia.Path()
            wall_path.moveTo(x1, y1)
            wall_path.lineTo(corner_x, corner_y)
            wall_path.lineTo(x2, y2)
            canvas.drawPath(wall_path, grid_paint)
            value = chart.value_min + chart.value_step * index
            label = self._chart_value_label(value, chart.axis_number_format)
            label_width = axis_font.measureText(label)
            baseline = y1 - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            canvas.drawString(
                label, x1 - 10.0 - label_width, baseline,
                axis_font, value_text_paint,
            )

        axis_path = self._skia.Path()
        axis_path.moveTo(*left_bottom)
        axis_path.lineTo(*right_bottom)
        canvas.drawPath(axis_path, grid_paint)

        def perspective(position):
            curvature = 0.42
            return (
                math.exp(curvature * position) - 1.0
            ) / (math.exp(curvature) - 1.0)

        def project(position, value):
            horizontal = perspective(position)
            base_x = left_bottom[0] + (
                right_bottom[0] - left_bottom[0]
            ) * horizontal
            base_y = left_bottom[1] + (
                right_bottom[1] - left_bottom[1]
            ) * horizontal
            ratio = (
                (float(value) - chart.value_min)
                / max(1.0, chart.value_max - chart.value_min)
            )
            top_x = left_top[0] + (right_top[0] - left_top[0]) * horizontal
            top_y = left_top[1] + (right_top[1] - left_top[1]) * horizontal
            return (
                base_x + (top_x - base_x) * ratio,
                base_y + (top_y - base_y) * ratio,
            )

        for label, position in chart.axis_categories:
            x, y = project(position, chart.value_min)
            canvas.drawLine(x, y, x, y + 2.8, grid_paint)
            label_width = axis_font.measureText(label)
            canvas.save()
            canvas.translate(x, y + 7.0)
            canvas.rotate(-90.0)
            canvas.drawString(
                label, -label_width, 0.0,
                axis_font, category_text_paint,
            )
            canvas.restore()

        depth_x = ((7.5, 0.0), (28.25, 3.0), (46.0, 7.0))
        for series_index in reversed(range(len(chart.series))):
            series = chart.series[series_index]
            points = sorted(zip(chart.category_positions, series.values))
            projected_points = []
            for point_index, (position, value) in enumerate(points):
                x, y = project(position, value)
                horizontal = perspective(position)
                if series_index < len(depth_x):
                    left_shift, right_shift = depth_x[series_index]
                    x += left_shift * (1.0 - horizontal) + right_shift * horizontal
                y += self._chart_3d_series_y_offset(
                    series_index, horizontal
                )
                projected_points.append((x, y, horizontal))
            def polygon(vertices):
                path = self._skia.Path()
                path.moveTo(*vertices[0])
                for vertex in vertices[1:]:
                    path.lineTo(*vertex)
                path.close()
                return path

            body_color = self._chart_3d_body_color(
                series.color, series_index
            )
            side_color = self._chart_3d_side_color(
                series.color, series_index
            )
            edges = [
                self._chart_3d_face_edges(point)
                for point in projected_points
            ]
            for segment_index in range(len(projected_points) - 1):
                start_front, start_back = edges[segment_index]
                end_front, end_back = edges[segment_index + 1]
                segment_dy = end_front[1] - start_front[1]
                midpoint_horizontal = (
                    projected_points[segment_index][2]
                    + projected_points[segment_index + 1][2]
                ) / 2.0
                strong_side = self._chart_3d_has_strong_side(
                    midpoint_horizontal, segment_dy
                )
                canvas.drawPath(
                    polygon((start_front, end_front, end_back, start_back)),
                    self._paint(
                        'FF' + (side_color if strong_side else body_color),
                        fill=True,
                    ),
                )
                if not strong_side:
                    canvas.drawLine(
                        *start_front, *end_front,
                        self._paint(
                            'FF' + series.color, stroke_width=0.55
                        ),
                    )

        for series_index, series in enumerate(chart.series):
            canvas.drawString(
                series.name,
                chart.x + 320.5,
                chart.y + 105.0 + series_index * 12.0,
                axis_font,
                legend_text_paint,
            )

        self._draw_3d_chart_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
        )

    def _draw_3d_column_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        is_percent_cylinder = (
            getattr(chart, 'bar_shape', 'box') == 'cylinder'
            and chart.grouping == 'percentStacked'
        )
        is_percent_box = (
            getattr(chart, 'bar_shape', 'box') == 'box'
            and chart.grouping == 'percentStacked'
        )
        is_standard_depth = (
            getattr(chart, 'bar_shape', 'box') == 'box'
            and chart.grouping == 'standard'
            and chart.show_series_axis
        )
        is_high_perspective_standard = (
            is_standard_depth and chart.perspective >= 65.0
        )

        width_factor = max(0.0, min(
            1.0, (chart.width - 384.36) / (461.04 - 384.36)
        ))
        height_factor = max(0.0, min(
            1.0, (chart.height - 195.36) / (246.48 - 195.36)
        ))

        def boundary(position):
            if is_percent_cylinder:
                return (
                    chart.x + chart.width * (
                        0.2842 + 0.2535 * position
                        + 0.1239 * position ** 2
                    ),
                    chart.y + chart.height * (
                        0.4610 + 0.1400 * position
                        + 0.0820 * position ** 2
                    ),
                )
            if is_percent_box:
                return (
                    chart.x + 83.46 + 187.28 * position
                    + 43.11 * position ** 2,
                    chart.y + 118.6305 + 23.3979 * position
                    + 5.6271 * position ** 2,
                )
            if is_high_perspective_standard:
                return self._chart_3d_high_perspective_ground_point(
                    chart, position, 0.0
                )
            if is_standard_depth:
                return (
                    chart.x + 65.883 + 195.752 * position
                    + 50.352 * position ** 2,
                    chart.y + 116.830 + 25.457 * position
                    + 5.728 * position ** 2,
                )
            x_coefficients = (
                93.73 + 5.35 * width_factor,
                184.6479 + 58.2535 * width_factor,
                43.11 + 5.2929 * width_factor,
            )
            y_coefficients = (
                118.6305 + 41.7905 * height_factor,
                23.3979 + 7.8085 * height_factor,
                5.6271 + 0.7029 * height_factor,
            )
            return (
                chart.x + x_coefficients[0]
                + x_coefficients[1] * position
                + x_coefficients[2] * position ** 2,
                chart.y + y_coefficients[0]
                + y_coefficients[1] * position
                + y_coefficients[2] * position ** 2,
            )

        def wall_points(ratio):
            front_x, front_y = boundary(0.0)
            if is_percent_cylinder:
                return (
                    (
                        front_x - chart.width * 0.0084 * ratio,
                        front_y - chart.height * 0.2500 * ratio,
                    ),
                    (
                        front_x + chart.width * (0.0760 - 0.0084 * ratio),
                        front_y - chart.height * (
                            0.0180 + 0.2500 * ratio
                        ),
                    ),
                    (
                        chart.x + chart.width * (
                            0.7205 + 0.0094 * ratio
                        ),
                        chart.y + chart.height * (
                            0.6205 - 0.3230 * ratio
                        ),
                    ),
                )
            if is_standard_depth:
                if is_high_perspective_standard:
                    lower = (
                        self._chart_3d_high_perspective_ground_point(
                            chart, 0.0, 0.0
                        ),
                        self._chart_3d_high_perspective_ground_point(
                            chart, 0.0, 1.0
                        ),
                        self._chart_3d_high_perspective_ground_point(
                            chart, 1.0, 1.0
                        ),
                    )
                    upper = self._chart_3d_high_perspective_wall_top(chart)
                    curve_offset = (
                        chart.height * 0.0242 * ratio * (1.0 - ratio)
                    )
                    return tuple(
                        (
                            bottom[0] + (top[0] - bottom[0]) * ratio,
                            bottom[1] + (top[1] - bottom[1]) * ratio
                            + curve_offset,
                        )
                        for bottom, top in zip(lower, upper)
                    )
                rendered_ratio = ratio * (0.953 + 0.047 * ratio)
                return (
                    (
                        chart.x + 63.70 - 5.01 * rendered_ratio,
                        chart.y + 116.13 - 62.04 * rendered_ratio,
                    ),
                    (
                        chart.x + 122.92 - 1.08 * rendered_ratio,
                        chart.y + 90.66 - 52.50 * rendered_ratio,
                    ),
                    (
                        chart.x + 322.50 + 5.40 * rendered_ratio,
                        chart.y + 106.84 - 58.97 * rendered_ratio,
                    ),
                )
            far_x, far_y = boundary(1.0)
            rendered_ratio = ratio * (0.953 + 0.047 * ratio)
            if is_percent_box:
                front_x -= 0.64
                front_y -= 0.53
                front_scale = 0.665
            elif chart.grouping in ('stacked', 'percentStacked'):
                front_y += 0.43
                front_scale = 0.675
            else:
                front_y -= 0.10
                front_scale = 0.689
            front_rise = -chart.plot_height * front_scale * rendered_ratio
            far_rise = -chart.plot_height * 0.823 * rendered_ratio
            return (
                (front_x + front_rise * 0.0706, front_y + front_rise),
                (
                    front_x + chart.width * 0.026,
                    front_y - chart.height * 0.026 + front_rise * 0.98,
                ),
                (
                    far_x + chart.width * 0.023,
                    far_y - chart.height * 0.026 + far_rise,
                ),
            )

        if is_standard_depth and (
            chart.back_wall_fill_image or chart.side_wall_fill_image
        ):
            self._draw_3d_column_walls(
                canvas, chart, wall_points(0.0), wall_points(1.0)
            )

        tick_count = int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        ))
        for index in range(tick_count + 1):
            ratio = index / max(1, tick_count)
            wall = wall_points(ratio)
            path = self._skia.Path()
            path.moveTo(*wall[0])
            path.lineTo(*wall[1])
            path.lineTo(*wall[2])
            canvas.drawPath(path, grid_paint)
            value = chart.value_min + chart.value_step * index
            label = self._chart_value_label(value, chart.axis_number_format)
            label_width = axis_font.measureText(label)
            baseline = wall[0][1] - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            canvas.drawString(
                label,
                wall[0][0]
                - (9.3 if is_high_perspective_standard else 8.0)
                - label_width,
                baseline,
                axis_font, value_text_paint,
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
                    wall = wall_points(ratio)
                    path = self._skia.Path()
                    path.moveTo(*wall[0])
                    path.lineTo(*wall[1])
                    path.lineTo(*wall[2])
                    canvas.drawPath(path, minor_paint)

        if chart.value_axis_color:
            axis_paint = self._paint(
                'FF' + chart.value_axis_color, stroke_width=0.75
            )
            bottom = wall_points(0.0)[0]
            top = wall_points(1.0)[0]
            canvas.drawLine(*bottom, *top, axis_paint)
            if chart.value_major_tick_mark == 'cross':
                for index in range(tick_count + 1):
                    point = wall_points(index / max(1, tick_count))[0]
                    canvas.drawLine(
                        point[0] - 2.8, point[1],
                        point[0] + 2.8, point[1], axis_paint,
                    )

        category_count = max(1, len(chart.categories))
        axis_category_count = (
            max(1, len(chart.axis_categories))
            if is_high_perspective_standard else category_count
        )
        boundaries = tuple(
            boundary(index / axis_category_count)
            for index in range(axis_category_count + 1)
        )
        floor_path = self._skia.Path()
        floor_path.moveTo(*boundaries[0])
        control = (
            2.0 * boundaries[len(boundaries) // 2][0]
            - (boundaries[0][0] + boundaries[-1][0]) / 2.0,
            2.0 * boundaries[len(boundaries) // 2][1]
            - (boundaries[0][1] + boundaries[-1][1]) / 2.0,
        )
        floor_path.quadTo(*control, *boundaries[-1])
        canvas.drawPath(floor_path, grid_paint)
        tick_points = (
            tuple(
                boundary(position)
                for _, position in chart.axis_categories
            )
            if is_high_perspective_standard else boundaries
        )
        for point in tick_points:
            canvas.drawLine(
                point[0], point[1], point[0], point[1] + 2.8, grid_paint
            )

        cylinders = (
            self._chart_3d_stacked_column_cylinders(chart)
            if is_percent_cylinder else ()
        )
        cylinder_bases = {
            category_index: lower_center
            for (
                series_index, category_index, lower_center, _, _, _, _, _
            ) in cylinders
            if series_index == 0
        }
        if is_high_perspective_standard:
            category_labels = tuple(
                (
                    label,
                    *self._chart_3d_high_perspective_category_label_point(
                        chart, position
                    ),
                )
                for label, position in chart.axis_categories
            )
        else:
            category_labels = tuple(
                (
                    label,
                    *cylinder_bases.get(category_index, (
                        (
                            boundaries[category_index][0]
                            + boundaries[category_index + 1][0]
                        ) / 2.0,
                        (
                            boundaries[category_index][1]
                            + boundaries[category_index + 1][1]
                        ) / 2.0,
                    )),
                )
                for category_index, label in enumerate(chart.categories)
            )
        for label, center_x, center_y in category_labels:
            if is_percent_cylinder:
                center_x -= chart.width * 0.0145
                center_y += chart.height * 0.0250
            label_width = axis_font.measureText(label)
            if chart.category_rotation:
                canvas.save()
                canvas.translate(center_x, center_y + 8.0)
                canvas.rotate(chart.category_rotation)
                canvas.drawString(
                    label, -label_width, 0.0,
                    axis_font, category_text_paint,
                )
                canvas.restore()
            else:
                canvas.drawString(
                    label, center_x - label_width / 2.0,
                    center_y + 15.0, axis_font, category_text_paint,
                )

        if is_percent_cylinder:
            self._draw_3d_stacked_column_cylinders(canvas, chart, cylinders)
            prisms = ()
        elif is_high_perspective_standard:
            prisms = self._chart_3d_high_perspective_column_prisms(chart)
        elif is_standard_depth:
            prisms = self._chart_3d_standard_column_prisms(chart, boundaries)
        elif chart.grouping in ('stacked', 'percentStacked'):
            prisms = self._chart_3d_stacked_column_prisms(chart, boundaries)
        else:
            prisms = self._chart_3d_column_prisms(chart, boundaries)
        for series_index, _, front, side, top in prisms:
            series = chart.series[series_index]
            canvas.drawPath(
                self._polygon_path(side),
                self._paint(
                    'FF' + self._scale_chart_color(series.color, 0.64),
                    fill=True,
                ),
            )
            canvas.drawPath(
                self._polygon_path(front),
                self._paint('FF' + series.color, fill=True),
            )
            if top is not None:
                canvas.drawPath(
                    self._polygon_path(top),
                    self._paint(
                        'FF' + self._scale_chart_color(series.color, 0.76),
                        fill=True,
                    ),
                )

        if chart.show_data_labels:
            metrics = axis_font.getMetrics()
            data_label_paint = self._chart_text_paint(
                chart.data_label_color, None
            )
            if is_percent_cylinder:
                for (
                    series_index, category_index, _, upper_center,
                    _, _, _, _,
                ) in cylinders:
                    value = chart.series[series_index].values[category_index]
                    text = self._chart_value_label(
                        value, chart.axis_number_format
                    )
                    text_width = axis_font.measureText(text)
                    canvas.drawString(
                        text, upper_center[0] - text_width / 2.0,
                        upper_center[1] - 8.0 - metrics.fDescent,
                        axis_font, data_label_paint,
                    )
            for series_index, category_index, front, _, _ in prisms:
                value = chart.series[series_index].values[category_index]
                text = self._chart_value_label(
                    value, chart.axis_number_format
                )
                text_width = axis_font.measureText(text)
                top_x = (front[2][0] + front[3][0]) / 2.0
                top_y = (front[2][1] + front[3][1]) / 2.0
                canvas.drawString(
                    text, top_x - text_width / 2.0,
                    top_y - 8.0 - metrics.fDescent,
                    axis_font, data_label_paint,
                )

        if is_standard_depth:
            self._draw_3d_column_series_axis(
                canvas, chart, axis_font, category_text_paint
            )

        self._draw_3d_chart_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
        )

    def _draw_3d_column_walls(self, canvas, chart, lower, upper):
        side = self._polygon_path((
            lower[0], lower[1], upper[1], upper[0],
        ))
        back = self._polygon_path((
            lower[1], lower[2], upper[2], upper[1],
        ))
        if chart.back_wall_fill_image:
            canvas.drawPath(
                back,
                self._chart_bitmap_paint(
                    chart.back_wall_fill_image, 0.36
                ),
            )
        if chart.side_wall_fill_image:
            canvas.drawPath(
                side,
                self._chart_bitmap_paint(
                    chart.side_wall_fill_image, 0.36
                ),
            )
            canvas.drawPath(
                side, self._paint('55000000', fill=True)
            )

    @staticmethod
    def _chart_3d_high_perspective_ground_point(chart, category, depth):
        width_scale = chart.width / 384.36
        height_scale = chart.height / 195.36
        denominator = (
            1.0 - 0.48041958 * category + 0.79254079 * depth
        )
        relative_x = (
            74.73471698
            + 72.65750365 * category
            + 182.35621837 * depth
        ) / denominator
        relative_y = (
            91.13751485
            - 22.95093244 * category
            + 52.33637257 * depth
        ) / denominator
        return (
            chart.x + relative_x * width_scale,
            chart.y + relative_y * height_scale,
        )

    @staticmethod
    def _chart_3d_high_perspective_wall_top(chart):
        width_scale = chart.width / 384.36
        height_scale = chart.height / 195.36
        return tuple(
            (
                chart.x + relative_x * width_scale,
                chart.y + relative_y * height_scale,
            )
            for relative_x, relative_y in (
                (62.14792453, 50.68309215),
                (139.46679245, 52.47311085),
                (257.06339623, 51.75710337),
            )
        )

    @staticmethod
    def _chart_3d_high_perspective_category_label_point(chart, position):
        width_scale = chart.width / 384.36
        height_scale = chart.height / 195.36
        x_projection = position / (1.0 - 0.54071755 * position)
        y_projection = position / (1.0 - 0.54091269 * position)
        return (
            chart.x + (
                78.33708858 + 95.57968917 * x_projection
            ) * width_scale,
            chart.y + (
                88.93293655 + 18.18544995 * y_projection
            ) * height_scale,
        )

    @staticmethod
    def _chart_3d_high_perspective_project(chart, category, depth, value):
        ground_x, ground_y = (
            PdfRenderer3DLineColumnMixin._chart_3d_high_perspective_ground_point(
                chart, category, depth
            )
        )
        denominator = (
            1.0 - 0.48041958 * category + 0.79254079 * depth
        )
        vertical_scale = (
            0.10425 + 0.01200 * category + 0.01955 * depth
        ) / denominator
        vertical_scale *= 0.94
        vanishing_x = chart.x + chart.width * (198.08528302 / 384.36)
        vanishing_y = chart.y + chart.height * (479.21356981 / 195.36)
        return (
            ground_x + value * vertical_scale * (
                ground_x - vanishing_x
            ),
            ground_y + value * vertical_scale * (
                ground_y - vanishing_y
            ),
        )

    @staticmethod
    def _chart_3d_high_perspective_column_prisms(chart):
        category_count = max(1, len(chart.categories))
        axis_count = max(1, len(chart.axis_categories))
        series_count = max(1, len(chart.series))
        gap_depth = 1.5
        total_depth = series_count + (series_count - 1) * gap_depth
        bar_depth = 1.0 / total_depth
        bar_width = 0.36 / axis_count
        y_offset = -chart.height * 0.0073

        def point(category, depth, value, series_index, far_category):
            x, y = PdfRenderer3DLineColumnMixin._chart_3d_high_perspective_project(
                chart, category, depth, value
            )
            if far_category:
                x += chart.width * 0.002808 * series_index
                y -= chart.height * 0.00550
                if series_count > 1 and series_index == series_count - 1:
                    y += chart.height * 0.00550
            return x, y + y_offset

        prisms = []
        category_order = sorted(
            range(category_count),
            key=chart.category_positions.__getitem__,
        )
        for series_index in reversed(range(series_count)):
            depth_start = (
                series_index * (1.0 + gap_depth) / total_depth
            )
            if series_count > 1 and series_index == series_count - 1:
                depth_start -= bar_depth / 2.0
            depth_end = depth_start + bar_depth
            depth_center = (depth_start + depth_end) / 2.0
            for category_index in category_order:
                raw_category = chart.category_positions[category_index]
                category_center = raw_category + (1.0 - raw_category) * (
                    0.02191474
                    - 0.15776136 * raw_category
                    + 0.08273078 * raw_category ** 2
                )
                far_category = raw_category >= 0.75
                category_width = bar_width * (
                    0.97 if far_category else 1.0
                )
                # At the far end of a high-perspective date axis, the visible
                # category face expands toward the depth vanishing point.
                category_start = category_center - category_width * (
                    0.80 if far_category else 0.50
                )
                category_end = category_center + category_width / 2.0
                ratio = max(0.0, min(
                    1.0,
                    (
                        float(chart.series[series_index].values[category_index])
                        - chart.value_min
                    ) / max(1.0, chart.value_max - chart.value_min),
                ))
                depth_face = (
                    point(
                        category_start, depth_center, 0.0,
                        series_index, far_category,
                    ),
                    point(
                        category_end, depth_center, 0.0,
                        series_index, far_category,
                    ),
                    point(
                        category_end, depth_center, ratio,
                        series_index, far_category,
                    ),
                    point(
                        category_start, depth_center, ratio,
                        series_index, far_category,
                    ),
                )
                front_depth_x = point(
                    category_center, depth_start, 0.0,
                    series_index, far_category,
                )[0]
                back_depth_x = point(
                    category_center, depth_end, 0.0,
                    series_index, far_category,
                )[0]
                visible_depth_end = depth_end + (
                    0.38 if far_category else 0.08
                ) * bar_depth
                side_category = (
                    category_end
                    if back_depth_x > front_depth_x
                    else category_start
                )
                side = (
                    point(
                        side_category, depth_center, 0.0,
                        series_index, far_category,
                    ),
                    point(
                        side_category, visible_depth_end, 0.0,
                        series_index, far_category,
                    ),
                    point(
                        side_category, visible_depth_end, ratio,
                        series_index, far_category,
                    ),
                    point(
                        side_category, depth_center, ratio,
                        series_index, far_category,
                    ),
                )
                top = (
                    point(
                        category_start, depth_center, ratio,
                        series_index, far_category,
                    ),
                    point(
                        category_end, depth_center, ratio,
                        series_index, far_category,
                    ),
                    point(
                        category_end, visible_depth_end, ratio,
                        series_index, far_category,
                    ),
                    point(
                        category_start, visible_depth_end, ratio,
                        series_index, far_category,
                    ),
                )
                front = depth_face
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @staticmethod
    def _chart_3d_standard_column_prisms(chart, boundaries):
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        vanishing_x, vanishing_y = (
            PdfRenderer3DLineColumnMixin._chart_3d_column_vertical_vanishing_point(chart)
        )

        def project_x(base_x, base_y, target_y):
            denominator = vanishing_y - base_y
            if abs(denominator) < 1e-9:
                return base_x
            return base_x + (
                (vanishing_x - base_x)
                * (target_y - base_y) / denominator
            )

        prisms = []
        for series_index in reversed(range(series_count)):
            for category_index in reversed(range(category_count)):
                left = boundaries[category_index]
                right = boundaries[category_index + 1]
                depth = (category_index + 0.5) / category_count
                series_shift_x = chart.width * (
                    0.0580 - 0.0430 * depth
                ) * series_index
                series_shift_x -= chart.width * (
                    0.0080 + 0.0080 * depth
                ) * (series_index * (series_index - 1) / 2.0)
                series_shift_y = -chart.height * (
                    0.0462 + 0.0264 * depth
                ) * series_index
                interval = max(1.0, right[0] - left[0])
                floor_slope = (right[1] - left[1]) / interval
                center_x = (
                    (left[0] + right[0]) / 2.0
                    + chart.width * (0.0230 - 0.0250 * depth)
                    + series_shift_x
                )
                center_y = (
                    (left[1] + right[1]) / 2.0
                    - chart.height * 0.0265
                    + series_shift_y
                )
                bar_width = chart.width * (
                    0.0390 + 0.0015 * depth
                )
                base_left = (
                    center_x - bar_width / 2.0,
                    center_y - bar_width * floor_slope / 2.0,
                )
                base_right = (
                    center_x + bar_width / 2.0,
                    center_y + bar_width * floor_slope / 2.0,
                )
                value = float(
                    chart.series[series_index].values[category_index]
                )
                ratio = max(0.0, min(
                    1.0,
                    (value - chart.value_min)
                    / max(1.0, chart.value_max - chart.value_min),
                ))
                rise_factor = (
                    0.524 + 0.140 * depth
                    - (0.025 + 0.035 * depth) * series_index
                )
                rise = -chart.plot_height * rise_factor * ratio
                top_left_y = base_left[1] + rise
                top_right_y = base_right[1] + rise
                top_left = (
                    project_x(*base_left, top_left_y), top_left_y
                )
                top_right = (
                    project_x(*base_right, top_right_y), top_right_y
                )
                face_depth_x = chart.width * 0.0170
                face_depth_y = -chart.height * 0.0140
                front = (
                    base_left, base_right, top_right, top_left,
                )
                side = (
                    base_right,
                    (
                        base_right[0] + face_depth_x,
                        base_right[1] + face_depth_y,
                    ),
                    (
                        top_right[0] + face_depth_x,
                        top_right[1] + face_depth_y,
                    ),
                    top_right,
                )
                top = (
                    top_left,
                    top_right,
                    (
                        top_right[0] + face_depth_x,
                        top_right[1] + face_depth_y,
                    ),
                    (
                        top_left[0] + face_depth_x,
                        top_left[1] + face_depth_y,
                    ),
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    def _draw_3d_column_series_axis(
        self, canvas, chart, axis_font, text_paint
    ):
        metrics = axis_font.getMetrics()
        perspective = max(0.0, min(
            1.0, (chart.perspective - 30.0) / 70.0
        ))
        for series_index, series in enumerate(chart.series):
            x = chart.x + (
                321.70 + (281.02 - 321.70) * perspective
                + (3.90 + (-14.575 - 3.90) * perspective) * series_index
                + (-0.33 + (2.715 + 0.33) * perspective)
                * series_index ** 2
            )
            top = chart.y + (
                133.61 + (109.75 - 133.61) * perspective
                + (-15.64 + (-21.095 + 15.64) * perspective)
                * series_index
                + (1.19 + (3.935 - 1.19) * perspective)
                * series_index ** 2
            )
            canvas.drawString(
                series.name, x, top - metrics.fAscent,
                axis_font, text_paint,
            )

    @staticmethod
    def _chart_3d_column_prisms(chart, boundaries):
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        perspective_factor = max(0.0, min(
            1.0,
            (chart.height - 195.36) / (246.48 - 195.36),
        ))
        vertical_vanishing_x, vertical_vanishing_y = (
            PdfRenderer3DLineColumnMixin._chart_3d_column_vertical_vanishing_point(chart)
        )

        def projected_top_x(bottom_x, base_y, top_y):
            denominator = vertical_vanishing_y - base_y
            if abs(denominator) < 1e-9:
                return bottom_x
            return bottom_x + (
                (vertical_vanishing_x - bottom_x)
                * (top_y - base_y) / denominator
            )

        prisms = []
        for category_index in reversed(range(category_count)):
            left = boundaries[category_index]
            right = boundaries[category_index + 1]
            center_x = (
                (left[0] + right[0]) / 2.0 + chart.width * 0.009
            )
            center_y = (left[1] + right[1]) / 2.0
            interval = max(1.0, right[0] - left[0])
            floor_slope = (right[1] - left[1]) / interval
            bar_width = interval * 0.207
            series_step_x = interval * 0.222
            cluster_width = bar_width + (series_count - 1) * series_step_x
            for series_index in range(series_count):
                value = float(chart.series[series_index].values[category_index])
                ratio = max(0.0, min(
                    1.0,
                    (value - chart.value_min)
                    / max(1.0, chart.value_max - chart.value_min),
                ))
                category_depth = (category_index + 0.5) / category_count
                depth_shift_x = perspective_factor * (
                    0.22 + 3.4 * category_depth
                )
                x = (
                    center_x - cluster_width / 2.0
                    + series_index * series_step_x
                    - depth_shift_x
                )
                depth_shift_y = perspective_factor * (
                    3.42 + 0.93 * category_depth
                )
                bottom_y = center_y + (
                    x + bar_width / 2.0 - center_x
                ) * floor_slope - depth_shift_y
                bottom_right_y = bottom_y + bar_width * floor_slope
                rise_factor = (
                    0.699 + 0.138 * category_depth
                    - perspective_factor * (
                        0.0251 - 0.00635 * category_depth
                    )
                )
                rise_y = -chart.plot_height * rise_factor * ratio
                top_left_y = bottom_y + rise_y
                top_right_y = bottom_right_y + rise_y
                top_left_x = projected_top_x(x, bottom_y, top_left_y)
                top_right_x = projected_top_x(
                    x + bar_width, bottom_right_y, top_right_y
                )
                depth_x = chart.width * 0.011
                depth_y = -chart.height * 0.010
                front = (
                    (x, bottom_y),
                    (x + bar_width, bottom_right_y),
                    (top_right_x, top_right_y),
                    (top_left_x, top_left_y),
                )
                side = (
                    front[1],
                    (front[1][0] + depth_x, front[1][1] + depth_y),
                    (front[2][0] + depth_x, front[2][1] + depth_y),
                    front[2],
                )
                top = (
                    front[3], front[2],
                    (front[2][0] + depth_x, front[2][1] + depth_y),
                    (front[3][0] + depth_x, front[3][1] + depth_y),
                )
                prisms.append((
                    series_index, category_index, front, side, top
                ))
        return tuple(prisms)

    @staticmethod
    def _chart_3d_column_vertical_vanishing_point(chart):
        # A shared vertical vanishing point keeps every column face on one
        # projective plane while preserving the series-specific edge angles.
        frame_factor = max(0.0, min(
            1.0,
            (chart.height - 195.36) / (246.48 - 195.36),
        ))
        x_ratio = 0.5558794 + (0.5406128 - 0.5558794) * frame_factor
        y_ratio = 8.5953190 + (9.8459911 - 8.5953190) * frame_factor
        return (
            chart.x + chart.width * x_ratio,
            chart.y + chart.height * y_ratio,
        )

    @staticmethod
    def _chart_3d_stacked_column_prisms(chart, boundaries):
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        vanishing_x, vanishing_y = (
            PdfRenderer3DLineColumnMixin._chart_3d_column_vertical_vanishing_point(chart)
        )

        def project_x(base_x, base_y, target_y):
            denominator = vanishing_y - base_y
            if abs(denominator) < 1e-9:
                return base_x
            return base_x + (
                (vanishing_x - base_x)
                * (target_y - base_y) / denominator
            )

        prisms = []
        value_range = max(1.0, chart.value_max - chart.value_min)
        for category_index in reversed(range(category_count)):
            left = boundaries[category_index]
            right = boundaries[category_index + 1]
            category_depth = (category_index + 0.5) / category_count
            category_center_x = (left[0] + right[0]) / 2.0
            category_center_y = (left[1] + right[1]) / 2.0
            interval = max(1.0, right[0] - left[0])
            floor_slope = (right[1] - left[1]) / interval
            bar_width = interval * 0.367
            center_offset_x = chart.width * (
                0.020879 - 0.014398 * category_depth
                - 0.010989 * category_depth ** 2
            )
            bar_center_x = category_center_x + center_offset_x
            depth_shift_y = chart.height * (
                0.022047 + 0.003107 * category_depth
            )
            bar_center_y = (
                category_center_y
                + center_offset_x * floor_slope
                - depth_shift_y
            )
            base_left = (
                bar_center_x - bar_width / 2.0,
                bar_center_y - bar_width * floor_slope / 2.0,
            )
            base_right = (
                bar_center_x + bar_width / 2.0,
                bar_center_y + bar_width * floor_slope / 2.0,
            )
            rise_factor = 0.624251 + 0.147962 * category_depth
            depth_x = chart.width * (
                0.019531 - 0.006666 * category_depth
                - 0.008414 * category_depth ** 2
            )
            depth_y = -chart.height * 0.0165
            values = [
                max(0.0, float(series.values[category_index]))
                for series in chart.series
            ]
            category_total = max(1.0, sum(values))
            cumulative_ratio = 0.0
            for series_index, value in enumerate(values):
                lower_ratio = cumulative_ratio
                if chart.grouping == 'percentStacked':
                    cumulative_ratio += value / category_total
                else:
                    cumulative_ratio += value / value_range
                upper_ratio = min(1.0, cumulative_ratio)
                lower_left_y = (
                    base_left[1]
                    - chart.plot_height * rise_factor * lower_ratio
                )
                lower_right_y = (
                    base_right[1]
                    - chart.plot_height * rise_factor * lower_ratio
                )
                upper_left_y = (
                    base_left[1]
                    - chart.plot_height * rise_factor * upper_ratio
                )
                upper_right_y = (
                    base_right[1]
                    - chart.plot_height * rise_factor * upper_ratio
                )
                lower_left = (
                    project_x(*base_left, lower_left_y), lower_left_y
                )
                lower_right = (
                    project_x(*base_right, lower_right_y), lower_right_y
                )
                upper_left = (
                    project_x(*base_left, upper_left_y), upper_left_y
                )
                upper_right = (
                    project_x(*base_right, upper_right_y), upper_right_y
                )
                front = (
                    lower_left, lower_right, upper_right, upper_left,
                )
                side = (
                    lower_right,
                    (lower_right[0] + depth_x, lower_right[1] + depth_y),
                    (upper_right[0] + depth_x, upper_right[1] + depth_y),
                    upper_right,
                )
                top = None
                if series_index == series_count - 1:
                    top = (
                        upper_left,
                        upper_right,
                        (upper_right[0] + depth_x, upper_right[1] + depth_y),
                        (upper_left[0] + depth_x, upper_left[1] + depth_y),
                    )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @staticmethod
    def _chart_3d_stacked_column_cylinders(chart):
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        cylinders = []
        for category_index in reversed(range(category_count)):
            depth = (category_index + 0.5) / category_count
            base_center = (
                chart.x + chart.width * (
                    0.31521 + 0.25348 * depth
                    + 0.12387 * depth ** 2
                ),
                chart.y + chart.height * (
                    0.44733 + 0.14010 * depth
                    + 0.06598 * depth ** 2
                ),
            )
            top_center = (
                chart.x + chart.width * (
                    0.31041 + 0.25768 * depth
                    + 0.12928 * depth ** 2
                ),
                chart.y + chart.height * (
                    0.20583 + 0.07192 * depth
                    + 0.03296 * depth ** 2
                ),
            )
            base_width = chart.width * (
                0.03142 + 0.00373 * depth
                + 0.00541 * depth ** 2
            )
            top_width = chart.width * (
                0.03117 + 0.00633 * depth
                + 0.00601 * depth ** 2
            )
            radius_y = chart.height * (0.0072 + 0.0008 * depth)
            values = [
                max(0.0, float(series.values[category_index]))
                for series in chart.series
            ]
            category_total = max(1.0, sum(values))

            def cross_section(ratio):
                return (
                    (
                        base_center[0]
                        + (top_center[0] - base_center[0]) * ratio,
                        base_center[1]
                        + (top_center[1] - base_center[1]) * ratio,
                    ),
                    (
                        base_width
                        + (top_width - base_width) * ratio
                    ) / 2.0,
                )

            cumulative_ratio = 0.0
            for series_index, value in enumerate(values):
                lower_ratio = cumulative_ratio
                cumulative_ratio = min(
                    1.0, cumulative_ratio + value / category_total
                )
                lower_center, lower_radius = cross_section(lower_ratio)
                upper_center, upper_radius = cross_section(cumulative_ratio)
                cylinders.append((
                    series_index,
                    category_index,
                    lower_center,
                    upper_center,
                    lower_radius,
                    upper_radius,
                    radius_y,
                    series_index == series_count - 1,
                ))
        return tuple(cylinders)

    def _draw_3d_stacked_column_cylinders(
        self, canvas, chart, cylinders
    ):
        for (
            series_index, _, lower_center, upper_center,
            lower_radius, upper_radius, radius_y, show_top,
        ) in cylinders:
            color = chart.series[series_index].color
            body = self._chart_cylinder_body_path(
                lower_center, upper_center,
                lower_radius, upper_radius, radius_y,
            )
            canvas.drawPath(
                body,
                self._chart_cylinder_paint(
                    color,
                    min(
                        lower_center[0] - lower_radius,
                        upper_center[0] - upper_radius,
                    ),
                    max(
                        lower_center[0] + lower_radius,
                        upper_center[0] + upper_radius,
                    ),
                ),
            )
            canvas.drawPath(
                body,
                self._paint(
                    'FF' + self._scale_chart_color(color, 0.45),
                    stroke_width=0.35,
                ),
            )
            if show_top:
                cap = self._chart_cylinder_cap_path(
                    upper_center, upper_radius, radius_y
                )
                canvas.drawPath(
                    cap,
                    self._chart_cylinder_paint(
                        color,
                        upper_center[0] - upper_radius,
                        upper_center[0] + upper_radius,
                        cap=True,
                    ),
                )
                canvas.drawPath(
                    cap,
                    self._paint(
                        'FF' + self._scale_chart_color(color, 0.45),
                        stroke_width=0.35,
                    ),
                )

    def _chart_cylinder_body_path(
        self, lower_center, upper_center,
        lower_radius, upper_radius, radius_y,
    ):
        lower_x, lower_y = lower_center
        upper_x, upper_y = upper_center
        path = self._skia.Path()
        path.moveTo(lower_x - lower_radius, lower_y)
        path.lineTo(upper_x - upper_radius, upper_y)
        path.cubicTo(
            upper_x - upper_radius, upper_y - radius_y,
            upper_x + upper_radius, upper_y - radius_y,
            upper_x + upper_radius, upper_y,
        )
        path.lineTo(lower_x + lower_radius, lower_y)
        path.cubicTo(
            lower_x + lower_radius, lower_y + radius_y,
            lower_x - lower_radius, lower_y + radius_y,
            lower_x - lower_radius, lower_y,
        )
        path.close()
        return path

    def _chart_cylinder_cap_path(self, center, radius_x, radius_y):
        center_x, center_y = center
        control = 0.55228475
        path = self._skia.Path()
        path.moveTo(center_x - radius_x, center_y)
        path.cubicTo(
            center_x - radius_x, center_y - control * radius_y,
            center_x - control * radius_x, center_y - radius_y,
            center_x, center_y - radius_y,
        )
        path.cubicTo(
            center_x + control * radius_x, center_y - radius_y,
            center_x + radius_x, center_y - control * radius_y,
            center_x + radius_x, center_y,
        )
        path.cubicTo(
            center_x + radius_x, center_y + control * radius_y,
            center_x + control * radius_x, center_y + radius_y,
            center_x, center_y + radius_y,
        )
        path.cubicTo(
            center_x - control * radius_x, center_y + radius_y,
            center_x - radius_x, center_y + control * radius_y,
            center_x - radius_x, center_y,
        )
        path.close()
        return path

    def _chart_cylinder_paint(self, color, left, right, cap=False):
        factors = (
            (0.0, 0.38 if cap else 0.48),
            (0.14, 0.82 if cap else 0.92),
            (0.32, 0.92 if cap else 1.08),
            (0.48, 0.84 if cap else 1.00),
            (0.70, 0.62 if cap else 0.76),
            (1.0, 0.30 if cap else 0.38),
        )
        colors = []
        positions = []
        for position, factor in factors:
            shaded = self._scale_chart_color(color, factor)
            alpha, red, green, blue = self._argb_channels(shaded)
            colors.append(self._skia.ColorSetARGB(alpha, red, green, blue))
            positions.append(position)
        paint = self._skia.Paint(AntiAlias=True)
        paint.setShader(self._skia.GradientShader.MakeLinear(
            (
                self._skia.Point(left, 0.0),
                self._skia.Point(max(left + 0.01, right), 0.0),
            ),
            colors,
            positions,
            self._skia.TileMode.kClamp,
        ))
        return paint

    def _polygon_path(self, points):
        path = self._skia.Path()
        path.moveTo(*points[0])
        for point in points[1:]:
            path.lineTo(*point)
        path.close()
        return path

    @staticmethod
    def _scale_chart_color(color, factor):
        channels = [int(color[index:index + 2], 16) for index in (0, 2, 4)]
        return ''.join(
            f'{max(0, min(255, round(channel * factor))):02X}'
            for channel in channels
        )

    def _draw_perspective_3d_line_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        tick_count = int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        ))
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        category_step = 1.0 / max(1, category_count - 1)
        depth_scale = max(0.1, chart.depth_percent / 100.0)
        series_step = depth_scale / max(1, series_count - 1)
        near_depth = -series_step / 2.0
        far_depth = depth_scale + series_step / 2.0
        category_end = 1.0 + category_step / 2.0

        def project(category, depth, value_ratio):
            return self._chart_3d_perspective_point(
                chart, category, depth, value_ratio
            )

        for index in range(tick_count + 1):
            ratio = index / max(1, tick_count)
            rendered_ratio = 0.03 + 0.97 * ratio
            category_start = (
                -0.424 * category_step
                - 0.116 * category_step * ratio
            )
            wall_points = (
                project(category_start, near_depth, rendered_ratio),
                project(category_start, far_depth, rendered_ratio),
                project(category_end, far_depth, rendered_ratio),
            )
            wall_path = self._skia.Path()
            wall_path.moveTo(*wall_points[0])
            wall_path.lineTo(*wall_points[1])
            wall_path.lineTo(*wall_points[2])
            canvas.drawPath(wall_path, grid_paint)

            value = chart.value_min + chart.value_step * index
            label = self._chart_value_label(value, chart.axis_number_format)
            label_width = axis_font.measureText(label)
            label_x, label_y = wall_points[0]
            baseline = label_y - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            canvas.drawString(
                label, label_x - 10.0 - label_width, baseline,
                axis_font, value_text_paint,
            )

        floor_ratio = 0.03
        floor_start_category = -0.424 * category_step
        floor_start = project(
            floor_start_category, near_depth, floor_ratio
        )
        floor_end = project(category_end, near_depth, floor_ratio)
        canvas.drawLine(*floor_start, *floor_end, grid_paint)

        if category_count == 1:
            category_positions = (0.5,)
        else:
            category_positions = tuple(
                index / (category_count - 1)
                for index in range(category_count)
            )
        for label, position in zip(chart.categories, category_positions):
            x, y = project(position, near_depth, floor_ratio)
            canvas.drawLine(x, y, x + 0.6, y + 3.0, grid_paint)
            label_width = axis_font.measureText(label)
            center_offset = 3.75 - 3.25 * position
            canvas.drawString(
                label,
                x + center_offset - label_width / 2.0,
                y + 17.5,
                axis_font,
                category_text_paint,
            )

        face_depth = series_step * 0.375
        for series_index in reversed(range(series_count)):
            series = chart.series[series_index]
            depth = series_index * series_step
            projected_edges = []
            for position, value in zip(category_positions, series.values):
                value_ratio = (
                    (float(value) - chart.value_min)
                    / max(1.0, chart.value_max - chart.value_min)
                )
                projected_edges.append(
                    self._chart_3d_perspective_face_edges(
                        chart, position, depth, value_ratio, face_depth
                    )
                )
            for segment_index in range(len(projected_edges) - 1):
                start_front, start_back = projected_edges[segment_index]
                end_front, end_back = projected_edges[segment_index + 1]
                path = self._skia.Path()
                path.moveTo(*start_front)
                path.lineTo(*end_front)
                path.lineTo(*end_back)
                path.lineTo(*start_back)
                path.close()
                color = self._chart_3d_perspective_face_color(
                    series.color, series_index, segment_index
                )
                canvas.drawPath(
                    path, self._paint('FF' + color, fill=True)
                )

        series_axis_x = category_end
        for series_index, series in enumerate(chart.series):
            depth = series_index * series_step
            x, y = project(series_axis_x, depth, floor_ratio)
            baseline = y - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            canvas.drawString(
                series.name, x + 6.0, baseline,
                axis_font, category_text_paint,
            )

        self._draw_3d_chart_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
        )

    def _draw_3d_chart_title_and_legend(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint, legend_series=None,
        legend_offset_x=0.0, legend_font_scale_x=1.0,
        legend_item_gap=7.0, title_font_scale_x=1.0,
    ):
        render_title_font = title_font
        if title_font_scale_x != 1.0:
            render_title_font = self._font(
                chart.title_font_resolution,
                getattr(chart.title_font_resolution, 'size_points', 14.0),
            )
            render_title_font.setScaleX(title_font_scale_x)
        title_width = render_title_font.measureText(chart.title)
        title_metrics = render_title_font.getMetrics()
        if chart.title_layout_x is None:
            title_x = chart.x + (chart.width - title_width) / 2.0
        else:
            title_x = chart.x + chart.width * chart.title_layout_x + 3.0
        if chart.title_layout_y is None:
            title_y = chart.y + 7.5 - title_metrics.fAscent
        else:
            title_y = (
                chart.y + chart.height * chart.title_layout_y
                + 1.54 - title_metrics.fAscent
            )
        canvas.drawString(
            chart.title, title_x, title_y, render_title_font, title_paint
        )
        if not chart.show_legend:
            return

        legend_font = axis_font
        if legend_font_scale_x != 1.0:
            legend_font = self._font(chart.axis_font_resolution, 9.0)
            legend_font.setScaleX(legend_font_scale_x)
        axis_metrics = legend_font.getMetrics()
        legend_series = (
            tuple(chart.series)
            if legend_series is None else tuple(legend_series)
        )
        item_widths = [
            9.0 + legend_font.measureText(series.name) + legend_item_gap
            for series in legend_series
        ]
        legend_x = chart.x + (chart.width - sum(item_widths)) / 2.0 + 7.0
        legend_x += legend_offset_x
        legend_y = chart.y + chart.height - 20.5
        legend_baseline = legend_y - axis_metrics.fAscent
        for series, item_width in zip(legend_series, item_widths):
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(legend_x, legend_y + 3.47, 5.0, 5.0),
                self._paint('FF' + series.color, fill=True),
            )
            canvas.drawString(
                series.name, legend_x + 8.0, legend_baseline,
                legend_font, legend_text_paint,
            )
            legend_x += item_width

    @staticmethod
    def _chart_3d_perspective_point(chart, category, depth, value_ratio):
        # Office's default 15/20-degree camera at perspective 60, normalized
        # to the chart frame. Keeping this projective mapping shared makes the
        # walls, axes, depth labels, and series faces meet at the same points.
        denominator = (
            1.0 - 0.147965 * category
            + 1.115199 * depth
            - 0.063767 * value_ratio
        )
        normalized_x = (
            0.37311492 + 0.16015329 * category
            + 0.79855236 * depth
            - 0.02877723 * value_ratio
        ) / denominator
        normalized_y = (
            0.59233634 - 0.03100446 * category
            + 0.21998678 * depth
            - 0.40400057 * value_ratio
        ) / denominator
        return (
            chart.x + chart.width * normalized_x,
            chart.y + chart.height * normalized_y,
        )

    @staticmethod
    def _chart_3d_perspective_face_color(
        color, series_index, segment_index
    ):
        factors = (
            (0.81, 0.80),
            (0.745, 0.745),
            (0.80, 0.82),
            (0.815, 0.82),
            (0.82, 0.83),
            (0.82, 0.82),
        )
        series_factors = factors[min(series_index, len(factors) - 1)]
        factor = series_factors[min(segment_index, len(series_factors) - 1)]
        channels = [int(color[index:index + 2], 16) for index in (0, 2, 4)]
        return ''.join(f'{round(channel * factor):02X}' for channel in channels)

    @classmethod
    def _chart_3d_perspective_face_edges(
        cls, chart, category, depth, value_ratio, face_depth
    ):
        half_depth = face_depth / 2.0
        return (
            cls._chart_3d_perspective_point(
                chart, category, depth - half_depth, value_ratio
            ),
            cls._chart_3d_perspective_point(
                chart, category, depth + half_depth, value_ratio
            ),
        )

    @staticmethod
    def _chart_3d_body_color(color, series_index):
        factors = (0.80, 0.815, 0.77)
        factor = factors[series_index % len(factors)]
        channels = [int(color[index:index + 2], 16) for index in (0, 2, 4)]
        return ''.join(f'{round(channel * factor):02X}' for channel in channels)

    @staticmethod
    def _chart_3d_side_color(color, series_index):
        factors = (0.77, 0.40, 0.70)
        factor = factors[series_index % len(factors)]
        channels = [int(color[index:index + 2], 16) for index in (0, 2, 4)]
        return ''.join(f'{round(channel * factor):02X}' for channel in channels)

    @staticmethod
    def _chart_3d_has_strong_side(midpoint_horizontal, segment_dy):
        # These projected orientations face away from Excel's default light.
        return (
            segment_dy < -12.0
            or (
                0.0 <= segment_dy <= 2.0
                and 0.25 <= midpoint_horizontal <= 0.45
            )
        )

    @classmethod
    def _chart_3d_face_edges(cls, point):
        x, y, horizontal = point
        depth_x, depth_y = cls._chart_3d_depth_vector(horizontal)
        return (
            (x, y),
            (x + depth_x, y + depth_y),
        )

    @staticmethod
    def _chart_3d_depth_vector(horizontal):
        # Excel projects line depth mostly along the receding category axis.
        # The vector turns upward as it approaches the front of the plot.
        depth_x = 0.75 + 7.0 * (1.0 - horizontal) ** 1.3
        depth_y = -2.5 - 2.0 * horizontal ** 5
        return depth_x, depth_y

    @staticmethod
    def _chart_3d_series_y_offset(series_index, horizontal):
        back_ratio = max(0.0, 1.0 - horizontal / 0.4)
        if series_index == 0:
            return 4.0 * (1.0 - horizontal) - 5.65 * back_ratio
        if series_index == 1:
            return (
                -11.1 * back_ratio
                - 18.0 * max(0.0, (horizontal - 0.4) / 0.6)
            )
        if series_index == 2:
            return (
                -25.0 * horizontal
                - 18.0 * (1.0 - horizontal) ** 3
            )
        return 0.0

    @staticmethod
    def _chart_3d_wall_points(chart, ratio):
        bottom = (
            (chart.x + 74.2, chart.y + 102.88),
            (chart.x + 128.56, chart.y + 78.18),
            (chart.x + 313.56, chart.y + 94.48),
        )
        top = (
            (chart.x + 69.2, chart.y + 52.3),
            (chart.x + 128.56, chart.y + 38.08),
            (chart.x + 318.6, chart.y + 47.4),
        )
        return tuple(
            (
                bottom[index][0]
                + (top[index][0] - bottom[index][0]) * ratio,
                bottom[index][1]
                + (top[index][1] - bottom[index][1]) * ratio,
            )
            for index in range(3)
        )

