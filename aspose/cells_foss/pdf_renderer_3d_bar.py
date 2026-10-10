"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRenderer3DBarMixin:
    @staticmethod
    def _chart_3d_bar_plot_point(chart, pixel_x, pixel_y):
        return (
            chart.plot_x + chart.plot_width * pixel_x / 836.0,
            chart.plot_y + chart.plot_height * pixel_y / 254.0,
        )

    @staticmethod
    def _uses_fan_3d_bar_projection(chart):
        return (
            chart.grouping == 'percentStacked'
            and chart.bar_shape == 'box'
            and abs(chart.rotation_x) < 1.0
            and 25.0 <= chart.rotation_y <= 35.0
            and 35.0 <= chart.perspective <= 45.0
        )

    @staticmethod
    def _uses_parallel_percent_3d_bar_projection(chart):
        return (
            chart.grouping == 'percentStacked'
            and chart.bar_shape == 'box'
            and 14.0 <= chart.rotation_x <= 16.0
            and 19.0 <= chart.rotation_y <= 21.0
            and 29.0 <= chart.perspective <= 31.0
        )

    @staticmethod
    def _uses_high_perspective_3d_bar_projection(chart):
        return (
            chart.grouping == 'clustered'
            and chart.bar_shape == 'box'
            and 14.0 <= chart.rotation_x <= 16.0
            and 19.0 <= chart.rotation_y <= 21.0
            and chart.perspective >= 65.0
        )

    @staticmethod
    def _chart_3d_high_perspective_bar_wall_corners(chart):
        width_scale = chart.width / 384.36
        height_scale = chart.height / 195.36

        def point(relative_x, relative_y):
            return (
                chart.x + relative_x * width_scale,
                chart.y + relative_y * height_scale,
            )

        return (
            point(74.73, 49.98),
            point(77.61, 51.05),
            point(285.13, 40.66),
            point(269.30, 147.43),
            point(90.55, 109.45),
            point(89.83, 109.81),
        )

    @staticmethod
    def _chart_3d_high_perspective_bar_value_ratio(ratio):
        ratio = max(0.0, min(1.0, float(ratio)))
        return 0.552 * ratio / max(1e-9, 1.0 - 0.448 * ratio)

    @classmethod
    def _chart_3d_high_perspective_bar_surface_point(
        cls, chart, category_position, value_ratio,
    ):
        category = max(0.0, min(1.0, float(category_position)))
        top, bottom = cls._chart_3d_bar_wall_points(chart, value_ratio)
        return (
            bottom[0] + (top[0] - bottom[0]) * category,
            bottom[1] + (top[1] - bottom[1]) * category,
        )

    @staticmethod
    def _chart_3d_fan_ratio(ratio):
        ratio = max(0.0, min(1.0, float(ratio)))
        return (
            0.71299279 * ratio
            + 0.06425005 * ratio ** 2
            + 0.22275716 * ratio ** 3
        )

    @staticmethod
    def _chart_3d_fan_plot_point(chart, pixel_x, pixel_y):
        return (
            chart.plot_x + chart.plot_width * pixel_x / 895.0,
            chart.plot_y + chart.plot_height * pixel_y / 313.0,
        )

    @staticmethod
    def _chart_3d_parallel_percent_ratio(ratio):
        ratio = max(0.0, min(1.0, float(ratio)))
        return (
            0.81801617 * ratio
            + 0.17873821 * ratio ** 2
            + 0.00324562 * ratio ** 3
        )

    @staticmethod
    def _chart_3d_parallel_percent_plot_point(chart, pixel_x, pixel_y):
        return (
            chart.plot_x + chart.plot_width * pixel_x / 895.0,
            chart.plot_y + chart.plot_height * pixel_y / 312.0,
        )

    @staticmethod
    def _chart_3d_bar_x_scales(chart):
        baseline = 300.54 / 836.0
        plot_scale = chart.plot_width / 836.0
        growth = max(0.0, plot_scale - baseline)
        return (
            baseline + growth * 0.28,
            plot_scale + growth * 0.35,
            baseline + growth * 0.40,
            plot_scale + growth * 0.28,
            baseline + growth * 0.32,
        )

    @staticmethod
    def _chart_3d_bar_y_scales(chart):
        baseline = 91.07 / 254.0
        plot_scale = chart.plot_height / 254.0
        growth = max(0.0, plot_scale - baseline)
        return (
            baseline,
            growth,
            baseline + growth * 0.57,
        )

    @classmethod
    def _chart_3d_bar_wall_points(cls, chart, ratio):
        ratio = max(0.0, min(1.0, float(ratio)))
        if cls._uses_high_perspective_3d_bar_projection(chart):
            (
                _, top_left, top_right,
                bottom_right, bottom_left, _,
            ) = cls._chart_3d_high_perspective_bar_wall_corners(chart)
            projected = cls._chart_3d_high_perspective_bar_value_ratio(ratio)
            return (
                (
                    top_left[0] + (top_right[0] - top_left[0]) * projected,
                    top_left[1] + (top_right[1] - top_left[1]) * projected,
                ),
                (
                    bottom_left[0]
                    + (bottom_right[0] - bottom_left[0]) * projected,
                    bottom_left[1]
                    + (bottom_right[1] - bottom_left[1]) * projected,
                ),
            )
        if cls._uses_fan_3d_bar_projection(chart):
            projected = cls._chart_3d_fan_ratio(ratio)
            return (
                cls._chart_3d_fan_plot_point(
                    chart,
                    148.0 + 574.0 * projected,
                    57.0 - 46.0 * projected,
                ),
                cls._chart_3d_fan_plot_point(
                    chart,
                    148.0 + 575.0 * projected,
                    233.0 + 64.0 * projected,
                ),
            )
        if cls._uses_parallel_percent_3d_bar_projection(chart):
            projected = cls._chart_3d_parallel_percent_ratio(ratio)
            horizontal = (
                0.90303511 * ratio
                - 0.01745248 * ratio ** 2
                + 0.11441737 * ratio ** 3
            )
            vertical = (
                0.83312604 * ratio
                + 0.09639303 * ratio ** 2
                + 0.07048093 * ratio ** 3
            )
            return (
                cls._chart_3d_parallel_percent_plot_point(
                    chart,
                    123.0 + 656.0 * projected,
                    3.0 + 35.0 * projected,
                ),
                cls._chart_3d_parallel_percent_plot_point(
                    chart,
                    134.0 + 620.0 * horizontal,
                    201.0 + 105.0 * vertical,
                ),
            )
        (
            _, _, wall_start_scale, wall_value_scale, _,
        ) = cls._chart_3d_bar_x_scales(chart)
        top_x = (
            97.1 + 550.33333 * ratio - 24.0 * ratio ** 2
            + 90.66667 * ratio ** 3
        )
        top_y = 2.02857 + 30.57143 * ratio + 3.42857 * ratio ** 2
        bottom_x = (
            109.97143 + 544.09524 * ratio - 75.42857 * ratio ** 2
            + 117.33333 * ratio ** 3
        )
        bottom_y = (
            175.97143 + 72.09524 * ratio - 19.42857 * ratio ** 2
            + 21.33333 * ratio ** 3
        )
        top_start_x = 97.1
        bottom_start_x = 109.97143
        width_growth = max(0.0, chart.plot_width - 300.54)
        height_growth = max(0.0, chart.plot_height - 91.07)
        middle_x_adjustment = (
            width_growth * 0.2915 * ratio ** 2 * (1.0 - ratio)
        )
        bottom_y_adjustment = height_growth * (
            0.0083 * ratio
            + 0.0830 * ratio * (1.0 - ratio)
        )
        points = (
            (
                chart.plot_x + top_start_x * wall_start_scale
                + (top_x - top_start_x) * wall_value_scale
                + middle_x_adjustment,
                chart.plot_y + chart.plot_height * top_y / 254.0,
            ),
            (
                chart.plot_x + bottom_start_x * wall_start_scale
                + (bottom_x - bottom_start_x) * wall_value_scale
                + middle_x_adjustment,
                chart.plot_y + chart.plot_height * bottom_y / 254.0
                + height_growth * 0.1383 * (1.0 - ratio)
                + bottom_y_adjustment,
            ),
        )
        if cls._uses_expanded_3d_bar_date_axis(chart):
            horizontal_shift = 1.0 + 4.1 * ratio
            return (
                (points[0][0] + horizontal_shift, points[0][1]),
                (
                    points[1][0] + horizontal_shift,
                    points[1][1] - 3.0,
                ),
            )
        return points

    @classmethod
    def _chart_3d_bar_wall_hook_point(cls, chart, ratio, bottom):
        if cls._uses_high_perspective_3d_bar_projection(chart):
            return bottom
        if cls._uses_fan_3d_bar_projection(chart):
            projected = cls._chart_3d_fan_ratio(ratio)
            return cls._chart_3d_fan_plot_point(
                chart,
                121.0 + 589.0 * projected,
                258.0 + 47.0 * projected,
            )
        if cls._uses_parallel_percent_3d_bar_projection(chart):
            ratio = max(0.0, min(1.0, float(ratio)))
            horizontal = (
                0.90303511 * ratio
                - 0.01745248 * ratio ** 2
                + 0.11441737 * ratio ** 3
            )
            vertical = (
                0.83312604 * ratio
                + 0.09639303 * ratio ** 2
                + 0.07048093 * ratio ** 3
            )
            return cls._chart_3d_parallel_percent_plot_point(
                chart,
                115.2 + 640.2 * horizontal,
                225.6 + 80.4 * vertical,
            )
        return (
            bottom[0] - chart.plot_width * 7.0 / 836.0,
            bottom[1] + chart.plot_height * 8.0 / 254.0,
        )

    @classmethod
    def _chart_3d_bar_category_axis_point(cls, chart, category_position):
        position = float(category_position)
        if cls._uses_high_perspective_3d_bar_projection(chart):
            outer_top, *_, outer_bottom = (
                cls._chart_3d_high_perspective_bar_wall_corners(chart)
            )
            ratio = max(0.0, min(1.0, position / 6.0))
            curve = ratio * (1.0 - ratio)
            width_scale = chart.width / 384.36
            height_scale = chart.height / 195.36
            return (
                outer_bottom[0]
                + (outer_top[0] - outer_bottom[0]) * ratio
                + (0.28 + 2.1 * curve) * width_scale,
                outer_bottom[1]
                + (outer_top[1] - outer_bottom[1]) * ratio
                + 6.5 * curve * height_scale,
            )
        if cls._uses_fan_3d_bar_projection(chart):
            return cls._chart_3d_fan_plot_point(
                chart, 121.8, 258.4 - 33.86 * position
            )
        if cls._uses_parallel_percent_3d_bar_projection(chart):
            return cls._chart_3d_parallel_percent_plot_point(
                chart,
                115.6005 - 2.63975 * position
                - 0.015625 * position ** 2,
                226.101 - 34.537 * position
                - 0.28 * position ** 2,
            )
        pixel_x = 112.9 - 2.3 * position - 0.025 * position ** 2
        pixel_y = 176.4 - 27.15 * position - 0.225 * position ** 2
        *_, category_x_scale = cls._chart_3d_bar_x_scales(chart)
        height_growth = max(0.0, chart.plot_height - 91.07)
        vertical_shift = height_growth * (
            0.0130 + (2.5 - position) * 0.0246
        )
        point = (
            chart.plot_x + pixel_x * category_x_scale,
            chart.plot_y + chart.plot_height * pixel_y / 254.0
            + vertical_shift,
        )
        if cls._uses_expanded_3d_bar_date_axis(chart):
            return (point[0] - 2.2, point[1])
        return point

    @staticmethod
    def _uses_expanded_3d_bar_date_axis(chart):
        return (
            len(chart.axis_categories) > len(chart.categories)
            and bool(chart.axis_categories)
        )

    @classmethod
    def _chart_3d_bar_data_category_position(cls, chart, category_index):
        category_count = max(1, len(chart.categories))
        if cls._uses_expanded_3d_bar_date_axis(chart):
            axis_start = float(chart.axis_categories[0][1])
            return (
                float(chart.category_positions[category_index]) - axis_start
            ) * 6.0
        return category_index * 5.0 / max(1, category_count - 1)

    @classmethod
    def _chart_3d_bar_axis_category_position(
        cls, chart, category_index, normalized_position,
    ):
        if cls._uses_expanded_3d_bar_date_axis(chart):
            axis_start = float(chart.axis_categories[0][1])
            return (float(normalized_position) - axis_start) * 6.0
        return float(category_index)

    @classmethod
    def _chart_3d_bar_prisms(cls, chart):
        if chart.grouping in ('stacked', 'percentStacked'):
            return cls._chart_3d_stacked_bar_prisms(chart)
        if cls._uses_high_perspective_3d_bar_projection(chart):
            return cls._chart_3d_high_perspective_bar_prisms(chart)

        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        value_range = max(1.0, chart.value_max - chart.value_min)
        (
            start_x_scale, value_x_scale, _, _, _,
        ) = cls._chart_3d_bar_x_scales(chart)
        (
            base_thickness_y_scale,
            thickness_y_growth,
            slope_y_scale,
        ) = cls._chart_3d_bar_y_scales(chart)
        center_y_scale = chart.plot_height / 254.0
        width_growth = max(0.0, chart.plot_width - 300.54)
        height_growth = max(0.0, chart.plot_height - 91.07)
        expanded_date_axis = cls._uses_expanded_3d_bar_date_axis(chart)
        prisms = []

        for category_index in reversed(range(category_count)):
            category = cls._chart_3d_bar_data_category_position(
                chart, category_index
            )
            for series_index in range(series_count):
                depth = max(0.0, series_count - 1.0 - series_index)
                start_x = (
                    113.98413 - 2.21786 * category
                    + 0.78571 * depth - 0.02976 * category ** 2
                    - 0.01429 * category * depth
                    - 0.08333 * depth ** 2
                )
                start_y = (
                    154.67561 - 27.17562 * category
                    + 6.13912 * depth - 0.25502 * category ** 2
                    + 0.07209 * category * depth
                    - 0.00167 * depth ** 2
                )
                value = float(
                    chart.series[series_index].values[category_index]
                )
                ratio = max(0.0, min(
                    1.0, (value - chart.value_min) / value_range
                ))
                delta_x = (
                    ratio * (
                        454.40213 + 4.11611 * category
                        - 2.61497 * depth
                    )
                    + 132.51470 * ratio ** 2
                )
                delta_y = (
                    ratio * (
                        55.70712 - 5.78819 * category
                        + 1.92851 * depth
                    )
                    + 9.20613 * ratio ** 2
                )
                half_height = (
                    (1.35 if expanded_date_axis else 2.75) + 0.2 * ratio
                )
                depth_x = 2.2
                depth_y = -1.8
                thickness_y_scale = (
                    base_thickness_y_scale
                    + thickness_y_growth * max(0.43, 1.0 - 0.285 * depth)
                )
                size_factor = min(1.0, height_growth / 72.33)
                series_plane_shift = size_factor * (
                    0.375 * depth + 0.125 * depth ** 2
                )
                category_shift = height_growth * (
                    0.0069 + 0.0216 * (5.0 - category)
                ) + series_plane_shift
                start_center_x = chart.plot_x + start_x * start_x_scale
                if expanded_date_axis:
                    start_center_x -= 2.0
                perspective_extension = (
                    width_growth * 0.0471
                    * 4.0 * ratio * (1.0 - ratio)
                )
                end_center_x = (
                    start_center_x + delta_x * value_x_scale
                    + perspective_extension
                )
                if expanded_date_axis:
                    end_center_x += 2.0 + 4.6 * ratio
                start_center_y = (
                    chart.plot_y + start_y * center_y_scale
                    + category_shift
                )
                if expanded_date_axis:
                    start_center_y += 6.75 - 1.6 * depth
                end_center_y = start_center_y + delta_y * slope_y_scale
                half_height_points = half_height * thickness_y_scale

                start_top = (
                    start_center_x, start_center_y - half_height_points,
                )
                start_bottom = (
                    start_center_x, start_center_y + half_height_points,
                )
                end_top = (
                    end_center_x, end_center_y - half_height_points,
                )
                end_bottom = (
                    end_center_x, end_center_y + half_height_points,
                )
                depth_vector = (
                    depth_x * start_x_scale,
                    depth_y * thickness_y_scale,
                )

                def shifted(point):
                    return (
                        point[0] + depth_vector[0],
                        point[1] + depth_vector[1],
                    )

                front = (
                    start_top, end_top, end_bottom, start_bottom,
                )
                side = (
                    end_top, shifted(end_top),
                    shifted(end_bottom), end_bottom,
                )
                top = (
                    start_top, shifted(start_top),
                    shifted(end_top), end_top,
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @classmethod
    def _chart_3d_high_perspective_bar_prisms(cls, chart):
        category_count = max(1, len(chart.categories))
        series_count = max(1, len(chart.series))
        value_range = max(1.0, chart.value_max - chart.value_min)
        half_height = chart.height * 0.0022
        width_scale = chart.width / 384.36
        height_scale = chart.height / 195.36
        prisms = []

        order = sorted(
            range(category_count),
            key=lambda index: chart.category_positions[index],
            reverse=True,
        )
        for category_index in order:
            category = float(chart.category_positions[category_index])
            category = max(0.0, min(1.0, category))
            category_curve = category * (1.0 - category)
            category_x_shift = (
                4.5 * category_curve - 1.3 * category ** 8
            ) * width_scale
            category_y_shift = (
                -0.75 + 3.7 * category - 6.5 * category ** 8
            ) * height_scale

            for series_index in range(series_count):
                depth = max(0.0, series_count - 1.0 - series_index)
                value = float(
                    chart.series[series_index].values[category_index]
                )
                ratio = max(0.0, min(
                    1.0, (value - chart.value_min) / value_range
                ))
                start = cls._chart_3d_high_perspective_bar_surface_point(
                    chart, category, 0.0
                )
                end = cls._chart_3d_high_perspective_bar_surface_point(
                    chart, category, ratio
                )
                series_x_shift = depth * 0.55 * width_scale
                series_y_shift = depth * (
                    0.95 + 1.2 * (1.0 - category)
                ) * height_scale
                shift_x = category_x_shift + series_x_shift
                shift_y = category_y_shift + series_y_shift

                start_top = (
                    start[0] + shift_x,
                    start[1] + shift_y - half_height,
                )
                start_bottom = (
                    start[0] + shift_x,
                    start[1] + shift_y + half_height,
                )
                end_top = (
                    end[0] + shift_x,
                    end[1] + shift_y - half_height,
                )
                end_bottom = (
                    end[0] + shift_x,
                    end[1] + shift_y + half_height,
                )
                depth_vector = (
                    (0.6 + 3.2 * (1.0 - category)) * width_scale,
                    (-0.2 - 0.6 * (1.0 - category)) * height_scale,
                )

                def shifted(point):
                    return (
                        point[0] + depth_vector[0],
                        point[1] + depth_vector[1],
                    )

                front = (start_top, end_top, end_bottom, start_bottom)
                side = (
                    end_top, shifted(end_top),
                    shifted(end_bottom), end_bottom,
                )
                top = (
                    start_top, shifted(start_top),
                    shifted(end_top), end_top,
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @classmethod
    def _chart_3d_bar_wall_polygons(cls, chart):
        wall_points = tuple(
            cls._chart_3d_bar_wall_points(chart, index / 40.0)
            for index in range(41)
        )
        top = tuple(points[0] for points in wall_points)
        bottom = tuple(points[1] for points in wall_points)
        if cls._uses_high_perspective_3d_bar_projection(chart):
            outer_top, top_left, *_, bottom_left, outer_bottom = (
                cls._chart_3d_high_perspective_bar_wall_corners(chart)
            )
            return (
                top + tuple(reversed(bottom)),
                (outer_top, top_left, bottom_left, outer_bottom),
            )
        hooks = tuple(
            cls._chart_3d_bar_wall_hook_point(chart, index / 40.0, point)
            for index, point in enumerate(bottom)
        )
        return (
            top + tuple(reversed(bottom)),
            bottom + tuple(reversed(hooks)),
        )

    @classmethod
    def _chart_3d_high_perspective_bar_floor_polygon(cls, chart):
        inner = []
        bottom_edge = []
        for index in range(41):
            ratio = index / 40.0
            top, bottom = cls._chart_3d_bar_wall_points(chart, ratio)
            projected = cls._chart_3d_high_perspective_bar_value_ratio(ratio)
            thickness = 0.055 + 0.035 * projected
            inner.append((
                bottom[0] + (top[0] - bottom[0]) * thickness,
                bottom[1] + (top[1] - bottom[1]) * thickness,
            ))
            bottom_edge.append(bottom)
        return tuple(inner) + tuple(reversed(bottom_edge))

    @classmethod
    def _chart_3d_fan_stacked_bar_prisms(cls, chart):
        category_count = max(1, len(chart.categories))
        prisms = []

        for category_index in reversed(range(category_count)):
            category = (
                category_index * 5.0 / max(1, category_count - 1)
            )
            start_center_y = 239.64286 - 33.45714 * category
            end_center_y = 280.29762 - 49.38571 * category
            values = [
                float(series.values[category_index])
                for series in chart.series
            ]
            total = sum(max(0.0, value) for value in values)
            cumulative = 0.0

            for series_index, value in enumerate(values):
                start_ratio = cumulative / max(1.0, total)
                cumulative += max(0.0, value)
                end_ratio = cumulative / max(1.0, total)

                def projected_edge(ratio):
                    projected = cls._chart_3d_fan_ratio(ratio)
                    center_y = (
                        start_center_y
                        + (end_center_y - start_center_y) * projected
                    )
                    half_height = 6.5 + (
                        2.5 + 0.2 * (5.0 - category)
                    ) * projected
                    center_x = 130.0 + 584.0 * projected
                    return (
                        cls._chart_3d_fan_plot_point(
                            chart, center_x, center_y - half_height
                        ),
                        cls._chart_3d_fan_plot_point(
                            chart, center_x, center_y + half_height
                        ),
                    )

                start_top, start_bottom = projected_edge(start_ratio)
                end_top, end_bottom = projected_edge(end_ratio)
                depth_vector = (
                    chart.plot_width * 5.0 / 895.0,
                    -chart.plot_height * 2.0 / 313.0,
                )

                def shifted(point):
                    return (
                        point[0] + depth_vector[0],
                        point[1] + depth_vector[1],
                    )

                front = (start_top, end_top, end_bottom, start_bottom)
                side = (
                    end_top, shifted(end_top),
                    shifted(end_bottom), end_bottom,
                )
                top = (
                    start_top, shifted(start_top),
                    shifted(end_top), end_top,
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @classmethod
    def _chart_3d_parallel_percent_stacked_bar_prisms(cls, chart):
        category_count = max(1, len(chart.categories))
        prisms = []

        for category_index in reversed(range(category_count)):
            category = (
                category_index * 5.0 / max(1, category_count - 1)
            )
            start_center_x = 120.0 - 2.6 * category
            end_center_x = 756.0 + 3.0 * category
            start_center_y = 204.59524 - 35.57143 * category
            end_center_y = 279.19048 - 42.84286 * category
            values = [
                float(series.values[category_index])
                for series in chart.series
            ]
            total = sum(max(0.0, value) for value in values)
            cumulative = 0.0

            for series_index, value in enumerate(values):
                start_ratio = cumulative / max(1.0, total)
                cumulative += max(0.0, value)
                end_ratio = cumulative / max(1.0, total)

                def projected_edge(ratio):
                    projected = cls._chart_3d_parallel_percent_ratio(ratio)
                    center_x = (
                        start_center_x
                        + (end_center_x - start_center_x) * projected
                    )
                    center_y = (
                        start_center_y
                        + (end_center_y - start_center_y) * projected
                    )
                    half_height = 8.5 + 1.5 * projected
                    return (
                        cls._chart_3d_parallel_percent_plot_point(
                            chart, center_x, center_y - half_height
                        ),
                        cls._chart_3d_parallel_percent_plot_point(
                            chart, center_x, center_y + half_height
                        ),
                    )

                start_top, start_bottom = projected_edge(start_ratio)
                end_top, end_bottom = projected_edge(end_ratio)
                depth_vector = (
                    chart.plot_width * 4.0 / 895.0,
                    -chart.plot_height * 2.0 / 312.0,
                )

                def shifted(point):
                    return (
                        point[0] + depth_vector[0],
                        point[1] + depth_vector[1],
                    )

                front = (start_top, end_top, end_bottom, start_bottom)
                side = (
                    end_top, shifted(end_top),
                    shifted(end_bottom), end_bottom,
                )
                top = (
                    start_top, shifted(start_top),
                    shifted(end_top), end_top,
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @classmethod
    def _chart_3d_stacked_bar_prisms(cls, chart):
        if cls._uses_fan_3d_bar_projection(chart):
            return cls._chart_3d_fan_stacked_bar_prisms(chart)
        if cls._uses_parallel_percent_3d_bar_projection(chart):
            return cls._chart_3d_parallel_percent_stacked_bar_prisms(chart)

        category_count = max(1, len(chart.categories))
        value_range = max(1.0, chart.value_max - chart.value_min)
        (
            start_x_scale, value_x_scale, _, _, _,
        ) = cls._chart_3d_bar_x_scales(chart)
        (
            base_thickness_y_scale,
            thickness_y_growth,
            slope_y_scale,
        ) = cls._chart_3d_bar_y_scales(chart)
        center_y_scale = chart.plot_height / 254.0
        width_growth = max(0.0, chart.plot_width - 300.54)
        height_growth = max(0.0, chart.plot_height - 91.07)
        depth = 1.0
        thickness_y_scale = (
            base_thickness_y_scale + thickness_y_growth * 0.715
        )
        size_factor = min(1.0, height_growth / 72.33)
        prisms = []

        for category_index in reversed(range(category_count)):
            category = category_index * 5.0 / max(1, category_count - 1)
            start_x = (
                113.98413 - 2.21786 * category
                + 0.78571 * depth - 0.02976 * category ** 2
                - 0.01429 * category * depth
                - 0.08333 * depth ** 2
            )
            start_y = (
                154.67561 - 27.17562 * category
                + 6.13912 * depth - 0.25502 * category ** 2
                + 0.07209 * category * depth
                - 0.00167 * depth ** 2
            )
            category_shift = height_growth * (
                0.0069 + 0.0216 * (5.0 - category)
            ) + size_factor * 0.5 + 0.5
            start_center_x = (
                chart.plot_x + start_x * start_x_scale
                - 0.1 * category
            )
            start_center_y = (
                chart.plot_y + start_y * center_y_scale
                + category_shift
            )
            values = [
                float(series.values[category_index])
                for series in chart.series
            ]
            total = sum(max(0.0, value) for value in values)
            cumulative = 0.0
            stack_ratio = (
                1.0
                if chart.grouping == 'percentStacked'
                else max(0.0, min(
                    1.0, (total - chart.value_min) / value_range
                ))
            )
            half_height = (
                2.75 + 0.2 * stack_ratio
            ) * thickness_y_scale * 2.2

            for series_index, value in enumerate(values):
                segment_start = cumulative
                cumulative += max(0.0, value)
                if chart.grouping == 'percentStacked':
                    start_ratio = segment_start / max(1.0, total)
                    end_ratio = cumulative / max(1.0, total)
                else:
                    start_ratio = max(0.0, min(
                        1.0, (segment_start - chart.value_min) / value_range
                    ))
                    end_ratio = max(0.0, min(
                        1.0, (cumulative - chart.value_min) / value_range
                    ))

                def projected_offset(ratio):
                    delta_x = (
                        ratio * (
                            454.40213 + 4.11611 * category
                            - 2.61497 * depth
                        )
                        + 132.51470 * ratio ** 2
                    )
                    delta_y = (
                        ratio * (
                            55.70712 - 5.78819 * category
                            + 1.92851 * depth
                        )
                        + 9.20613 * ratio ** 2
                    )
                    perspective_extension = (
                        width_growth * 0.0471
                        * 4.0 * ratio * (1.0 - ratio)
                    )
                    stacked_extension = (
                        4.3 * ratio / (ratio + 0.1)
                        + 0.6 * ratio ** 6
                        if ratio > 0.0 else 0.0
                    )
                    return (
                        delta_x * value_x_scale
                        + perspective_extension + stacked_extension,
                        delta_y * slope_y_scale - 1.6 * ratio ** 2,
                    )

                start_offset = projected_offset(start_ratio)
                end_offset = projected_offset(end_ratio)
                segment_start_center = (
                    start_center_x + start_offset[0],
                    start_center_y + start_offset[1],
                )
                segment_end_center = (
                    start_center_x + end_offset[0],
                    start_center_y + end_offset[1],
                )
                start_top = (
                    segment_start_center[0],
                    segment_start_center[1] - half_height,
                )
                start_bottom = (
                    segment_start_center[0],
                    segment_start_center[1] + half_height,
                )
                end_top = (
                    segment_end_center[0],
                    segment_end_center[1] - half_height,
                )
                end_bottom = (
                    segment_end_center[0],
                    segment_end_center[1] + half_height,
                )
                depth_vector = (
                    2.2 * start_x_scale,
                    -1.8 * thickness_y_scale,
                )

                def shifted(point):
                    return (
                        point[0] + depth_vector[0],
                        point[1] + depth_vector[1],
                    )

                front = (start_top, end_top, end_bottom, start_bottom)
                side = (
                    end_top, shifted(end_top),
                    shifted(end_bottom), end_bottom,
                )
                top = (
                    start_top, shifted(start_top),
                    shifted(end_top), end_top,
                )
                prisms.append((
                    series_index, category_index, front, side, top,
                ))
        return tuple(prisms)

    @staticmethod
    def _chart_3d_bar_legend_series(chart):
        if chart.grouping in ('stacked', 'percentStacked'):
            return tuple(chart.series)
        return tuple(reversed(chart.series))

    def _draw_3d_bar_chart(
        self, canvas, chart, axis_font, axis_title_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        import math

        tick_count = max(1, int(round(
            (chart.value_max - chart.value_min) / chart.value_step
        )))
        value_grid_paint = self._paint(
            'FF' + chart.value_major_gridline_color,
            stroke_width=0.75,
        )

        back_wall, side_wall = self._chart_3d_bar_wall_polygons(chart)
        wall_color_matrix = None
        if self._uses_high_perspective_3d_bar_projection(chart):
            wall_color_matrix = (
                1.075, 0.0, 0.0, 0.0, 0.0,
                0.0, 1.088, 0.0, 0.0, 0.0,
                0.0, 0.0, 1.107, 0.0, 0.0,
                0.0, 0.0, 0.0, 1.0, 0.0,
            )
        if chart.back_wall_fill_image:
            back_wall_paint = self._chart_bitmap_paint(
                chart.back_wall_fill_image, 0.36,
                color_matrix=wall_color_matrix,
            )
            canvas.drawPath(
                self._polygon_path(back_wall),
                back_wall_paint,
            )
        if chart.side_wall_fill_image:
            side_wall_paint = self._chart_bitmap_paint(
                chart.side_wall_fill_image, 0.36,
                color_matrix=wall_color_matrix,
            )
            canvas.drawPath(
                self._polygon_path(side_wall),
                side_wall_paint,
            )
            canvas.drawPath(
                self._polygon_path(side_wall),
                self._paint('33000000', fill=True),
            )
        if self._uses_high_perspective_3d_bar_projection(chart):
            floor = self._chart_3d_high_perspective_bar_floor_polygon(chart)
            canvas.drawPath(
                self._polygon_path(floor),
                self._paint('30000000', fill=True),
            )

        if chart.show_minor_horizontal_gridlines:
            minor_grid_paint = self._paint(
                'FF' + chart.minor_gridline_color,
                stroke_width=0.75,
            )
            for major_index in range(tick_count):
                for minor_index in range(1, 5):
                    ratio = (
                        major_index + minor_index / 5.0
                    ) / tick_count
                    top, bottom = self._chart_3d_bar_wall_points(
                        chart, ratio
                    )
                    canvas.drawLine(*top, *bottom, minor_grid_paint)
                    if not self._uses_high_perspective_3d_bar_projection(
                        chart
                    ):
                        hook = self._chart_3d_bar_wall_hook_point(
                            chart, ratio, bottom
                        )
                        canvas.drawLine(*bottom, *hook, minor_grid_paint)

        for index in range(tick_count + 1):
            ratio = index / tick_count
            top, bottom = self._chart_3d_bar_wall_points(chart, ratio)
            canvas.drawLine(*top, *bottom, value_grid_paint)
            if (
                not self._uses_high_perspective_3d_bar_projection(chart)
                and (
                    index > 0
                    or self._uses_fan_3d_bar_projection(chart)
                    or self._uses_parallel_percent_3d_bar_projection(chart)
                )
            ):
                hook = self._chart_3d_bar_wall_hook_point(
                    chart, ratio, bottom
                )
                canvas.drawLine(*bottom, *hook, value_grid_paint)

        for series_index, _, front, side, top in (
            self._chart_3d_bar_prisms(chart)
        ):
            color = chart.series[series_index].color
            canvas.drawPath(
                self._polygon_path(side),
                self._paint(
                    'FF' + self._scale_chart_color(color, 0.70),
                    fill=True,
                ),
            )
            canvas.drawPath(
                self._polygon_path(front),
                self._paint('FF' + color, fill=True),
            )
            canvas.drawPath(
                self._polygon_path(top),
                self._paint(
                    'FF' + self._scale_chart_color(color, 0.76),
                    fill=True,
                ),
            )

        category_count = max(1, len(chart.categories))
        category_axis = self._skia.Path()
        for boundary_index in reversed(range(category_count + 1)):
            point = self._chart_3d_bar_category_axis_point(
                chart, boundary_index * 6.0 / category_count
            )
            if boundary_index == category_count:
                category_axis.moveTo(*point)
            else:
                category_axis.lineTo(*point)
        canvas.drawPath(category_axis, grid_paint)
        if (
            self._uses_fan_3d_bar_projection(chart)
            or self._uses_parallel_percent_3d_bar_projection(chart)
        ):
            point = self._chart_3d_bar_category_axis_point(chart, 6.0)
            canvas.drawLine(
                point[0] - 2.8, point[1],
                point[0], point[1], grid_paint,
            )

        if chart.value_axis_color:
            value_axis_paint = self._paint(
                'FF' + chart.value_axis_color, stroke_width=0.75
            )
            value_axis = self._skia.Path()
            for index in range(41):
                _, point = self._chart_3d_bar_wall_points(
                    chart, index / 40.0
                )
                if index == 0:
                    value_axis.moveTo(*point)
                else:
                    value_axis.lineTo(*point)
            canvas.drawPath(value_axis, value_axis_paint)

        axis_metrics = axis_font.getMetrics()
        category_label_stride = 1
        if len(chart.axis_categories) > 1:
            _, first_position = chart.axis_categories[0]
            _, second_position = chart.axis_categories[1]
            first_axis_position = self._chart_3d_bar_axis_category_position(
                chart, 0, first_position
            )
            second_axis_position = self._chart_3d_bar_axis_category_position(
                chart, 1, second_position
            )
            first_point = self._chart_3d_bar_category_axis_point(
                chart, first_axis_position
            )
            second_point = self._chart_3d_bar_category_axis_point(
                chart, second_axis_position
            )
            projected_spacing = abs(second_point[1] - first_point[1])
            label_height = axis_metrics.fDescent - axis_metrics.fAscent
            if projected_spacing < label_height + 3.0:
                category_label_stride = max(
                    2,
                    int(math.ceil(
                        (label_height + 3.0) / max(0.01, projected_spacing)
                    )),
                )
        for category_index, (label, position) in enumerate(
            chart.axis_categories
        ):
            if category_index % category_label_stride:
                continue
            axis_position = self._chart_3d_bar_axis_category_position(
                chart, category_index, position
            )
            point = self._chart_3d_bar_category_axis_point(
                chart, axis_position
            )
            canvas.drawLine(
                point[0] - 2.8, point[1],
                point[0], point[1], grid_paint,
            )
            label_width = axis_font.measureText(label)
            baseline = point[1] - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0 - 5.0
            if self._uses_parallel_percent_3d_bar_projection(chart):
                baseline -= 1.36 + 0.0975 * category_index
            elif self._uses_high_perspective_3d_bar_projection(chart):
                baseline += 3.2
            elif self._uses_expanded_3d_bar_date_axis(chart):
                baseline += 4.0
            canvas.drawString(
                label, point[0] - 8.75 - label_width, baseline,
                axis_font, category_text_paint,
            )

        for index in range(tick_count + 1):
            ratio = index / tick_count
            _, bottom = self._chart_3d_bar_wall_points(chart, ratio)
            label_point = self._chart_3d_bar_wall_hook_point(
                chart, ratio, bottom
            )
            value = chart.value_min + chart.value_step * index
            display_value = value / max(
                1e-12, chart.value_axis_display_unit
            )
            label = self._chart_value_label(
                display_value, chart.axis_number_format
            )
            label_width = axis_font.measureText(label)
            if self._uses_high_perspective_3d_bar_projection(chart):
                label_shift = 0.0
                label_x_adjustment = (
                    -1.22 + 3.4 * ratio - 3.0 * ratio ** 3
                )
                label_top = bottom[1] + (
                    7.24 if index == 0 else 10.08 - 0.77 * ratio
                )
            else:
                label_shift = 0.0
                if not self._uses_fan_3d_bar_projection(chart):
                    label_shift = (
                        chart.plot_width * 45.0 * ratio
                        * (1.0 - ratio) ** 2 / 836.0
                    )
                label_x_adjustment = 0.0
                if self._uses_expanded_3d_bar_date_axis(chart):
                    label_x_adjustment = -0.66 + 2.14 * ratio
                label_top = label_point[1] + 6.45 + ratio
            canvas.drawString(
                label,
                label_point[0] - label_width / 2.0 - label_shift
                + label_x_adjustment,
                label_top - axis_metrics.fAscent,
                axis_font, value_text_paint,
            )

        if chart.value_axis_display_unit_label:
            label_metrics = axis_title_font.getMetrics()
            label_top = chart.y + chart.height * 0.7497
            canvas.drawString(
                chart.value_axis_display_unit_label,
                chart.x + chart.width * 0.6725,
                label_top - label_metrics.fAscent,
                axis_title_font, value_text_paint,
            )

        self._draw_3d_chart_title_and_legend(
            canvas, chart, axis_font, title_font,
            legend_text_paint, title_paint,
            legend_series=self._chart_3d_bar_legend_series(chart),
            legend_offset_x=-1.24,
        )

