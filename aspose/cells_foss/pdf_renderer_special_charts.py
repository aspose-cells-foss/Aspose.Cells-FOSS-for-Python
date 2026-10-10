"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRendererSpecialChartsMixin:
    def _bubble_3d_paint(self, color, x, y, radius):
        channels = tuple(
            int(color[index:index + 2], 16) for index in (0, 2, 4)
        )

        def blend(target, amount):
            return self._skia.ColorSetARGB(
                255,
                *(round(value + (target - value) * amount)
                  for value in channels),
            )

        paint = self._skia.Paint(AntiAlias=True)
        paint.setShader(self._skia.GradientShader.MakeRadial(
            self._skia.Point(x - radius * 0.16, y - radius * 0.2),
            radius * 1.16,
            (
                blend(255, 0.5),
                blend(255, 0.25),
                blend(0, 0.0),
                blend(0, 0.32),
                blend(0, 0.66),
            ),
            (0.0, 0.18, 0.5, 0.82, 1.0),
            self._skia.TileMode.kClamp,
        ))
        return paint

    @classmethod
    def _map_features(cls):
        if cls._MAP_FEATURES is None:
            resource = os.path.join(
                os.path.dirname(__file__),
                'ne_110m_admin_0_countries.geojson',
            )
            with open(resource, 'r', encoding='utf-8') as stream:
                cls._MAP_FEATURES = tuple(json.load(stream)['features'])
        return cls._MAP_FEATURES

    def _draw_2d_sunburst_chart(
        self, canvas, chart, axis_font, title_font, title_paint
    ):
        if not chart.sunburst_paths:
            return

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font,
            title_paint,
        )

        totals = {}
        children = {}
        for path, value in chart.sunburst_paths:
            numeric = max(0.0, float(value))
            for depth in range(1, len(path) + 1):
                prefix = tuple(path[:depth])
                totals[prefix] = totals.get(prefix, 0.0) + numeric
                parent = prefix[:-1]
                children.setdefault(parent, [])
                if prefix not in children[parent]:
                    children[parent].append(prefix)
        root_total = sum(totals[node] for node in children.get((), ()))
        if root_total <= 0.0:
            return

        center_x = chart.x + chart.width * 0.5
        center_y = chart.y + chart.height * 0.56
        outer_radius = min(chart.width * 0.213, chart.height * 0.42)
        max_depth = max(len(path) for path, _ in chart.sunburst_paths)
        inner_radius = outer_radius * 0.26
        ring_width = (outer_radius - inner_radius) / max(1, max_depth)
        colors = ('156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E')
        label_font = self._font(chart.axis_font_resolution, 8.0)
        label_paint = self._paint('FFFFFFFF', fill=True)
        border_paint = self._paint('FFFFFFFF', stroke_width=1.0)
        top_color_indexes = {
            node: index for index, node in enumerate(children.get((), ()))
        }

        def draw_children(parent, start_angle, sweep_angle, top_index):
            nodes = sorted(
                children.get(parent, ()),
                key=lambda node: totals[node],
                reverse=True,
            )
            parent_total = sum(totals[node] for node in nodes)
            angle = start_angle
            for child_index, node in enumerate(nodes):
                node_sweep = (
                    sweep_angle * totals[node] / parent_total
                    if parent_total > 0.0 else 0.0
                )
                depth = len(node)
                color_index = (
                    top_color_indexes[node] if depth == 1 else top_index
                )
                base_color = colors[color_index % len(colors)]
                inner = inner_radius + (depth - 1) * ring_width
                outer = inner + ring_width
                outer_rect = self._skia.Rect.MakeXYWH(
                    center_x - outer, center_y - outer,
                    outer * 2.0, outer * 2.0,
                )
                inner_rect = self._skia.Rect.MakeXYWH(
                    center_x - inner, center_y - inner,
                    inner * 2.0, inner * 2.0,
                )
                path = self._skia.Path()
                path.arcTo(outer_rect, angle, node_sweep, True)
                path.arcTo(
                    inner_rect, angle + node_sweep, -node_sweep, False
                )
                path.close()
                canvas.drawPath(
                    path, self._paint('FF' + base_color, fill=True)
                )
                canvas.drawPath(path, border_paint)

                label = str(node[-1])
                middle_angle = angle + node_sweep / 2.0
                middle_radius = (inner + outer) / 2.0
                available = math.radians(abs(node_sweep)) * middle_radius
                text_width = label_font.measureText(label)
                if available >= text_width + 3.0 and ring_width >= 9.0:
                    radians = math.radians(middle_angle)
                    label_x = center_x + middle_radius * math.cos(radians)
                    label_y = center_y + middle_radius * math.sin(radians)
                    rotation = middle_angle
                    normalized = rotation % 360.0
                    if 90.0 < normalized < 270.0:
                        rotation += 180.0
                    metrics = label_font.getMetrics()
                    canvas.save()
                    canvas.translate(label_x, label_y)
                    canvas.rotate(rotation)
                    canvas.drawString(
                        label,
                        -text_width / 2.0,
                        -(metrics.fAscent + metrics.fDescent) / 2.0,
                        label_font,
                        label_paint,
                    )
                    canvas.restore()

                draw_children(node, angle, node_sweep, color_index)
                angle += node_sweep

        draw_children((), -90.0, 360.0, 0)

    @staticmethod
    def _treemap_rectangles(items, x, y, width, height):
        """Lay out ordered, weighted nodes using Excel-like hierarchical strips."""
        weighted = [
            (key, max(0.0, float(value)))
            for key, value in items
            if float(value) > 0.0
        ]
        weighted.sort(key=lambda item: item[1], reverse=True)
        if not weighted or width <= 0.0 or height <= 0.0:
            return ()
        result = []

        def partition(entries, left, top, available_width, available_height):
            if len(entries) == 1:
                key, value = entries[0]
                result.append((
                    key, value, left, top,
                    available_width, available_height,
                ))
                return

            total = sum(value for _, value in entries)
            if available_width >= available_height:
                key, value = entries[0]
                item_width = available_width * value / total
                result.append((
                    key, value, left, top, item_width, available_height,
                ))
                partition(
                    entries[1:], left + item_width, top,
                    available_width - item_width, available_height,
                )
                return

            row_count = (len(entries) + 1) // 2
            row = entries[:row_count]
            remainder = entries[row_count:]
            row_total = sum(value for _, value in row)
            row_height = available_height * row_total / total
            cursor = left
            for key, value in row:
                item_width = available_width * value / row_total
                result.append((
                    key, value, cursor, top, item_width, row_height,
                ))
                cursor += item_width
            if remainder:
                partition(
                    remainder, left, top + row_height,
                    available_width, available_height - row_height,
                )

        partition(weighted, x, y, width, height)
        return tuple(result)

    def _draw_2d_treemap_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        if not chart.treemap_paths:
            return

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + (
                7.5 if chart.height >= 250.0
                else 6.75 if chart.height >= 210.0
                else 4.8
            ) - title_metrics.fAscent,
            title_font,
            title_paint,
        )

        roots = []
        totals = {}
        children = {}
        for path, value in chart.treemap_paths:
            if not path:
                continue
            normalized = tuple(str(label) for label in path if str(label))
            if not normalized:
                continue
            root = normalized[0]
            if root not in roots:
                roots.append(root)
            numeric = max(0.0, float(value))
            for depth in range(1, len(normalized) + 1):
                node = normalized[:depth]
                totals[node] = totals.get(node, 0.0) + numeric
                parent = node[:-1]
                children.setdefault(parent, [])
                if node not in children[parent]:
                    children[parent].append(node)
        root_totals = [
            ((root,), totals.get((root,), 0.0))
            for root in roots
        ]
        root_rectangles = self._treemap_rectangles(
            root_totals,
            chart.plot_x, chart.plot_y,
            chart.plot_width, chart.plot_height,
        )
        if not root_rectangles:
            return

        colors = ('156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E')
        root_colors = {
            root: colors[index % len(colors)]
            for index, root in enumerate(roots)
        }
        label_font = self._font(chart.axis_font_resolution, 9.0)
        label_metrics = label_font.getMetrics()
        label_paint = self._paint('FFFFFFFF', fill=True)
        border_paint = self._paint('FFFFFFFF', stroke_width=1.0)

        def draw_leaf(label, root, item_x, item_y, item_width, item_height):
            label = str(label)
            rect = self._skia.Rect.MakeXYWH(
                item_x, item_y, item_width, item_height
            )
            canvas.drawRect(
                rect,
                self._paint('FF' + root_colors[root], fill=True),
            )
            canvas.drawRect(rect, border_paint)
            available = max(0.0, item_width - 7.0)
            if item_height < 10.0 or available < 4.0:
                return
            canvas.save()
            canvas.clipRect(rect)
            line_height = max(
                8.0, label_metrics.fDescent - label_metrics.fAscent
            )
            lines = []
            remaining = label
            while remaining and len(lines) < 2:
                take = len(remaining)
                while take > 1 and label_font.measureText(
                    remaining[:take]
                ) > available:
                    take -= 1
                lines.append(remaining[:take])
                remaining = remaining[take:]
            baseline = item_y + item_height - 4.0 - (
                len(lines) - 1
            ) * line_height
            for line in lines:
                canvas.drawString(
                    line, item_x + 3.5, baseline,
                    label_font, label_paint,
                )
                baseline += line_height
            canvas.restore()

        def draw_children(parent, root, left, top, width, height):
            nodes = children.get(parent, ())
            rectangles = self._treemap_rectangles(
                ((node, totals[node]) for node in nodes),
                left, top, width, height,
            )
            for node, _, item_x, item_y, item_width, item_height in rectangles:
                if children.get(node):
                    draw_children(
                        node, root, item_x, item_y, item_width, item_height
                    )
                else:
                    draw_leaf(
                        node[-1], root,
                        item_x, item_y, item_width, item_height,
                    )

        flat = all(len(path) == 1 for path, _ in chart.treemap_paths)
        for root_node, _, left, top, width, height in root_rectangles:
            root = root_node[0]
            if flat:
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(left, top, width, height),
                    self._paint('FF' + root_colors[root], fill=True),
                )
                continue

            draw_children(root_node, root, left, top, width, height)
            root_metrics = label_font.getMetrics()
            canvas.save()
            canvas.clipRect(
                self._skia.Rect.MakeXYWH(left, top, width, height)
            )
            canvas.drawString(
                root, left + 3.5, top + 3.0 - root_metrics.fAscent,
                label_font, label_paint,
            )
            canvas.restore()

        if chart.show_legend:
            legend_font = self._font(chart.axis_font_resolution, 8.5)
            legend_metrics = legend_font.getMetrics()
            square = 5.4
            gap = 5.0
            item_gap = 12.0
            widths = [
                square + gap + legend_font.measureText(root)
                for root in roots
            ]
            total_width = sum(widths) + item_gap * max(0, len(roots) - 1)
            cursor = chart.x + (chart.width - total_width) / 2.0
            legend_offset = (
                39.5 if chart.height >= 250.0
                else 37.4 if chart.height >= 210.0
                else 36.0
            )
            square_y = chart.y + legend_offset
            baseline = square_y + (
                square - legend_metrics.fAscent - legend_metrics.fDescent
            ) / 2.0
            for index, root in enumerate(roots):
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(cursor, square_y, square, square),
                    self._paint('FF' + root_colors[root], fill=True),
                )
                canvas.drawString(
                    root, cursor + square + gap, baseline,
                    legend_font, legend_text_paint,
                )
                cursor += widths[index] + item_gap

    @staticmethod
    def _map_fill_color(value, minimum, maximum):
        light = (217, 226, 243)
        dark = (47, 85, 151)
        ratio = (
            (float(value) - minimum) / (maximum - minimum)
            if maximum > minimum else 1.0
        )
        ratio = min(1.0, max(0.0, ratio))
        channels = tuple(
            round(start + (end - start) * ratio)
            for start, end in zip(light, dark)
        )
        return ''.join(f'{channel:02X}' for channel in channels)

    @staticmethod
    def _interpolate_chart_color(start, end, ratio):
        ratio = min(1.0, max(0.0, float(ratio)))
        start_channels = tuple(
            int(start[index:index + 2], 16) for index in (0, 2, 4)
        )
        end_channels = tuple(
            int(end[index:index + 2], 16) for index in (0, 2, 4)
        )
        return ''.join(
            f'{round(first + (last - first) * ratio):02X}'
            for first, last in zip(start_channels, end_channels)
        )

    @staticmethod
    def _map_diverging_fill_color(
        value, minimum, midpoint, maximum, colors,
    ):
        if float(value) <= midpoint:
            start, end = colors[0], colors[1]
            span_start, span_end = minimum, midpoint
        else:
            start, end = colors[1], colors[2]
            span_start, span_end = midpoint, maximum
        ratio = (
            (float(value) - span_start) / (span_end - span_start)
            if span_end > span_start else 0.0
        )
        ratio = min(1.0, max(0.0, ratio))
        start_channels = tuple(
            int(start[index:index + 2], 16) for index in (0, 2, 4)
        )
        end_channels = tuple(
            int(end[index:index + 2], 16) for index in (0, 2, 4)
        )
        return ''.join(
            f'{round(first + (last - first) * ratio):02X}'
            for first, last in zip(start_channels, end_channels)
        )

    def _draw_2d_scatter_chart(
        self, canvas, chart, axis_font, axis_title_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        effect_stops = chart.scatter_text_outline_gradient_stops

        def draw_text(text, x, y, font, normal_paint):
            if not effect_stops:
                canvas.drawString(text, x, y, font, normal_paint)
                return
            metrics = font.getMetrics()
            outline_paint = self._chart_text_paint(
                None, None, effect_stops,
                chart.scatter_text_outline_gradient_angle,
                (
                    x, y + metrics.fAscent,
                    max(1.0, font.measureText(text)),
                    max(1.0, metrics.fDescent - metrics.fAscent),
                ),
            )
            outline_paint.setStyle(self._skia.Paint.kStrokeAndFill_Style)
            outline_paint.setStrokeWidth(0.6)
            canvas.drawString(text, x, y, font, outline_paint)

        for value, ratio in self._chart_value_axis_ticks(chart):
            y = chart.plot_y + chart.plot_height * (1.0 - ratio)
            canvas.drawLine(
                chart.plot_x, y,
                chart.plot_x + chart.plot_width, y,
                grid_paint,
            )
            if not chart.scatter_vary_colors:
                display_value = value / max(
                    1e-12, chart.value_axis_display_unit
                )
                label = self._chart_value_label(
                    display_value, chart.axis_number_format
                )
                label_paint = value_text_paint
                if (
                    display_value < 0.0
                    and '[Red]' in chart.axis_number_format
                ):
                    label = f'(${abs(int(round(display_value))):,})'
                    label_paint = self._paint('FFFF0000', fill=True)
                label_width = axis_font.measureText(label)
                baseline = (
                    y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                    - 0.6
                )
                draw_text(
                    label, chart.plot_x - 10.0 - label_width, baseline,
                    axis_font, label_paint,
                )

        if chart.chart_kind == 'bubble' and chart.value_min >= 0.0:
            display_value = -chart.value_step / max(
                1e-12, chart.value_axis_display_unit
            )
            label = f'(${abs(int(round(display_value))):,})'
            label_width = axis_font.measureText(label)
            baseline = (
                chart.plot_y + chart.plot_height + 5.5
                - axis_metrics.fAscent
            )
            draw_text(
                label, chart.plot_x - 10.0 - label_width, baseline,
                axis_font, self._paint('FFFF0000', fill=True),
            )

        if chart.plot_border_color:
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    chart.plot_x, chart.plot_y,
                    chart.plot_width, chart.plot_height,
                ),
                self._paint(
                    'FF' + chart.plot_border_color,
                    stroke_width=0.75,
                ),
            )

        if chart.value_axis_display_unit_label:
            label = chart.value_axis_display_unit_label
            metrics = axis_title_font.getMetrics()
            if (
                chart.value_axis_display_unit_layout_x is not None
                and chart.value_axis_display_unit_layout_y is not None
            ):
                label_x = (
                    chart.x
                    + chart.width * chart.value_axis_display_unit_layout_x
                    - metrics.fTop
                )
                label_y = (
                    chart.y
                    + chart.height * chart.value_axis_display_unit_layout_y
                    + axis_title_font.measureText(label)
                    + metrics.fDescent
                )
            else:
                label_x = chart.x + 22.0
                label_y = chart.plot_y + 42.3
            canvas.save()
            canvas.translate(label_x, label_y)
            canvas.rotate(-90.0)
            canvas.drawString(
                label, 0.0, 0.0,
                axis_title_font, value_text_paint,
            )
            canvas.restore()

        axis_paint = self._paint(
            'FF' + (
                chart.category_major_gridline_color
                if chart.scatter_vary_colors else 'D9D9D9'
            ),
            stroke_width=0.75,
        )
        axis_y = chart.plot_y + chart.plot_height
        canvas.drawLine(
            chart.plot_x, axis_y,
            chart.plot_x + chart.plot_width, axis_y,
            axis_paint,
        )
        x_span = max(1e-12, chart.scatter_x_max - chart.scatter_x_min)
        tick_count = max(1, int(round(x_span / chart.scatter_x_step)))
        category_baseline = axis_y + 5.5 - axis_metrics.fAscent
        for index in range(tick_count + 1):
            value = chart.scatter_x_min + chart.scatter_x_step * index
            ratio = (value - chart.scatter_x_min) / x_span
            x = chart.plot_x + chart.plot_width * ratio
            if chart.show_vertical_gridlines:
                canvas.drawLine(
                    x, chart.plot_y, x, axis_y,
                    self._paint(
                        'FF' + chart.category_major_gridline_color,
                        stroke_width=0.75,
                    ),
                )
            if not chart.scatter_vary_colors:
                canvas.drawLine(x, axis_y, x, axis_y + 2.8, axis_paint)
            if not chart.scatter_vary_colors:
                label = self._chart_value_label(
                    value, chart.scatter_x_number_format
                )
                label_width = axis_font.measureText(label)
                draw_text(
                    label, x - label_width / 2.0, category_baseline,
                    axis_font, category_text_paint,
                )

        for series in chart.series:
            point_count = min(len(series.x_values), len(series.values))
            if point_count <= 0:
                continue
            points = []
            for x_value, y_value in zip(
                series.x_values[:point_count], series.values[:point_count]
            ):
                x_ratio = (float(x_value) - chart.scatter_x_min) / x_span
                y_ratio = self._chart_value_ratio(chart, y_value)
                points.append((
                    chart.plot_x + chart.plot_width * x_ratio,
                    chart.plot_y + chart.plot_height * (1.0 - y_ratio),
                ))
            if chart.chart_kind != 'bubble' and series.line_visible and len(points) > 1:
                path = self._skia.Path()
                path.moveTo(*points[0])
                if str(chart.scatter_style).lower().startswith('smooth'):
                    for index in range(len(points) - 1):
                        previous = points[max(0, index - 1)]
                        start = points[index]
                        end = points[index + 1]
                        following = points[min(len(points) - 1, index + 2)]
                        control1 = (
                            start[0] + (end[0] - previous[0]) / 6.0,
                            start[1] + (end[1] - previous[1]) / 6.0,
                        )
                        control2 = (
                            end[0] - (following[0] - start[0]) / 6.0,
                            end[1] - (following[1] - start[1]) / 6.0,
                        )
                        min_x, max_x = sorted((start[0], end[0]))
                        min_y, max_y = sorted((start[1], end[1]))
                        control1 = (
                            min(max(control1[0], min_x), max_x),
                            min(max(control1[1], min_y), max_y),
                        )
                        control2 = (
                            min(max(control2[0], min_x), max_x),
                            min(max(control2[1], min_y), max_y),
                        )
                        path.cubicTo(
                            *control1, *control2, end[0], end[1]
                        )
                else:
                    for point in points[1:]:
                        path.lineTo(*point)
                line_paint = self._paint(
                    'FF' + series.color,
                    stroke_width=max(0.1, series.line_width),
                )
                if series.line_cap == 'round':
                    line_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
                canvas.drawPath(path, line_paint)
            maximum_bubble_size = max(
                (
                    float(size)
                    for item in chart.series
                    for size in item.bubble_sizes
                    if float(size) > 0.0
                ),
                default=1.0,
            )
            for point_index, (x, y) in enumerate(points):
                if chart.chart_kind == 'bubble':
                    size = (
                        float(series.bubble_sizes[point_index])
                        if point_index < len(series.bubble_sizes) else 1.0
                    )
                    if size < 0.0:
                        continue
                    diameter = (
                        min(chart.plot_width, chart.plot_height)
                        * 0.28
                        * max(0.0, chart.bubble_scale) / 100.0
                        * math.sqrt(size / maximum_bubble_size)
                    )
                    radius = diameter / 2.0
                    bubble_paint = (
                        self._bubble_3d_paint(
                            series.color, x, y, radius
                        )
                        if series.bubble_3d else
                        self._paint('FF' + series.color, fill=True)
                    )
                    canvas.drawCircle(x, y, radius, bubble_paint)
                    if series.bubble_3d:
                        canvas.drawCircle(
                            x, y, radius,
                            self._paint(
                                'B3000000', stroke_width=0.55
                            ),
                        )
                    continue
                point_color = None
                if chart.scatter_vary_colors:
                    point_color = self._interpolate_chart_color(
                        '6B812E', 'D8E5C8',
                        point_index / max(1, len(points) - 1),
                    )
                self._draw_chart_marker(
                    canvas, series, x, y,
                    fill_color_override=point_color,
                    border_color_override=point_color,
                )

        if chart.scatter_vary_colors:
            overlay_title_paint = self._paint('FF000000', fill=True)
            overlay_title_paint.setStyle(
                self._skia.Paint.kStrokeAndFill_Style
            )
            overlay_title_paint.setStrokeWidth(0.45)
            if chart.category_axis_title:
                text_width = axis_title_font.measureText(
                    chart.category_axis_title
                )
                canvas.drawString(
                    chart.category_axis_title,
                    chart.plot_x + (chart.plot_width - text_width) / 2.0,
                    axis_y - 1.7,
                    axis_title_font, overlay_title_paint,
                )
            if chart.value_axis_title:
                text_width = axis_title_font.measureText(
                    chart.value_axis_title
                )
                canvas.save()
                canvas.translate(
                    chart.plot_x + 13.1,
                    chart.plot_y + chart.plot_height / 2.0 + 0.2,
                )
                canvas.rotate(-90.0)
                canvas.drawString(
                    chart.value_axis_title, -text_width / 2.0, 0.0,
                axis_title_font, overlay_title_paint,
            )
                canvas.restore()

        if not chart.scatter_vary_colors:
            self._draw_chart_axis_titles(
                canvas, chart, axis_title_font,
                value_text_paint, category_text_paint,
            )

        scatter_title_font = (
            axis_font if chart.scatter_vary_colors else title_font
        )
        title_width = scatter_title_font.measureText(chart.title)
        title_metrics = scatter_title_font.getMetrics()
        title_baseline = (
            chart.plot_y + 8.0
            if chart.scatter_vary_colors
            else chart.y + 7.5 - title_metrics.fAscent
        )
        if (
            not chart.show_legend
            and '%' in chart.axis_number_format
        ):
            title_baseline += 3.0
        if chart.scatter_vary_colors:
            canvas.save()
            canvas.translate(
                chart.x + chart.width / 2.0, title_baseline - 1.2
            )
            canvas.scale(1.133, 0.833)
            canvas.drawString(
                chart.title, -title_width / 2.0, 0.0,
                scatter_title_font, overlay_title_paint,
            )
            canvas.restore()
        else:
            title_x = chart.x + (chart.width - title_width) / 2.0
            if chart.title_border_color:
                title_box_y = (
                    title_baseline + title_metrics.fAscent - 1.51
                )
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        title_x - 3.0, title_box_y,
                        title_width + 6.0, 19.9,
                    ),
                    self._paint(
                        'FF' + chart.title_border_color,
                        stroke_width=0.75,
                    ),
                )
            draw_text(
                chart.title, title_x,
                title_baseline,
                scatter_title_font, title_paint,
            )

        if chart.show_legend and chart.series and chart.scatter_vary_colors:
            series = chart.series[0]
            count = min(len(series.x_values), len(series.values))
            point_legend_paint = self._paint('FF000000', fill=True)
            center_x = chart.plot_x + chart.plot_width + 17.75
            item_step = 18.1
            first_y = (
                chart.plot_y + chart.plot_height / 2.0
                - item_step * (count - 1) / 2.0
                + 1.25
            )
            for index, value in enumerate(series.x_values[:count]):
                center_y = first_y + item_step * index
                color = self._interpolate_chart_color(
                    '6B812E', 'D8E5C8',
                    index / max(1, count - 1),
                )
                self._draw_chart_marker(
                    canvas, series, center_x, center_y, 7.0,
                    fill_color_override=color,
                    border_color_override=color,
                )
                label = self._chart_value_label(
                    value, chart.scatter_x_number_format
                )
                baseline = center_y - (
                    axis_metrics.fAscent + axis_metrics.fDescent
                ) / 2.0
                canvas.save()
                canvas.translate(center_x + 7.0, baseline - 1.0)
                canvas.scale(1.113, 0.9)
                canvas.drawString(
                    label, 0.0, 0.0,
                    axis_font, point_legend_paint,
                )
                canvas.restore()
        elif chart.show_legend and chart.series:
            marker_size = 5.0
            text_gap = 4.0
            item_gap = 8.08 if chart.legend_layout_x is not None else 6.4
            widths = tuple(
                (22.0 if series.line_visible and chart.chart_kind != 'bubble'
                 else marker_size + text_gap)
                + axis_font.measureText(series.name)
                for series in chart.series
            )
            total_width = sum(widths) + item_gap * (len(widths) - 1)
            if chart.legend_layout_x is None:
                cursor = (
                    chart.x + (chart.width - total_width) / 2.0 + 0.67
                )
            else:
                cursor = (
                    chart.x + chart.width * chart.legend_layout_x + 10.5
                )
            if chart.legend_layout_y is None:
                center_y = chart.y + chart.height - 15.0
            else:
                center_y = (
                    chart.y + chart.height * chart.legend_layout_y + 7.1
                )
            baseline = center_y - (
                axis_metrics.fAscent + axis_metrics.fDescent
            ) / 2.0
            for series, item_width in zip(chart.series, widths):
                if series.line_visible and chart.chart_kind != 'bubble':
                    legend_paint = self._paint(
                        'FF' + series.color,
                        stroke_width=max(0.1, series.line_width),
                    )
                    if series.line_cap == 'round':
                        legend_paint.setStrokeCap(
                            self._skia.Paint.kRound_Cap
                        )
                    canvas.drawLine(
                        cursor, center_y, cursor + 19.2, center_y,
                        legend_paint,
                    )
                    self._draw_chart_marker(
                        canvas, series, cursor + 9.6, center_y,
                        marker_size,
                    )
                    text_x = cursor + 22.0
                else:
                    if chart.chart_kind == 'bubble':
                        legend_x = cursor + marker_size / 2.0
                        legend_radius = marker_size / 2.0
                        canvas.drawCircle(
                            legend_x, center_y, legend_radius,
                            (
                                self._bubble_3d_paint(
                                    series.color, legend_x, center_y,
                                    legend_radius,
                                )
                                if series.bubble_3d else
                                self._paint(
                                    'FF' + series.color, fill=True
                                )
                            ),
                        )
                    else:
                        self._draw_chart_marker(
                            canvas, series,
                            cursor + marker_size / 2.0, center_y,
                            marker_size,
                        )
                    text_x = cursor + marker_size + text_gap
                draw_text(
                    series.name, text_x,
                    baseline, axis_font, legend_text_paint,
                )
                cursor += item_width + item_gap

    def _draw_2d_waterfall_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        if not chart.series or not chart.series[0].values:
            return

        axis_metrics = axis_font.getMetrics()
        for value, ratio in self._chart_value_axis_ticks(chart):
            y = chart.plot_y + chart.plot_height * (1.0 - ratio)
            canvas.drawLine(
                chart.plot_x, y,
                chart.plot_x + chart.plot_width, y,
                grid_paint,
            )
            label = self._chart_value_label(value, chart.axis_number_format)
            label_width = axis_font.measureText(label)
            baseline = (
                y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                - 0.6
            )
            canvas.drawString(
                label, chart.plot_x - 6.0 - label_width, baseline,
                axis_font, value_text_paint,
            )

        axis_paint = self._paint('FFD9D9D9', stroke_width=0.75)
        canvas.drawLine(
            chart.plot_x, chart.plot_y,
            chart.plot_x, chart.plot_y + chart.plot_height,
            axis_paint,
        )
        canvas.drawLine(
            chart.plot_x, chart.plot_y + chart.plot_height,
            chart.plot_x + chart.plot_width,
            chart.plot_y + chart.plot_height,
            axis_paint,
        )

        values = tuple(float(value) for value in chart.series[0].values)
        count = len(values)
        slot_width = chart.plot_width / max(1, count)
        bar_width = slot_width / (1.0 + chart.waterfall_gap_width)
        subtotals = set(chart.waterfall_subtotals)
        colors = {
            'increase': '156082',
            'decrease': 'E97132',
            'total': '196B24',
        }
        connector_paint = self._paint('FFBFBFBF', stroke_width=0.75)
        label_metrics = axis_font.getMetrics()
        running = 0.0
        previous_right = None

        for index, value in enumerate(values):
            center_x = chart.plot_x + slot_width * (index + 0.5)
            left = center_x - bar_width / 2.0
            if index in subtotals:
                start_value = 0.0
                end_value = value
                kind = 'total'
            else:
                start_value = running
                end_value = running + value
                kind = 'increase' if value >= 0.0 else 'decrease'

            connector_y = chart.plot_y + chart.plot_height * (
                1.0 - self._chart_value_ratio(chart, running)
            )
            if previous_right is not None:
                canvas.drawLine(
                    previous_right, connector_y, left, connector_y,
                    connector_paint,
                )

            start_y = chart.plot_y + chart.plot_height * (
                1.0 - self._chart_value_ratio(chart, start_value)
            )
            end_y = chart.plot_y + chart.plot_height * (
                1.0 - self._chart_value_ratio(chart, end_value)
            )
            top = min(start_y, end_y)
            bottom = max(start_y, end_y)
            canvas.drawRect(
                self._skia.Rect.MakeXYWH(
                    left, top, bar_width, max(0.5, bottom - top)
                ),
                self._paint('FF' + colors[kind], fill=True),
            )

            if chart.show_data_labels:
                label = (
                    f'({abs(int(round(value))):,})'
                    if value < 0.0 else f'{int(round(value)):,}'
                )
                label_width = axis_font.measureText(label)
                if value < 0.0 and index not in subtotals:
                    label_baseline = bottom + 2.0 - label_metrics.fAscent
                else:
                    label_baseline = top - 2.0
                canvas.drawString(
                    label, center_x - label_width / 2.0, label_baseline,
                    axis_font, value_text_paint,
                )

            category = (
                str(chart.categories[index])
                if index < len(chart.categories) else str(index + 1)
            )
            category_width = axis_font.measureText(category)
            category_baseline = (
                chart.plot_y + chart.plot_height + 5.5
                - axis_metrics.fAscent
            )
            canvas.drawString(
                category, center_x - category_width / 2.0,
                category_baseline, axis_font, category_text_paint,
            )

            running = end_value
            previous_right = left + bar_width

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font, title_paint,
        )

        if chart.show_legend:
            legend_items = (
                ('Increase', colors['increase']),
                ('Decrease', colors['decrease']),
                ('Total', colors['total']),
            )
            square = 7.0
            text_gap = 4.0
            item_gap = 7.55
            widths = tuple(
                square + text_gap + axis_font.measureText(label)
                for label, _ in legend_items
            )
            total_width = sum(widths) + item_gap * (len(widths) - 1)
            cursor = chart.x + (chart.width - total_width) / 2.0 + 2.2
            square_y = chart.y + (
                36.04 if chart.show_data_labels else 38.37
            )
            baseline = square_y + (
                square - axis_metrics.fAscent - axis_metrics.fDescent
            ) / 2.0
            for (label, color), item_width in zip(legend_items, widths):
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(
                        cursor, square_y, square, square
                    ),
                    self._paint('FF' + color, fill=True),
                )
                canvas.drawString(
                    label, cursor + square + text_gap, baseline,
                    axis_font, legend_text_paint,
                )
                cursor += item_width + item_gap

    def _draw_2d_map_chart(
        self, canvas, chart, axis_font, title_font,
        legend_text_paint, title_paint,
    ):
        visible = tuple(series for series in chart.series if not series.hidden)
        series = visible[-1] if visible else (chart.series[-1] if chart.series else None)
        values = tuple(float(value) for value in series.values) if series else ()
        data = dict(zip(chart.categories, values))
        aliases = {'United States': 'United States of America'}
        data = {aliases.get(name, name): value for name, value in data.items()}
        minimum = min(values, default=0.0)
        maximum = max(values, default=0.0)
        diverging = len(chart.map_value_colors) == 3
        midpoint = (minimum + maximum) / 2.0

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font,
            title_paint,
        )

        map_x = chart.x + 17.753
        map_y = chart.y + 73.82
        map_width = 223.297
        map_height = 113.19
        map_rect = self._skia.Rect.MakeXYWH(
            map_x, map_y, map_width, map_height
        )
        border_paint = self._paint('FFFFFFFF', stroke_width=0.25)
        canvas.save()
        canvas.clipRect(map_rect, doAntiAlias=True)
        for feature in self._map_features():
            geometry = feature.get('geometry') or {}
            coordinates = geometry.get('coordinates') or ()
            polygons = (
                (coordinates,) if geometry.get('type') == 'Polygon'
                else coordinates if geometry.get('type') == 'MultiPolygon'
                else ()
            )
            country = feature.get('properties', {}).get('ADMIN', '')
            fill = (
                self._map_diverging_fill_color(
                    data[country], minimum, midpoint, maximum,
                    chart.map_value_colors,
                )
                if country in data and diverging else
                self._map_fill_color(data[country], minimum, maximum)
                if country in data else 'E0E0E0'
            )
            for polygon in polygons:
                path = self._skia.Path()
                path.setFillType(self._skia.PathFillType.kEvenOdd)
                for ring in polygon:
                    if not ring:
                        continue
                    first = ring[0]
                    path.moveTo(
                        map_x + (float(first[0]) + 180.0) / 360.0 * map_width,
                        map_y + (90.0 - float(first[1])) / 180.0 * map_height,
                    )
                    for longitude, latitude in ring[1:]:
                        path.lineTo(
                            map_x + (float(longitude) + 180.0) / 360.0 * map_width,
                            map_y + (90.0 - float(latitude)) / 180.0 * map_height,
                        )
                    path.close()
                canvas.drawPath(path, self._paint('FF' + fill, fill=True))
                canvas.drawPath(path, border_paint)
        canvas.restore()

        legend_x = chart.x + 254.64
        legend_y = chart.y + 51.84
        legend_width = 7.2
        legend_height = 64.8
        gradient = self._skia.Paint(AntiAlias=True)
        gradient_colors = (
            tuple(
                self._skia.ColorSetARGB(
                    255,
                    int(color[0:2], 16),
                    int(color[2:4], 16),
                    int(color[4:6], 16),
                )
                for color in reversed(chart.map_value_colors)
            )
            if diverging else (
                self._skia.ColorSetARGB(255, 47, 85, 151),
                self._skia.ColorSetARGB(255, 217, 226, 243),
            )
        )
        gradient_positions = (0.0, 0.5, 1.0) if diverging else (0.0, 1.0)
        gradient.setShader(self._skia.GradientShader.MakeLinear(
            (
                self._skia.Point(legend_x, legend_y),
                self._skia.Point(legend_x, legend_y + legend_height),
            ),
            gradient_colors,
            gradient_positions,
            self._skia.TileMode.kClamp,
        ))
        canvas.drawRect(
            self._skia.Rect.MakeXYWH(
                legend_x, legend_y, legend_width, legend_height
            ),
            gradient,
        )
        canvas.drawRect(
            self._skia.Rect.MakeXYWH(
                legend_x, legend_y, legend_width, legend_height
            ),
            self._paint('FFB7B7B7', stroke_width=0.25),
        )

        legend_title = series.name if series else ''
        canvas.drawString(
            legend_title,
            chart.x + 251.48,
            chart.y + 42.8,
            axis_font,
            legend_text_paint,
        )
        maximum_label = f'{maximum:g}' if diverging else f'{maximum:.3%}'
        minimum_label = f'{minimum:g}' if diverging else f'{minimum:.3%}'
        canvas.drawString(
            maximum_label,
            legend_x + 11.0,
            legend_y + 4.2,
            axis_font,
            legend_text_paint,
        )
        if diverging:
            canvas.drawString(
                f'{midpoint:g}',
                legend_x + 11.0,
                legend_y + legend_height / 2.0 + 2.1,
                axis_font,
                legend_text_paint,
            )
        canvas.drawString(
            minimum_label,
            legend_x + 11.0,
            legend_y + legend_height,
            axis_font,
            legend_text_paint,
        )

        attribution_font = self._font(chart.axis_font_resolution, 5.5)
        attribution_paint = self._paint('FF7F7F7F', fill=True)
        powered = 'Powered by Bing'
        canvas.drawString(
            powered,
            chart.x + chart.width - 9.0 - attribution_font.measureText(powered),
            chart.y + 225.5,
            attribution_font,
            attribution_paint,
        )
        copyright_text = (
            '\u00a9 DSAT for MSFT, GeoNames, Microsoft, Navteq, '
            'Thinkware Extract, Wikipedia'
        )
        canvas.drawString(
            copyright_text,
            chart.x + 48.04,
            chart.y + 233.8,
            attribution_font,
            attribution_paint,
        )

    def _draw_2d_box_whisker_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint, grid_paint,
    ):
        axis_metrics = axis_font.getMetrics()
        ticks = self._chart_value_axis_ticks(chart)
        for value, ratio in ticks:
            y = chart.plot_y + chart.plot_height * (1.0 - ratio)
            canvas.drawLine(
                chart.plot_x, y,
                chart.plot_x + chart.plot_width, y,
                grid_paint,
            )
            label = self._chart_value_label(value, chart.axis_number_format)
            label_width = axis_font.measureText(label)
            baseline = (
                y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                - 0.6
            )
            canvas.drawString(
                label, chart.plot_x - 6.0 - label_width, baseline,
                axis_font, value_text_paint,
            )

        axis_paint = self._paint(
            'FF' + (chart.value_axis_color or 'D9D9D9'),
            stroke_width=0.75,
        )
        plot_bottom = chart.plot_y + chart.plot_height
        canvas.drawLine(
            chart.plot_x, chart.plot_y, chart.plot_x, plot_bottom,
            axis_paint,
        )
        category_y = plot_bottom + 5.5 - axis_metrics.fAscent
        for label, position in chart.axis_categories:
            x = chart.plot_x + chart.plot_width * position
            canvas.drawLine(x, plot_bottom, x, plot_bottom + 2.8, grid_paint)
            label_width = axis_font.measureText(label)
            canvas.drawString(
                label, x - label_width / 2.0, category_y,
                axis_font, category_text_paint,
            )

        for series_index, series in enumerate(chart.series):
            paint = self._paint(
                'FF' + series.color,
                stroke_width=max(2.25, series.line_width),
            )
            paint.setStrokeCap(self._skia.Paint.kRound_Cap)
            for point_index, position in enumerate(chart.category_positions):
                if point_index >= len(series.values):
                    continue
                x, y = self._box_whisker_singleton_center(
                    chart, series_index, point_index
                )
                canvas.drawLine(x - 6.6, y, x + 6.6, y, paint)
                if chart.box_show_mean_marker:
                    canvas.drawLine(x - 3.0, y - 3.0, x + 3.0, y + 3.0, paint)
                    canvas.drawLine(x - 3.0, y + 3.0, x + 3.0, y - 3.0, paint)

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        canvas.drawString(
            chart.title,
            chart.x + (chart.width - title_width) / 2.0,
            chart.y + 7.5 - title_metrics.fAscent,
            title_font, title_paint,
        )

        if chart.show_legend and chart.series:
            item_widths = [
                16.0 + axis_font.measureText(series.name)
                for series in chart.series
            ]
            legend_x = chart.x + (chart.width - sum(item_widths)) / 2.0
            legend_y = chart.y + 55.8
            legend_baseline = legend_y - axis_metrics.fAscent - 1.1
            for series, item_width in zip(chart.series, item_widths):
                canvas.drawRect(
                    self._skia.Rect.MakeXYWH(legend_x, legend_y, 7.2, 7.2),
                    self._paint('FF' + series.color, fill=True),
                )
                canvas.drawString(
                    series.name, legend_x + 10.2, legend_baseline,
                    axis_font, legend_text_paint,
                )
                legend_x += item_width

    @classmethod
    def _box_whisker_singleton_center(
        cls, chart, series_index, point_index
    ):
        position = chart.category_positions[point_index]
        slot_width = chart.plot_width / max(1, len(chart.category_positions))
        marker_width = min(13.2, slot_width / max(1.0, len(chart.series) + 1.2))
        series_step = marker_width + 1.2
        series_offset = (
            series_index - (len(chart.series) - 1) / 2.0
        ) * series_step
        x = chart.plot_x + chart.plot_width * position + series_offset
        value = chart.series[series_index].values[point_index]
        y = chart.plot_y + chart.plot_height * (
            1.0 - cls._chart_value_ratio(chart, value)
        )
        return x, y

    def _draw_2d_radar_chart(
        self, canvas, chart, axis_font, title_font,
        value_text_paint, category_text_paint,
        legend_text_paint, title_paint,
    ):
        center_x, center_y, radius_x, radius_y, directions = (
            self._radar_chart_geometry(chart)
        )
        if not directions or radius_x <= 0.0 or radius_y <= 0.0:
            return

        compact_triangle = (
            chart.radar_style in ('standard', 'filled')
            and len(chart.categories) == 3
        )
        if compact_triangle:
            if chart.radar_style == 'filled':
                title_font = self._font(chart.title_font_resolution, 18.0)
            else:
                title_font = self._font(chart.title_font_resolution, 16.45)
                title_font.setScaleX(1.1975)
            legend_font = self._font(chart.axis_font_resolution, 10.0)
            legend_font.setScaleX(0.968)
        else:
            legend_font = axis_font
        if compact_triangle:
            plot_size = max(0.0, min(chart.width, chart.height) - 22.0)
            plot_rect = self._skia.Rect.MakeXYWH(
                center_x - plot_size / 2.0,
                center_y - plot_size / 2.0,
                plot_size,
                plot_size,
            )
            canvas.drawRect(
                plot_rect, self._paint('FFFFFFFF', fill=True)
            )
            canvas.save()
            canvas.clipRect(plot_rect)

        grid_paint = self._paint(
            'FF000000' if compact_triangle else (
                'FF' + chart.value_major_gridline_color
            ),
            stroke_width=1.0 if compact_triangle else 0.75,
        )
        ticks = self._chart_value_axis_ticks(chart)
        grid_path = self._skia.Path()
        grid_ticks = ticks[1:-1] if compact_triangle else ticks[1:]
        for _, ratio in grid_ticks:
            for index, (dx, dy) in enumerate(directions):
                x = center_x + dx * radius_x * ratio
                y = center_y + dy * radius_y * ratio
                if index == 0:
                    grid_path.moveTo(x, y)
                else:
                    grid_path.lineTo(x, y)
            grid_path.close()
        canvas.drawPath(grid_path, grid_paint)

        axis_metrics = axis_font.getMetrics()
        if not compact_triangle:
            for value, ratio in ticks:
                label = self._chart_value_label(
                    value, chart.axis_number_format
                )
                label_width = axis_font.measureText(label)
                y = center_y - radius_y * ratio
                baseline = (
                    y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                )
                canvas.drawString(
                    label, center_x - label_width - 10.6, baseline,
                    axis_font, value_text_paint,
                )

            for label, (dx, dy) in zip(chart.categories, directions):
                label = str(label)
                label_width = axis_font.measureText(label)
                x = center_x + dx * (radius_x + 2.0)
                y = center_y + dy * (radius_y + 9.0)
                if dx > 0.25:
                    text_x = x + 1.5
                elif dx < -0.25:
                    text_x = x - label_width - 1.5
                else:
                    text_x = x - label_width / 2.0
                baseline = (
                    y - (axis_metrics.fAscent + axis_metrics.fDescent) / 2.0
                )
                if dy < -0.75:
                    baseline -= 1.14
                elif dy < 0.0:
                    baseline += 2.18
                elif dy < 0.75:
                    baseline -= 3.22
                else:
                    baseline += 0.13
                canvas.drawString(
                    label, text_x, baseline,
                    axis_font, category_text_paint,
                )

        marker_series = []
        for series_index, series in enumerate(chart.series):
            points = []
            for value, (dx, dy) in zip(series.values, directions):
                ratio = self._chart_value_ratio(chart, float(value))
                points.append((
                    center_x + dx * radius_x * ratio,
                    center_y + dy * radius_y * ratio,
                ))
            if not points:
                continue
            path = self._skia.Path()
            path.moveTo(*points[0])
            for point in points[1:]:
                path.lineTo(*point)
            path.close()
            if chart.radar_style == 'filled':
                canvas.drawPath(
                    path,
                    self._paint(
                        ('FF' if compact_triangle else '55') + series.color,
                        fill=True,
                    ),
                )
            if not (compact_triangle and chart.radar_style == 'filled'):
                series_paint = self._paint(
                    'FF' + series.color,
                    stroke_width=self._chart_series_line_width(
                        chart, series_index
                    ),
                )
                if series.line_cap == 'round':
                    series_paint.setStrokeCap(self._skia.Paint.kRound_Cap)
                canvas.drawPath(path, series_paint)
            marker_series.append((series, points))
            if not compact_triangle and series.marker_symbol != 'none':
                for x, y in points:
                    self._draw_chart_marker(canvas, series, x, y)

        if compact_triangle:
            canvas.restore()
            for series, points in marker_series:
                for x, y in points:
                    self._draw_chart_marker(canvas, series, x, y)

        title_width = title_font.measureText(chart.title)
        title_metrics = title_font.getMetrics()
        if chart.title_layout_x is None:
            title_x = chart.x + (chart.width - title_width) / 2.0
        else:
            title_x = chart.x + chart.width * chart.title_layout_x + 3.0
        if chart.title_layout_y is None:
            title_y = chart.y + 7.5 - title_metrics.fAscent
            if compact_triangle:
                title_y += 0.64
        else:
            title_y = (
                chart.y + chart.height * chart.title_layout_y
                + 1.54 - title_metrics.fAscent
            )
        canvas.drawString(
            chart.title, title_x, title_y, title_font, title_paint
        )

        if not chart.show_legend or not chart.series:
            return
        if (
            chart.radar_style == 'marker'
            and len(chart.categories) == 3
            and chart.x < 0.0
        ):
            # Excel keeps only the chart-area border on this continuation
            # page instead of repeating the bottom legend into the clipped
            # right-hand fragment.
            return
        if chart.legend_position == 'r':
            legend_metrics = legend_font.getMetrics()
            legend_x = chart.x + chart.width - 62.36
            legend_y = center_y - (len(chart.series) - 1) * 9.04
            for index, series in enumerate(chart.series):
                line_y = legend_y + index * 18.08
                legend_paint = self._paint(
                    'FF' + series.color, stroke_width=series.line_width
                )
                if chart.radar_style == 'filled':
                    canvas.drawRect(
                        self._skia.Rect.MakeXYWH(
                            legend_x + 13.37, line_y - 2.746,
                            5.492, 5.492,
                        ),
                        self._paint('FF' + series.color, fill=True),
                    )
                else:
                    canvas.drawLine(
                        legend_x, line_y,
                        legend_x + 19.2, line_y,
                        legend_paint,
                    )
                if (
                    chart.radar_style != 'filled'
                    and series.marker_symbol != 'none'
                ):
                    self._draw_chart_marker(
                        canvas, series,
                        legend_x + 9.6, line_y,
                        size=6.0,
                    )
                baseline = (
                    line_y
                    - (
                        legend_metrics.fAscent + legend_metrics.fDescent
                    ) / 2.0
                    - 0.55
                )
                canvas.drawString(
                    series.name, legend_x + 21.30, baseline,
                    legend_font, legend_text_paint,
                )
            return
        item_gap = 9.0
        item_widths = [
            24.0 + axis_font.measureText(series.name) + item_gap
            for series in chart.series
        ]
        legend_x = (
            chart.x + (chart.width - sum(item_widths)) / 2.0 + 7.0
        )
        legend_y = chart.y + chart.height - 20.6
        legend_baseline = legend_y - axis_metrics.fAscent
        for series, item_width in zip(chart.series, item_widths):
            legend_paint = self._paint(
                'FF' + series.color, stroke_width=series.line_width
            )
            canvas.drawLine(
                legend_x, legend_y + 6.16,
                legend_x + 19.2, legend_y + 6.16,
                legend_paint,
            )
            if chart.radar_style == 'marker':
                self._draw_chart_marker(
                    canvas, series,
                    legend_x + 9.6, legend_y + 6.16,
                )
            canvas.drawString(
                series.name, legend_x + 22.0, legend_baseline,
                axis_font, legend_text_paint,
            )
            legend_x += item_width

    @staticmethod
    def _radar_chart_geometry(chart):
        category_count = len(chart.categories)
        if category_count < 3:
            return chart.x, chart.y, 0.0, 0.0, ()
        if category_count == 3:
            center_x = chart.x + chart.width / 2.0
            if chart.radar_style in ('standard', 'filled'):
                center_y = chart.y + chart.height / 2.0
                radius_x = chart.height * 0.50879
            else:
                center_y = chart.y + chart.height * 0.5301
                radius_x = max(0.0, min(
                    chart.width * 0.20145,
                    chart.height * 0.35363,
                ))
        else:
            center_x = chart.x + chart.width / 2.0 + 1.14
            center_y = chart.y + chart.height * 0.5301
            radius_x = max(0.0, min(
                chart.width * 0.20145,
                chart.height * 0.35363,
            ))
        radius_y = radius_x
        directions = tuple(
            (
                math.cos(-math.pi / 2.0 + 2.0 * math.pi * index / category_count),
                math.sin(-math.pi / 2.0 + 2.0 * math.pi * index / category_count),
            )
            for index in range(category_count)
        )
        return center_x, center_y, radius_x, radius_y, directions
