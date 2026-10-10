"""Rendering methods split from pdf_renderer.py by responsibility."""

import json
import math
import os


class PdfRendererDrawingMixin:
    def _chart_text_paint(
        self, color, image_bytes, gradient_stops=(),
        gradient_angle=0.0, bounds=None,
    ):
        paint = self._paint('FF' + (color or '595959'), fill=True)
        if gradient_stops and bounds is not None:
            x, y, width, height = bounds
            angle = math.radians(gradient_angle)
            direction_x = math.cos(angle)
            direction_y = math.sin(angle)
            center_x = x + width / 2.0
            center_y = y + height / 2.0
            half_length = (
                abs(width * direction_x) + abs(height * direction_y)
            ) / 2.0
            colors = []
            positions = []
            for position, value in sorted(gradient_stops):
                alpha, red, green, blue = self._argb_channels(value)
                colors.append(self._skia.ColorSetARGB(
                    alpha, red, green, blue
                ))
                positions.append(max(0.0, min(1.0, position)))
            paint.setShader(self._skia.GradientShader.MakeLinear(
                (
                    self._skia.Point(
                        center_x - direction_x * half_length,
                        center_y - direction_y * half_length,
                    ),
                    self._skia.Point(
                        center_x + direction_x * half_length,
                        center_y + direction_y * half_length,
                    ),
                ),
                colors,
                positions,
                self._skia.TileMode.kClamp,
            ))
        if not image_bytes:
            return paint
        return self._chart_bitmap_paint(image_bytes, 0.96, paint)

    def _chart_bitmap_paint(
        self, image_bytes, scale, paint=None, color_matrix=None,
    ):
        if paint is None:
            paint = self._paint('FFFFFFFF', fill=True)
        source = self._picture_images.get(image_bytes)
        if source is None:
            data = self._skia.Data.MakeWithoutCopy(image_bytes)
            image = self._skia.Image.MakeFromEncoded(data)
            if image is None:
                return paint
            source = (image, data)
            self._picture_images[image_bytes] = source

        shader_image = source[0]
        if color_matrix is not None:
            matrix = tuple(color_matrix)
            adjusted_key = (image_bytes, matrix)
            adjusted = self._picture_images.get(adjusted_key)
            if adjusted is None:
                surface = self._skia.Surface(
                    shader_image.width(), shader_image.height()
                )
                canvas = surface.getCanvas()
                canvas.clear(self._skia.ColorTRANSPARENT)
                filter_paint = self._skia.Paint(AntiAlias=True)
                filter_paint.setColorFilter(
                    self._skia.ColorFilters.Matrix(matrix)
                )
                canvas.drawImage(
                    shader_image, 0.0, 0.0,
                    self._skia.SamplingOptions(
                        self._skia.FilterMode.kLinear
                    ),
                    filter_paint,
                )
                snapshot = surface.makeImageSnapshot()
                encoded = snapshot.encodeToData(
                    self._skia.EncodedImageFormat.kJPEG, 90
                )
                adjusted_image = self._skia.Image.MakeFromEncoded(encoded)
                adjusted = (adjusted_image, encoded)
                self._picture_images[adjusted_key] = adjusted
            shader_image = adjusted[0]

        paint.setShader(shader_image.makeShader(
            self._skia.TileMode.kRepeat,
            self._skia.TileMode.kRepeat,
            self._skia.SamplingOptions(self._skia.FilterMode.kLinear),
            self._skia.Matrix.Scale(scale, scale),
        ))
        return paint

    def _draw_chart_marker(
        self, canvas, series, x, y, size=None,
        fill_color_override=None, border_color_override=None,
    ):
        marker_size = series.marker_size if size is None else float(size)
        if series.marker_symbol == 'none' or marker_size <= 0:
            return
        if series.marker_symbol == 'square':
            radius = marker_size * 0.5
        elif series.marker_symbol in ('diamond', 'triangle', 'x'):
            radius = marker_size * 0.4933
        else:
            radius = marker_size * 0.504
        fill_color = (
            fill_color_override or series.marker_fill_color or series.color
        )
        border_color = (
            border_color_override or series.marker_border_color or fill_color
        )
        marker_path = None
        marker_rect = None
        if series.marker_symbol == 'circle':
            canvas.drawCircle(
                x, y, radius,
                self._paint('FF' + fill_color, fill=True),
            )
        elif series.marker_symbol == 'square':
            marker_rect = self._skia.Rect.MakeXYWH(
                x - radius, y - radius, 2.0 * radius, 2.0 * radius
            )
            canvas.drawRect(
                marker_rect, self._paint('FF' + fill_color, fill=True)
            )
        elif series.marker_symbol in ('diamond', 'triangle'):
            marker_path = self._skia.Path()
            marker_path.moveTo(x, y - radius)
            if series.marker_symbol == 'diamond':
                marker_path.lineTo(x + radius, y)
                marker_path.lineTo(x, y + radius)
                marker_path.lineTo(x - radius, y)
            else:
                marker_path.lineTo(x + radius, y + radius)
                marker_path.lineTo(x - radius, y + radius)
            marker_path.close()
            canvas.drawPath(
                marker_path, self._paint('FF' + fill_color, fill=True)
            )
        elif series.marker_symbol == 'x':
            marker_paint = self._paint(
                'FF' + border_color,
                stroke_width=max(1.0, series.marker_border_width),
            )
            canvas.drawLine(
                x - radius, y - radius, x + radius, y + radius, marker_paint
            )
            canvas.drawLine(
                x - radius, y + radius, x + radius, y - radius, marker_paint
            )
        else:
            return
        if series.marker_border_width > 0:
            border_paint = self._paint(
                'FF' + border_color,
                stroke_width=series.marker_border_width,
            )
            if series.marker_symbol == 'circle':
                canvas.drawCircle(x, y, radius, border_paint)
            elif marker_rect is not None:
                canvas.drawRect(marker_rect, border_paint)
            elif marker_path is not None:
                canvas.drawPath(marker_path, border_paint)

    @staticmethod
    def _chart_series_line_width(chart, series_index):
        series = chart.series[series_index]
        if (
            chart.grouping == 'percentStacked'
            and series_index == len(chart.series) - 1
            and chart.chart_border_width > 1.0
        ):
            return max(0.1, series.line_width - 0.25)
        return series.line_width

    @staticmethod
    def _chart_value_label(value, number_format):
        if '%' in number_format:
            decimals = 0
            decimal_part = number_format.split('%', 1)[0]
            if '.' in decimal_part:
                decimals = sum(
                    character in ('0', '#')
                    for character in decimal_part.rsplit('.', 1)[1]
                )
            return f'{value * 100:.{decimals}f}%'
        rounded = int(round(value))
        if '$' in number_format:
            return f'${rounded:,}'
        if number_format.strip() == '0':
            return str(rounded)
        lowered = number_format.lower()
        if 'mmm' in lowered and ('yy' in lowered or 'yyyy' in lowered):
            if rounded == 0:
                return 'Jan-00' if 'yyyy' not in lowered else 'Jan-1900'
            from datetime import datetime, timedelta

            rendered = datetime(1899, 12, 30) + timedelta(days=rounded)
            return rendered.strftime(
                '%b-%Y' if 'yyyy' in lowered else '%b-%y'
            )
        return f'{rounded:,}'

    def _draw_shape(self, canvas, shape):
        if shape.is_connector:
            self._draw_connector(canvas, shape)
            return

        rect = self._skia.Rect.MakeXYWH(
            shape.x, shape.y, shape.width, shape.height
        )
        fill = self._paint('FF' + shape.fill_color, fill=True)
        stroke = self._paint(
            'FF' + shape.line_color, stroke_width=shape.line_width
        )
        if shape.preset_geometry == 'downArrow':
            path = self._down_arrow_path(shape)
            if shape.fill_visible:
                canvas.drawPath(path, fill)
            if shape.line_visible:
                canvas.drawPath(path, stroke)
        elif shape.preset_geometry == 'rightArrow':
            path = self._right_arrow_path(shape)
            if shape.fill_visible:
                canvas.drawPath(path, fill)
            if shape.line_visible:
                canvas.drawPath(path, stroke)
        elif shape.preset_geometry == 'mathPlus':
            path = self._math_plus_path(shape)
            if shape.fill_visible:
                canvas.drawPath(path, fill)
            if shape.line_visible:
                canvas.drawPath(path, stroke)
        else:
            if shape.fill_visible:
                canvas.drawRect(rect, fill)
            if shape.line_visible:
                canvas.drawRect(rect, stroke)

        if shape.text and shape.font_resolution is not None:
            self._draw_shape_text(canvas, shape)

    def _draw_connector(self, canvas, shape):
        use_connections = shape.preset_geometry in (
            'bentConnector2', 'curvedConnector2'
        )
        if (
            use_connections
            and shape.connector_start is not None
            and shape.connector_end is not None
        ):
            x1, y1 = shape.connector_start
            x2, y2 = shape.connector_end
        else:
            x1, x2 = shape.x, shape.x + shape.width
            if shape.flip_vertical:
                y1, y2 = shape.y + shape.height, shape.y
            else:
                y1, y2 = shape.y, shape.y + shape.height
        paint = self._paint(
            'FF' + shape.line_color, stroke_width=shape.line_width
        )
        path = self._skia.Path()
        path.moveTo(x1, y1)
        arrow_from = (x1, y1)
        if shape.preset_geometry == 'bentConnector2':
            path.lineTo(x1, y2)
            path.lineTo(x2, y2)
            arrow_from = (x1, y2)
        elif shape.preset_geometry == 'curvedConnector2':
            dx = x2 - x1
            dy = y2 - y1
            control1 = (x1 + dx * 0.55, y1)
            control2 = (x2, y2 - dy * 0.55)
            path.cubicTo(*control1, *control2, x2, y2)
            arrow_from = control2
        else:
            path.lineTo(x2, y2)
        canvas.drawPath(path, paint)
        if shape.tail_end == 'triangle':
            self._draw_arrow_head(
                canvas, arrow_from[0], arrow_from[1], x2, y2,
                shape.line_color,
            )

    def _draw_arrow_head(self, canvas, x1, y1, x2, y2, color):
        import math

        angle = math.atan2(y2 - y1, x2 - x1)
        length = 6.0
        half_width = 3.0
        base_x = x2 - length * math.cos(angle)
        base_y = y2 - length * math.sin(angle)
        path = self._skia.Path()
        path.moveTo(x2, y2)
        path.lineTo(
            base_x + half_width * math.sin(angle),
            base_y - half_width * math.cos(angle),
        )
        path.lineTo(
            base_x - half_width * math.sin(angle),
            base_y + half_width * math.cos(angle),
        )
        path.close()
        canvas.drawPath(path, self._paint('FF' + color, fill=True))

    def _draw_shape_text(self, canvas, shape):
        font = self._font(shape.font_resolution, shape.font.size)
        metrics = font.getMetrics()
        paint = self._paint('FF' + shape.font.color, fill=True)
        lines = (
            self._wrap_shape_text(
                font, shape.text, max(0.0, shape.width - 14.4)
            )
            if shape.text_wrapped else tuple(shape.text.split('\n'))
        )
        line_height = max(
            metrics.fDescent - metrics.fAscent,
            shape.font.size * 1.2,
        )
        text_height = line_height * len(lines)
        if int(shape.text_vertical_alignment) == 1:
            baseline = (
                shape.y + (shape.height - text_height) / 2.0
                - metrics.fAscent
            )
        elif int(shape.text_vertical_alignment) == 2:
            baseline = (
                shape.y + shape.height - 4.0 - text_height
                - metrics.fAscent
            )
        else:
            baseline = shape.y + 4.0 - metrics.fAscent

        if shape.text_direction == 'wordArtVertRtl':
            characters = list(reversed(shape.text.replace('\n', '')))
            widths = [font.measureText(character) for character in characters]
            gap = shape.font.size * 0.85
            total_width = sum(widths) + gap * max(0, len(characters) - 1)
            x = shape.x + (shape.width - total_width) / 2.0
            for character, width in zip(characters, widths):
                canvas.drawString(character, x, baseline, font, paint)
                x += width + gap
            return

        for line in lines:
            text_width = font.measureText(line)
            if int(shape.text_horizontal_alignment) == 1:
                x = shape.x + (shape.width - text_width) / 2.0
            elif int(shape.text_horizontal_alignment) == 2:
                x = shape.x + shape.width - 7.2 - text_width
            else:
                x = shape.x + 7.2
            canvas.drawString(line, x, baseline, font, paint)
            baseline += line_height

    @staticmethod
    def _wrap_shape_text(font, text, max_width):
        if max_width <= 0.0:
            return (text,)
        lines = []
        for paragraph in text.split('\n'):
            words = paragraph.split()
            if not words:
                lines.append('')
                continue
            line = words[0]
            for word in words[1:]:
                candidate = f'{line} {word}'
                if font.measureText(candidate) <= max_width:
                    line = candidate
                else:
                    lines.append(line)
                    line = word
            lines.append(line)
        return tuple(lines) or ('',)

    def _down_arrow_path(self, shape):
        x, y, width, height = shape.x, shape.y, shape.width, shape.height
        path = self._skia.Path()
        path.moveTo(x, y + height * 0.5)
        path.lineTo(x + width * 0.25, y + height * 0.5)
        path.lineTo(x + width * 0.25, y)
        path.lineTo(x + width * 0.75, y)
        path.lineTo(x + width * 0.75, y + height * 0.5)
        path.lineTo(x + width, y + height * 0.5)
        path.lineTo(x + width * 0.5, y + height)
        path.close()
        return path

    def _right_arrow_path(self, shape):
        x, y, width, height = shape.x, shape.y, shape.width, shape.height
        head = min(width * 0.25, max(height, 6.0))
        shaft_top = y + height / 3.0
        shaft_bottom = y + height * 2.0 / 3.0
        path = self._skia.Path()
        path.moveTo(x, shaft_top)
        path.lineTo(x + width - head, shaft_top)
        path.lineTo(x + width - head, y)
        path.lineTo(x + width, y + height / 2.0)
        path.lineTo(x + width - head, y + height)
        path.lineTo(x + width - head, shaft_bottom)
        path.lineTo(x, shaft_bottom)
        path.close()
        return path

    def _math_plus_path(self, shape):
        inset_x = shape.width * 2.0 / 15.0
        inset_y = shape.height * 2.0 / 15.0
        x = shape.x + inset_x
        y = shape.y + inset_y
        width = shape.width - 2.0 * inset_x
        height = shape.height - 2.0 * inset_y
        left_arm = x + width * 0.4
        right_arm = x + width * 0.6
        top_arm = y + height / 3.0
        bottom_arm = y + height * 2.0 / 3.0
        points = (
            (x, top_arm), (left_arm, top_arm), (left_arm, y),
            (right_arm, y), (right_arm, top_arm), (x + width, top_arm),
            (x + width, bottom_arm), (right_arm, bottom_arm),
            (right_arm, y + height), (left_arm, y + height),
            (left_arm, bottom_arm), (x, bottom_arm),
        )
        path = self._skia.Path()
        path.moveTo(*points[0])
        for point in points[1:]:
            path.lineTo(*point)
        path.close()
        return path

    def _draw_fill(self, canvas, cell):
        rect = self._skia.Rect.MakeXYWH(cell.x, cell.y, cell.width, cell.height)
        fill = cell.style.fill
        pattern_type = fill.pattern_type
        if pattern_type in (None, 'none'):
            return
        if pattern_type == 'solid':
            paint = self._paint(fill.foreground_color, fill=True)
            # Axis-aligned adjacent cells must share a seamless fill edge.
            paint.setAntiAlias(False)
            canvas.drawRect(rect, paint)
            return

        if pattern_type in self._BITMAP_PATTERN_MASKS:
            self._draw_bitmap_pattern(
                canvas,
                rect,
                pattern_type,
                fill.foreground_color,
                fill.background_color,
            )

    def _draw_data_bar(self, canvas, cell):
        bar = cell.data_bar
        if bar is None:
            return
        rect = self._skia.Rect.MakeXYWH(
            bar.x, bar.y, bar.width, bar.height
        )
        paint = self._paint(bar.color, fill=True)
        paint.setAntiAlias(False)
        if bar.gradient:
            alpha, red, green, blue = self._argb_channels(bar.color)
            paint.setShader(self._skia.GradientShader.MakeLinear(
                (
                    self._skia.Point(bar.gradient_start_x, bar.y),
                    self._skia.Point(bar.gradient_end_x, bar.y),
                ),
                (
                    self._skia.ColorSetARGB(alpha, red, green, blue),
                    self._skia.ColorSetARGB(alpha, 255, 255, 255),
                ),
                (0.0, 1.0),
                self._skia.TileMode.kClamp,
            ))
        canvas.drawRect(rect, paint)

    def _draw_icon_set(self, canvas, cell):
        icon = cell.icon_set
        if icon is None or icon.icon_set_type != '3TrafficLights1':
            return

        palette = (
            ('FFD65532', 'FFB04123'),
            ('FFEAC282', 'FFBC9749'),
            ('FF68A490', 'FF43806C'),
        )
        icon_index = min(max(int(icon.icon_index), 0), len(palette) - 1)
        fill_color, border_color = palette[icon_index]
        scale = min(icon.width, icon.height) / 12.0
        # Excel's 24px source icon leaves one transparent pixel around the
        # circle; account for half of the stroke so the visible bounds match.
        inset = 0.875 * scale
        oval = self._skia.Rect.MakeXYWH(
            icon.x + inset,
            icon.y + inset,
            max(0.0, icon.width - 2.0 * inset),
            max(0.0, icon.height - 2.0 * inset),
        )
        canvas.drawOval(oval, self._paint(fill_color, fill=True))
        canvas.drawOval(
            oval,
            self._paint(border_color, stroke_width=0.75 * scale),
        )

    def _draw_bitmap_pattern(
        self, canvas, rect, pattern_type, foreground_color, background_color
    ):
        image = self._bitmap_pattern_image(
            pattern_type, foreground_color, background_color
        )
        shader = image.makeShader(
            self._skia.TileMode.kRepeat,
            self._skia.TileMode.kRepeat,
            self._skia.SamplingOptions(self._skia.FilterMode.kNearest),
            self._skia.Matrix.Scale(0.12, 0.12),
        )
        paint = self._paint('FFFFFFFF', fill=True)
        paint.setShader(shader)
        canvas.drawRect(rect, paint)

    def _bitmap_pattern_image(
        self, pattern_type, foreground_color, background_color
    ):
        key = (pattern_type, foreground_color, background_color)
        cached = self._pattern_images.get(key)
        if cached is not None:
            return cached[0]

        foreground = self._rgba_channels(foreground_color)
        background = self._rgba_channels(background_color)
        mask = self._BITMAP_PATTERN_MASKS[pattern_type]
        pixels = bytes(
            channel
            for y in range(8)
            for x in range(8)
            for channel in (foreground if (x, y) in mask else background)
        )
        info = self._skia.ImageInfo.Make(
            8,
            8,
            self._skia.ColorType.kRGBA_8888_ColorType,
            self._skia.AlphaType.kUnpremul_AlphaType,
        )
        image = self._skia.Image.MakeRasterData(info, pixels, 8 * 4)
        if image is None:
            raise RuntimeError('Skia could not create an Excel fill pattern')
        # MakeRasterData shares storage with the image, so retain both together.
        self._pattern_images[key] = (image, pixels)
        return image

    def _draw_gridline(self, canvas, cell):
        rect = self._skia.Rect.MakeXYWH(cell.x, cell.y, cell.width, cell.height)
        canvas.drawRect(
            rect,
            self._paint('FFD1D6D9', stroke_width=0.35),
        )

    def _draw_cell_borders(self, canvas, cell):
        rect = self._skia.Rect.MakeXYWH(cell.x, cell.y, cell.width, cell.height)
        self._draw_borders(canvas, rect, cell.style.borders)

    def _draw_text(self, canvas, page, cell):
        if not cell.text or (
            cell.data_bar is not None
            and not cell.data_bar.show_value
        ) or (
            cell.icon_set is not None
            and not cell.icon_set.show_value
        ):
            return
        rect = self._skia.Rect.MakeXYWH(
            cell.x, cell.y, cell.text_clip_width, cell.height
        )
        style = cell.style
        canvas.save()
        canvas.clipRect(rect)
        paint = self._paint(style.font.color, fill=True)
        font_size = style.font.size * page.scale * cell.font_scale
        font = self._font(cell.font_resolution, font_size)
        metrics = font.getMetrics()
        line_height = cell.line_height
        prepared_lines = []
        for runs in cell.line_runs or ((),):
            prepared_runs = []
            for run in runs:
                run_font = self._text_run_font(
                    run,
                    run.font_size * page.scale * cell.font_scale,
                )
                prepared_runs.append((
                    run,
                    run_font,
                    run_font.getMetrics(),
                    run.baseline_shift * page.scale * cell.font_scale,
                    self._paint(run.color, fill=True),
                    (
                        run_font.measureText(run.text)
                        + run.advance_adjustment
                        * page.scale * cell.font_scale
                    ),
                ))
            line_width = sum(
                advance
                for _, _, _, _, _, advance in prepared_runs
            )
            prepared_lines.append((prepared_runs, line_width))

        line_count = max(1, len(cell.lines))
        block_height = line_height * line_count
        rotation_value = int(style.alignment.text_rotation or 0)
        if rotation_value == 255:
            self._draw_stacked_text(
                canvas, cell, font, metrics, line_height, paint
            )
            canvas.restore()
            return

        if cell.rich_text:
            self._draw_rich_text(
                canvas, page, cell, prepared_lines
            )
            canvas.restore()
            return

        rotation = self._excel_rotation_degrees(rotation_value)
        if rotation:
            self._draw_rotated_text(
                canvas,
                page,
                cell,
                prepared_lines,
                metrics,
                line_height,
                block_height,
                rotation,
                paint,
            )
            canvas.restore()
            return

        vertical = style.alignment.vertical
        visual_line_height = metrics.fDescent - metrics.fAscent
        visual_block_height = (
            visual_line_height + line_height * (line_count - 1)
        )
        if vertical in ('center', 'distributed', 'justify'):
            baseline = (
                cell.y + (cell.height - visual_block_height) / 2.0
                - metrics.fAscent
            )
            if (
                line_count == 1
                and not getattr(cell, 'merged', False)
                and getattr(
                    style.font, 'vertical_alignment', 'baseline'
                ) == 'baseline'
            ):
                metric_gap = max(0.0, line_height - visual_line_height)
                baseline += metric_gap / 2.0
                if self._font_underline_type(style.font) in (
                    'singleAccounting', 'doubleAccounting'
                ):
                    baseline -= metric_gap
        elif vertical == 'top':
            baseline = cell.y + page.scale - metrics.fAscent
        else:
            bottom_gap = 0.25
            if (
                cell.font_resolution.resolved_family.lower() == 'calibri'
            ):
                bottom_gap = 0.4
            baseline = (
                cell.y + cell.height - bottom_gap * page.scale
                - metrics.fDescent - line_height * (line_count - 1)
            )

        for prepared_runs, line_width in prepared_lines:
            horizontal = style.alignment.horizontal
            indent_width = cell.indent_width
            if horizontal in ('center', 'centerContinuous'):
                text_x = (
                    cell.x + (cell.width - line_width) / 2.0
                    + 0.75 * page.scale
                )
            elif horizontal == 'right' or (horizontal == 'general' and cell.numeric):
                text_x = (
                    cell.x + cell.width - line_width
                    - 1.5 * page.scale - indent_width
                )
            else:
                padding = (
                    2.0
                    if getattr(
                        style.font, 'vertical_alignment', 'baseline'
                    ) in ('superscript', 'subscript')
                    else 2.4
                )
                text_x = cell.x + padding * page.scale + indent_width
            line_x = text_x
            for (
                run, run_font, _, baseline_shift, run_paint, advance
            ) in prepared_runs:
                self._draw_text_run(
                    canvas,
                    run.text,
                    text_x,
                    baseline + baseline_shift,
                    run_font,
                    style,
                    run_paint,
                    run,
                )
                text_x += advance
            self._draw_hyperlink_annotation(
                canvas, cell.hyperlink, line_x, baseline, line_width, metrics
            )
            baseline += line_height
        canvas.restore()

    def _draw_rich_text(self, canvas, page, cell, prepared_lines):
        prepared_runs, line_width = prepared_lines[0]
        if not prepared_runs:
            return
        line_top = min(
            baseline_shift + metrics.fAscent
            for _, _, metrics, baseline_shift, _, _ in prepared_runs
        )
        line_bottom = max(
            baseline_shift + metrics.fDescent
            for _, _, metrics, baseline_shift, _, _ in prepared_runs
        )
        visual_height = line_bottom - line_top
        vertical = cell.style.alignment.vertical
        if vertical in ('center', 'distributed', 'justify'):
            baseline = (
                cell.y + (cell.height - visual_height) / 2.0 - line_top
            )
        elif vertical == 'top':
            baseline = cell.y - line_top
        else:
            baseline = cell.y + cell.height - visual_height - line_top

        horizontal = cell.style.alignment.horizontal
        indent_width = cell.indent_width
        if horizontal in ('center', 'centerContinuous'):
            text_x = (
                cell.x + (cell.width - line_width) / 2.0
                + 0.75 * page.scale
            )
        elif horizontal == 'right' or (
            horizontal == 'general' and cell.numeric
        ):
            text_x = (
                cell.x + cell.width - line_width
                - 1.5 * page.scale - indent_width
            )
        else:
            text_x = cell.x + 2.4 * page.scale + indent_width
        line_x = text_x
        for (
            run, run_font, _, baseline_shift, run_paint, advance
        ) in prepared_runs:
            self._draw_text_run(
                canvas,
                run.text,
                text_x,
                baseline + baseline_shift,
                run_font,
                cell.style,
                run_paint,
                run,
            )
            text_x += advance
        first_metrics = prepared_runs[0][2]
        self._draw_hyperlink_annotation(
            canvas,
            cell.hyperlink,
            line_x,
            baseline,
            line_width,
            first_metrics,
        )

    def _draw_stacked_text(
        self, canvas, cell, font, metrics, line_height, paint
    ):
        columns = str(cell.text).split() or ('',)
        block_width = line_height * len(columns)
        block_height = line_height * max(len(column) for column in columns)
        left = cell.x + (cell.width - block_width) / 2.0
        top = cell.y + (cell.height - block_height) / 2.0
        baseline = top - metrics.fAscent

        for index, column in enumerate(columns):
            slot = len(columns) - index - 1
            slot_left = left + slot * line_height
            for char_index, char in enumerate(column):
                char_width = font.measureText(char)
                char_x = slot_left + (line_height - char_width) / 2.0
                char_baseline = baseline + char_index * line_height
                self._draw_text_run(
                    canvas,
                    char,
                    char_x,
                    char_baseline,
                    font,
                    cell.style,
                    paint,
                )

    def _draw_rotated_text(
        self, canvas, page, cell, prepared_lines, metrics, line_height,
        block_height, rotation, paint,
    ):
        import math

        block_width = max((width for _, width in prepared_lines), default=0.0)
        radians = math.radians(rotation)
        cosine = math.cos(radians)
        sine = math.sin(radians)
        corners = (
            (0.0, 0.0),
            (block_width, 0.0),
            (0.0, block_height),
            (block_width, block_height),
        )
        rotated_x = [x * cosine - y * sine for x, y in corners]
        rotated_y = [x * sine + y * cosine for x, y in corners]
        min_x, max_x = min(rotated_x), max(rotated_x)
        min_y, max_y = min(rotated_y), max(rotated_y)
        rotated_width = max_x - min_x
        rotated_height = max_y - min_y

        horizontal = cell.style.alignment.horizontal
        # Excel leaves about five screen pixels between rotated text and the
        # cell edge. Clockwise general-aligned text is anchored from the
        # right edge so its rotated bounds stay inside the source column.
        padding = 3.75 * page.scale
        indent_width = cell.indent_width
        if horizontal in ('center', 'centerContinuous'):
            target_x = cell.x + (cell.width - rotated_width) / 2.0
        elif horizontal == 'right' or (
            horizontal == 'general' and cell.numeric
        ) or (
            horizontal == 'general' and rotation > 0
        ):
            target_x = (
                cell.x + cell.width - rotated_width - padding - indent_width
            )
        else:
            target_x = cell.x + padding + indent_width

        vertical = cell.style.alignment.vertical
        if vertical in ('center', 'distributed', 'justify'):
            target_y = cell.y + (cell.height - rotated_height) / 2.0
        elif vertical == 'top':
            target_y = cell.y
        else:
            target_y = cell.y + cell.height - rotated_height

        canvas.translate(target_x - min_x, target_y - min_y)
        canvas.rotate(rotation)
        baseline = -metrics.fAscent
        for prepared_runs, line_width in prepared_lines:
            if horizontal in ('center', 'centerContinuous'):
                text_x = (block_width - line_width) / 2.0
            elif horizontal == 'right' or (
                horizontal == 'general' and cell.numeric
            ):
                text_x = block_width - line_width
            else:
                text_x = 0.0
            for (
                run, run_font, _, baseline_shift, run_paint, advance
            ) in prepared_runs:
                self._draw_text_run(
                    canvas, run.text, text_x, baseline + baseline_shift,
                    run_font, cell.style, run_paint, run,
                )
                text_x += advance
            baseline += line_height

    def _draw_hyperlink_annotation(
        self, canvas, target, x, baseline, width, metrics
    ):
        if not target or width <= 0:
            return
        rect = self._skia.Rect.MakeXYWH(
            x,
            baseline + metrics.fAscent,
            width,
            metrics.fDescent - metrics.fAscent,
        )
        data = self._skia.Data.MakeWithCopy(
            str(target).encode('utf-8') + b'\0'
        )
        canvas.drawAnnotation(rect, 'SkAnnotationKey_URL', data)

    def _draw_text_run(
        self, canvas, text, x, baseline, font, style, paint, font_style=None
    ):
        width = font.measureText(text)
        canvas.drawString(text, x, baseline, font, paint)
        metrics = font.getMetrics()
        font_style = font_style or style.font
        underline_type = self._font_underline_type(font_style)
        if underline_type != 'none':
            position = float(
                getattr(metrics, 'fUnderlinePosition', 0.0)
                or max(0.5, font.getSize() * 0.08)
            )
            thickness = float(
                getattr(metrics, 'fUnderlineThickness', 0.0)
                or max(0.5, font.getSize() / 18.0)
            )
            if underline_type in ('singleAccounting', 'doubleAccounting'):
                position += font.getSize() * 0.08
                underline_width = width + font.getSize() * 0.38
            else:
                underline_width = width
            underline_paint = self._paint(
                font_style.color, stroke_width=thickness
            )
            positions = [position]
            if underline_type in ('double', 'doubleAccounting'):
                positions.append(position + max(
                    thickness * 2.4, font.getSize() * 0.12
                ))
            for underline_position in positions:
                canvas.drawLine(
                    x,
                    baseline + underline_position,
                    x + underline_width,
                    baseline + underline_position,
                    underline_paint,
                )
        if font_style.strikethrough:
            position = float(
                getattr(metrics, 'fStrikeoutPosition', 0.0)
                or -font.getSize() * 0.3
            )
            thickness = float(
                getattr(metrics, 'fStrikeoutThickness', 0.0)
                or max(0.5, font.getSize() / 18.0)
            )
            canvas.drawLine(
                x,
                baseline + position,
                x + width,
                baseline + position,
                self._paint(font_style.color, stroke_width=thickness),
            )
        return width

    @staticmethod
    def _font_underline_type(font_style):
        if not bool(getattr(font_style, 'underline', False)):
            return 'none'
        underline_type = getattr(font_style, 'underline_type', None)
        supported = {
            'single', 'double', 'singleAccounting', 'doubleAccounting',
        }
        return underline_type if underline_type in supported else 'single'

    @staticmethod
    def _excel_rotation_degrees(rotation):
        rotation = int(rotation or 0)
        if 1 <= rotation <= 90:
            return -rotation
        if 91 <= rotation <= 180:
            return rotation - 90
        return 0

    def _draw_borders(self, canvas, rect, borders):
        edges = (
            (borders.top, rect.left(), rect.top(), rect.right(), rect.top()),
            (borders.bottom, rect.left(), rect.bottom(), rect.right(), rect.bottom()),
            (borders.left, rect.left(), rect.top(), rect.left(), rect.bottom()),
            (borders.right, rect.right(), rect.top(), rect.right(), rect.bottom()),
        )
        for border, x1, y1, x2, y2 in edges:
            if border.line_style in (None, 'none'):
                continue
            paint = self._paint(
                border.color,
                stroke_width=self._border_width(border),
                line_style=border.line_style,
            )
            canvas.drawLine(x1, y1, x2, y2, paint)

        diagonal = borders.diagonal
        if diagonal.line_style in (None, 'none'):
            return
        paint = self._paint(
            diagonal.color,
            stroke_width=self._border_width(diagonal),
            line_style=diagonal.line_style,
        )
        if borders.diagonal_up:
            canvas.drawLine(
                rect.left(), rect.bottom(), rect.right(), rect.top(), paint
            )
        if borders.diagonal_down:
            canvas.drawLine(
                rect.left(), rect.top(), rect.right(), rect.bottom(), paint
            )

    def _font(self, resolution, size):
        key = (
            resolution.font_path,
            resolution.resolved_family,
            resolution.bold,
            resolution.italic,
        )
        if key not in self._typefaces:
            style = self._font_style(resolution.bold, resolution.italic)
            typeface = None
            if resolution.font_path:
                typeface = self._skia.Typeface.MakeFromFile(resolution.font_path)
            if typeface is None:
                typeface = self._skia.Typeface.MakeFromName(
                    resolution.resolved_family, style
                )
            if typeface is None:
                typeface = self._skia.Typeface.MakeFromName('Arial', style)
            self._typefaces[key] = typeface
        return self._skia.Font(self._typefaces[key], max(1.0, float(size)))

    def _text_run_font(self, run, size):
        font = self._font(run.font_resolution, size)
        if run.synthetic_italic:
            font.setSkewX(self._SYNTHETIC_ITALIC_SKEW)
        return font

    def _font_style(self, bold, italic):
        if bold and italic:
            return self._skia.FontStyle.BoldItalic()
        if bold:
            return self._skia.FontStyle.Bold()
        if italic:
            return self._skia.FontStyle.Italic()
        return self._skia.FontStyle.Normal()

    def _paint(self, color, fill=False, stroke_width=1.0, line_style=None):
        paint = self._skia.Paint(
            AntiAlias=True,
            Style=(
                self._skia.Paint.kFill_Style
                if fill else self._skia.Paint.kStroke_Style
            ),
            StrokeWidth=float(stroke_width),
        )
        alpha, red, green, blue = self._argb_channels(color)
        paint.setColor4f(
            self._skia.Color4f(*(
                round(channel / 255.0, 3)
                for channel in (red, green, blue, alpha)
            )),
            None,
        )
        if line_style == 'slantDashDot':
            paint.setPathEffect(
                self._skia.DashPathEffect.Make(
                    [10.0, 5.5, 1.5, 5.5], 0.0
                )
            )
        elif line_style in ('dotted', 'hair'):
            paint.setPathEffect(self._skia.DashPathEffect.Make([1.0, 2.0], 0.0))
        elif 'dash' in str(line_style or '').lower():
            paint.setPathEffect(self._skia.DashPathEffect.Make([4.0, 2.0], 0.0))
        return paint

    @staticmethod
    def _argb_channels(value):
        text = str(value or 'FF000000').lstrip('#')
        if len(text) == 6:
            text = 'FF' + text
        if len(text) != 8:
            text = 'FF000000'
        try:
            alpha, red, green, blue = (
                int(text[index:index + 2], 16) for index in (0, 2, 4, 6)
            )
        except ValueError:
            alpha, red, green, blue = 255, 0, 0, 0
        return alpha, red, green, blue

    @classmethod
    def _rgba_channels(cls, value):
        alpha, red, green, blue = cls._argb_channels(value)
        return red, green, blue, alpha

    @staticmethod
    def _border_width(border):
        if (
            border.line_style == 'thin'
            and getattr(border, 'automatic_color', False)
        ):
            return 0.75
        style_widths = {
            'hair': 0.05, 'thin': 0.14, 'medium': 1.0,
            'thick': 1.5, 'double': 1.5, 'slantDashDot': 1.5,
        }
        return style_widths.get(
            border.line_style, max(0.5, float(border.weight or 1) * 0.5)
        )
