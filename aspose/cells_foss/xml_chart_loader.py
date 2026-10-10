"""
Aspose.Cells for Python - Chart XML Loader Module

This module loads chart settings from XLSX parts into chart APIs.
"""

import posixpath
import re
import xml.etree.ElementTree as ET
from .chart import ChartType, ChartAxis, ChartErrorBars
from .xml_picture_loader import PictureXmlLoader
from .xml_shape_loader import ShapeXmlLoader


# Map from XML element local-name to ChartType
_CHART_ELEM_TYPE_MAP = {
    "lineChart": ChartType.LINE,
    "line3DChart": ChartType.LINE,
    "barChart": ChartType.BAR,
    "bar3DChart": ChartType.BAR,
    "pieChart": ChartType.PIE,
    "doughnutChart": ChartType.PIE,
    "pie3DChart": ChartType.PIE,
    "ofPieChart": ChartType.PIE,
    "areaChart": ChartType.AREA,
    "area3DChart": ChartType.AREA,
    "scatterChart": ChartType.SCATTER,
    "bubbleChart": ChartType.SCATTER,
    "stockChart": ChartType.STOCK,
    "surface3DChart": ChartType.SURFACE,
    "surfaceChart": ChartType.SURFACE,
    "radarChart": ChartType.RADAR,
}

_3D_ELEMS = {"line3DChart", "bar3DChart", "pie3DChart", "area3DChart", "surface3DChart"}


class ChartXmlLoader:
    """Loads worksheet chart settings from drawing/chart XML parts."""

    def __init__(self, worksheet_ns):
        self._ws_ns = worksheet_ns
        self._rels_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        self._xdr_ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        self._c_ns = {
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        self._cx_ns = {
            'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        self._r_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        self._picture_loader = PictureXmlLoader()
        self._shape_loader = ShapeXmlLoader()

    def load_charts(self, worksheet, worksheet_root, zipf, sheet_num, content_type_overrides=None, content_type_defaults=None):
        """
        Loads charts for a worksheet from drawing/chart parts.
        """
        drawing_elem = worksheet_root.find('main:drawing', namespaces=self._ws_ns)
        if drawing_elem is None:
            return

        drawing_rel_id = drawing_elem.get(self._r_attr)
        if not drawing_rel_id:
            return

        sheet_path = f'xl/worksheets/sheet{sheet_num}.xml'
        sheet_rels_path = f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels'
        sheet_rels = self._read_relationships(zipf, sheet_rels_path)
        drawing_target = sheet_rels.get(drawing_rel_id)
        if not drawing_target:
            return

        drawing_path = self._resolve_target(sheet_path, drawing_target)
        drawing_bytes = self._try_read_bytes(zipf, drawing_path)
        drawing_root = None
        if drawing_bytes is not None:
            try:
                drawing_root = ET.fromstring(drawing_bytes)
            except ET.ParseError:
                drawing_root = None
        if drawing_root is None:
            return

        drawing_rels_path = self._rels_path_for_part(drawing_path)
        drawing_rels = self._read_relationships(zipf, drawing_rels_path)
        drawing_rel_types = self._read_relationship_types(zipf, drawing_rels_path)
        drawing_rels_bytes = self._try_read_bytes(zipf, drawing_rels_path)
        worksheet._source_drawing_xml = drawing_bytes
        worksheet._source_drawing_rels_xml = drawing_rels_bytes
        worksheet._source_drawing_part_path = drawing_path  # e.g. "xl/drawings/drawing5.xml"
        worksheet._drawing_dirty = False
        self._picture_loader.collect_drawing_image_parts(
            worksheet,
            drawing_path,
            drawing_rels,
            drawing_rel_types,
            zipf,
            content_type_overrides=content_type_overrides,
            content_type_defaults=content_type_defaults,
        )
        self._picture_loader.load_pictures(
            worksheet,
            drawing_root,
            drawing_path,
            drawing_rels,
            zipf,
            self._get_anchor_int,
            content_type_overrides=content_type_overrides,
            content_type_defaults=content_type_defaults,
            drawing_rel_types=drawing_rel_types,
        )
        self._shape_loader.load_shapes(
            worksheet,
            drawing_root,
            drawing_path,
            drawing_rels,
            zipf,
            self._get_anchor_int,
            drawing_rel_types=drawing_rel_types,
        )

        # oneCellAnchor holds a chart just as validly as twoCellAnchor; skipping it
        # drops the chart while its rel stays in the preserved drawingN.xml.rels.
        anchors = []
        anchors.extend(drawing_root.findall('xdr:twoCellAnchor', namespaces=self._xdr_ns))
        anchors.extend(drawing_root.findall('xdr:oneCellAnchor', namespaces=self._xdr_ns))
        for anchor in anchors:
            # One anchor can hold several graphicFrames -- charts grouped in an
            # xdr:grpSp all share a single anchor. Every c:chart/cx:chart below it
            # must be loaded: a chart we skip is still referenced by the preserved
            # drawingN.xml.rels, leaving a rel pointing at a part never written.
            chart_elems = [
                (el, False) for el in anchor.findall('.//c:chart', namespaces=self._xdr_ns)
            ]
            chart_elems += [
                (el, True) for el in anchor.findall('.//cx:chart', namespaces=self._xdr_ns)
            ]
            for chart_elem, is_chart_ex in chart_elems:
                chart_rel_id = chart_elem.get(self._r_attr)
                if not chart_rel_id:
                    continue

                chart_target = drawing_rels.get(chart_rel_id)
                if not chart_target:
                    continue

                chart_path = self._resolve_target(drawing_path, chart_target)
                chart_bytes = self._try_read_bytes(zipf, chart_path)
                if chart_bytes is None:
                    continue
                try:
                    chart_root = ET.fromstring(chart_bytes)
                except ET.ParseError:
                    continue
                if chart_root is None:
                    continue

                if not is_chart_ex:
                    rel_type = drawing_rel_types.get(chart_rel_id, "")
                    is_chart_ex = rel_type.endswith('/chartEx')

                chart_type = None
                plot_area = None
                is_3d = False

                if is_chart_ex:
                    plot_chart = chart_root.find('.//cx:plotAreaRegion', namespaces=self._cx_ns)
                    if plot_chart is not None:
                        chart_type = self._detect_chart_ex_type(plot_chart)
                else:
                    plot_area = chart_root.find('.//c:chart/c:plotArea', namespaces=self._c_ns)
                    if plot_area is not None:
                        chart_type, is_3d = self._detect_standard_chart_type(plot_area)

                if chart_type is None:
                    continue

                from_col = self._get_anchor_int(anchor, 'xdr:from/xdr:col')
                from_row = self._get_anchor_int(anchor, 'xdr:from/xdr:row')
                if from_col is None or from_row is None:
                    continue
                # oneCellAnchor sizes itself with xdr:ext instead of xdr:to, so fall
                # back to a one-cell span the way the picture loader does.
                to_col = self._get_anchor_int(anchor, 'xdr:to/xdr:col', default=from_col + 1)
                to_row = self._get_anchor_int(anchor, 'xdr:to/xdr:row', default=from_row + 1)

                chart_idx = worksheet.charts.add(chart_type, from_row, from_col, to_row, to_col)
                chart = worksheet.charts[chart_idx]
                chart._suspend_dirty_tracking = True
                chart._upper_left_column_offset = self._get_anchor_int(anchor, 'xdr:from/xdr:colOff', default=0)
                chart._upper_left_row_offset = self._get_anchor_int(anchor, 'xdr:from/xdr:rowOff', default=0)
                chart._lower_right_column_offset = self._get_anchor_int(anchor, 'xdr:to/xdr:colOff', default=0)
                chart._lower_right_row_offset = self._get_anchor_int(anchor, 'xdr:to/xdr:rowOff', default=0)
                anchor_extent = anchor.find(
                    'xdr:ext', namespaces=self._xdr_ns
                )
                if anchor_extent is not None:
                    try:
                        chart._anchor_extent_width = int(
                            float(anchor_extent.get('cx'))
                        )
                        chart._anchor_extent_height = int(
                            float(anchor_extent.get('cy'))
                        )
                    except (TypeError, ValueError):
                        chart._anchor_extent_width = None
                        chart._anchor_extent_height = None
                chart._is_3d = is_3d
                chart._source_chart_xml = chart_bytes
                chart._source_chart_rels_xml = None
                chart._source_chart_extra_parts = []
                chart._source_chart_part_path = chart_path
                chart._source_chart_content_type = None
                if content_type_overrides:
                    chart._source_chart_content_type = content_type_overrides.get(f'/{chart_path}')
                chart._source_is_chart_ex = is_chart_ex
                chart._source_chart_part_index = None
                chart_match = re.search(r'chart(\d+)\.xml$', chart_path)
                if chart_match:
                    chart._source_chart_part_index = int(chart_match.group(1))

                chart_rels_path = self._rels_path_for_part(chart_path)
                chart_rels_bytes = self._try_read_bytes(zipf, chart_rels_path)
                if chart_rels_bytes is not None:
                    chart._source_chart_rels_xml = chart_rels_bytes
                    try:
                        chart_rels_root = ET.fromstring(chart_rels_bytes)
                        for rel in chart_rels_root.findall('rel:Relationship', namespaces=self._rels_ns):
                            target = rel.get('Target')
                            if not target:
                                continue
                            part_path = self._resolve_target(chart_path, target)
                            part_bytes = self._try_read_bytes(zipf, part_path)
                            if part_bytes is None:
                                continue
                            content_type = None
                            if content_type_overrides:
                                content_type = content_type_overrides.get(f'/{part_path}')
                            chart._source_chart_extra_parts.append((part_path, part_bytes, content_type))
                    except ET.ParseError:
                        pass

                if is_chart_ex:
                    plot_chart = chart_root.find('.//cx:plotAreaRegion', namespaces=self._cx_ns)
                    self._load_chart_ex_settings(chart, chart_root, plot_chart, worksheet)
                else:
                    self._load_chart_settings(chart, chart_root, plot_area, is_3d)
                chart._suspend_dirty_tracking = False

    def _detect_standard_chart_type(self, plot_area):
        """
        Detects the chart type from a plotArea element.

        Returns:
            tuple(ChartType, bool): (chart_type, is_3d).
            chart_type is COMBO if multiple sub-chart types found.
        """
        c_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        found_types = []
        for child in plot_area:
            local = child.tag.replace(f'{{{c_uri}}}', '')
            if local in _CHART_ELEM_TYPE_MAP:
                found_types.append(local)

        if not found_types:
            return None, False

        is_3d = any(t in _3D_ELEMS for t in found_types)

        if len(found_types) == 1:
            return _CHART_ELEM_TYPE_MAP[found_types[0]], is_3d

        unique_types = set(found_types)
        if "stockChart" in unique_types and unique_types.issubset({"stockChart", "barChart"}):
            return ChartType.STOCK, False

        return ChartType.COMBO, False

    def _load_chart_settings(self, chart, chart_root, plot_area, is_3d=False):
        """Loads settings for standard (non-chartEx) charts, including combo."""
        c_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

        # Collect all sub-chart elements in order
        sub_chart_elems = []
        for child in plot_area:
            local = child.tag.replace(f'{{{c_uri}}}', '')
            if local in _CHART_ELEM_TYPE_MAP:
                sub_chart_elems.append((local, child))

        stock_like = any(name == "stockChart" for name, _ in sub_chart_elems)
        if stock_like:
            self._load_stock_chart_settings(chart, plot_area, sub_chart_elems)
        else:
            is_combo = len(sub_chart_elems) > 1
            if is_combo:
                self._load_combo_chart_settings(chart, chart_root, plot_area, sub_chart_elems)
            elif sub_chart_elems:
                elem_name, plot_chart_elem = sub_chart_elems[0]
                self._load_single_chart_settings(chart, chart_root, plot_chart_elem, is_3d)

        # Common: title and legend apply to both
        title_elem = chart_root.find('.//c:chart/c:title', namespaces=self._c_ns)
        if title_elem is not None:
            text_elem = title_elem.find('.//a:t', namespaces=self._c_ns)
            if text_elem is not None:
                chart.title = text_elem.text if text_elem.text is not None else ""
            else:
                # Empty title element means auto-title deleted or overlay-only
                overlay = title_elem.find('c:overlay', namespaces=self._c_ns)
                if overlay is not None:
                    chart.title = None  # represents empty/no text title
                else:
                    chart.title = ""

        auto_title_deleted = chart_root.find('.//c:chart/c:autoTitleDeleted', namespaces=self._c_ns)
        if auto_title_deleted is not None:
            chart._auto_title_deleted = auto_title_deleted.get('val', '0') in ('1', 'true', 'True')

        legend_elem = chart_root.find('.//c:chart/c:legend', namespaces=self._c_ns)
        if legend_elem is not None:
            legend_pos = legend_elem.find('c:legendPos', namespaces=self._c_ns)
            if legend_pos is not None and legend_pos.get('val'):
                try:
                    chart.legend_position = legend_pos.get('val')
                except ValueError:
                    pass
            legend_delete = legend_elem.find('c:delete', namespaces=self._c_ns)
            if legend_delete is not None:
                chart.show_legend = legend_delete.get('val', '0') not in ('1', 'true', 'True')
        else:
            chart.show_legend = False

        disp_blanks = chart_root.find('.//c:chart/c:dispBlanksAs', namespaces=self._c_ns)
        if disp_blanks is not None and disp_blanks.get('val'):
            chart._disp_blanks_as = disp_blanks.get('val')

        if is_3d:
            view3d_elem = chart_root.find('.//c:chart/c:view3D', namespaces=self._c_ns)
            if view3d_elem is not None:
                self._load_view3d(chart, view3d_elem)

    def _load_stock_chart_settings(self, chart, plot_area, sub_chart_elems):
        """Loads stock chart settings from stockChart or bar+stock structure."""
        bar_elem = None
        stock_elem = None
        for elem_name, elem in sub_chart_elems:
            if elem_name == "barChart":
                bar_elem = elem
            elif elem_name == "stockChart":
                stock_elem = elem
        if stock_elem is None:
            return

        has_volume = bar_elem is not None
        has_open = stock_elem.find('c:upDownBars', namespaces=self._c_ns) is not None
        if has_volume and has_open:
            chart._stock_style = "volume_open_high_low_close"
        elif has_volume:
            chart._stock_style = "volume_high_low_close"
        elif has_open:
            chart._stock_style = "open_high_low_close"
        else:
            chart._stock_style = "high_low_close"

        if has_volume:
            for ser in bar_elem.findall('c:ser', namespaces=self._c_ns):
                self._load_series(chart, ser, chart_type=ChartType.STOCK)

        for ser in stock_elem.findall('c:ser', namespaces=self._c_ns):
            self._load_series(chart, ser, chart_type=ChartType.STOCK)

        self._load_axes(chart, plot_area)

    def _load_single_chart_settings(self, chart, chart_root, plot_chart_elem, is_3d=False):
        """Loads settings for a simple single-type chart into the chart object."""
        chart._is_bubble = plot_chart_elem.tag.endswith('}bubbleChart')
        bubble_scale_elem = plot_chart_elem.find(
            'c:bubbleScale', namespaces=self._c_ns
        )
        if bubble_scale_elem is not None:
            try:
                chart._bubble_scale = float(bubble_scale_elem.get('val', 100))
            except (TypeError, ValueError):
                chart._bubble_scale = 100.0
        grouping_elem = plot_chart_elem.find('c:grouping', namespaces=self._c_ns)
        if grouping_elem is not None and grouping_elem.get('val'):
            try:
                chart.grouping = grouping_elem.get('val')
            except ValueError:
                pass

        bar_dir_elem = plot_chart_elem.find('c:barDir', namespaces=self._c_ns)
        if bar_dir_elem is not None and bar_dir_elem.get('val'):
            try:
                chart.bar_direction = bar_dir_elem.get('val')
            except ValueError:
                pass

        gap_width_elem = plot_chart_elem.find('c:gapWidth', namespaces=self._c_ns)
        if gap_width_elem is not None and gap_width_elem.get('val'):
            try:
                chart.gap_width = int(gap_width_elem.get('val'))
            except ValueError:
                pass

        overlap_elem = plot_chart_elem.find('c:overlap', namespaces=self._c_ns)
        if overlap_elem is not None and overlap_elem.get('val'):
            try:
                chart.overlap = int(overlap_elem.get('val'))
            except ValueError:
                pass

        vary_colors_elem = plot_chart_elem.find('c:varyColors', namespaces=self._c_ns)
        if vary_colors_elem is not None:
            chart.vary_colors = vary_colors_elem.get('val', '0') in ('1', 'true', 'True')

        first_slice_ang = plot_chart_elem.find('c:firstSliceAng', namespaces=self._c_ns)
        if first_slice_ang is not None and first_slice_ang.get('val'):
            try:
                chart.first_slice_angle = int(first_slice_ang.get('val'))
            except ValueError:
                pass

        of_pie_type_elem = plot_chart_elem.find('c:ofPieType', namespaces=self._c_ns)
        if of_pie_type_elem is not None and of_pie_type_elem.get('val'):
            chart.is_of_pie = True
            try:
                chart.of_pie_type = of_pie_type_elem.get('val')
            except ValueError:
                pass

        second_pie_size_elem = plot_chart_elem.find('c:secondPieSize', namespaces=self._c_ns)
        if second_pie_size_elem is not None and second_pie_size_elem.get('val'):
            chart.is_of_pie = True
            try:
                chart.second_pie_size = int(second_pie_size_elem.get('val'))
            except ValueError:
                pass

        smooth_elem = plot_chart_elem.find('c:smooth', namespaces=self._c_ns)
        if smooth_elem is not None:
            chart.smooth = smooth_elem.get('val', '0') in ('1', 'true', 'True')

        scatter_style_elem = plot_chart_elem.find('c:scatterStyle', namespaces=self._c_ns)
        if scatter_style_elem is not None and scatter_style_elem.get('val'):
            chart._scatter_style = scatter_style_elem.get('val')

        wireframe_elem = plot_chart_elem.find('c:wireframe', namespaces=self._c_ns)
        if wireframe_elem is not None:
            chart._wireframe = wireframe_elem.get('val', '0') in ('1', 'true', 'True')

        radar_style_elem = plot_chart_elem.find('c:radarStyle', namespaces=self._c_ns)
        if radar_style_elem is not None and radar_style_elem.get('val'):
            chart._radar_style = radar_style_elem.get('val')

        if is_3d:
            gap_depth_elem = plot_chart_elem.find('c:gapDepth', namespaces=self._c_ns)
            if gap_depth_elem is not None and gap_depth_elem.get('val'):
                try:
                    chart.gap_depth = int(gap_depth_elem.get('val'))
                except ValueError:
                    pass

        # Load series 鈥?pass actual chart type so scatter charts use xVal/yVal
        for ser in plot_chart_elem.findall('c:ser', namespaces=self._c_ns):
            self._load_series(chart, ser, chart_type=chart.type)

    def _load_combo_chart_settings(self, chart, chart_root, plot_area, sub_chart_elems):
        """
        Loads a combo chart where plotArea has multiple chart type elements.
        Populates chart.sub_charts and chart.n_series with per-series chart_type.
        """
        chart._sub_charts = []

        for elem_name, sub_elem in sub_chart_elems:
            sub_type = _CHART_ELEM_TYPE_MAP[elem_name]

            sub_chart_info = {
                "type": sub_type,
                "series": [],
                "bar_direction": "col",
                "grouping": "clustered",
                "scatter_style": "lineMarker",
                "vary_colors": False,
                "gap_width": 150,
                "ax_ids": [],
            }

            bar_dir = sub_elem.find('c:barDir', namespaces=self._c_ns)
            if bar_dir is not None and bar_dir.get('val'):
                sub_chart_info["bar_direction"] = bar_dir.get('val')
                if len(chart._sub_charts) == 0:
                    try:
                        chart.bar_direction = bar_dir.get('val')
                    except ValueError:
                        pass

            grouping = sub_elem.find('c:grouping', namespaces=self._c_ns)
            if grouping is not None and grouping.get('val'):
                sub_chart_info["grouping"] = grouping.get('val')

            scatter_style = sub_elem.find('c:scatterStyle', namespaces=self._c_ns)
            if scatter_style is not None and scatter_style.get('val'):
                sub_chart_info["scatter_style"] = scatter_style.get('val')
                chart._scatter_style = scatter_style.get('val')

            vary_colors = sub_elem.find('c:varyColors', namespaces=self._c_ns)
            if vary_colors is not None:
                sub_chart_info["vary_colors"] = vary_colors.get('val', '0') in ('1', 'true', 'True')

            gap_width = sub_elem.find('c:gapWidth', namespaces=self._c_ns)
            if gap_width is not None and gap_width.get('val'):
                try:
                    sub_chart_info["gap_width"] = int(gap_width.get('val'))
                except ValueError:
                    pass

            for ax_id_elem in sub_elem.findall('c:axId', namespaces=self._c_ns):
                try:
                    sub_chart_info["ax_ids"].append(int(ax_id_elem.get('val', 0)))
                except ValueError:
                    pass

            # Load all series for this sub-chart
            for ser in sub_elem.findall('c:ser', namespaces=self._c_ns):
                ser_list_idx = len(chart.n_series._series)
                self._load_series(chart, ser, chart_type=sub_type)
                if len(chart.n_series._series) > ser_list_idx:
                    sub_chart_info["series"].append(ser_list_idx)

            chart._sub_charts.append(sub_chart_info)

        # Load all axes from the plot area
        self._load_axes(chart, plot_area)

        # Set primary chart properties from first sub-chart for convenience
        if chart._sub_charts:
            first = chart._sub_charts[0]
            try:
                chart.grouping = first["grouping"]
            except ValueError:
                pass

    def _load_series(self, chart, ser_elem, chart_type=None):
        """Loads a single series element into chart.n_series."""
        # Determine series idx and order
        idx_elem = ser_elem.find('c:idx', namespaces=self._c_ns)
        order_elem = ser_elem.find('c:order', namespaces=self._c_ns)
        series_idx = int(idx_elem.get('val', 0)) if idx_elem is not None else None
        series_order = int(order_elem.get('val', 0)) if order_elem is not None else None

        if chart_type == ChartType.SCATTER:
            x_values_formula = self._extract_formula(ser_elem, 'c:xVal')
            values_formula = self._extract_formula(ser_elem, 'c:yVal')
            bubble_sizes = self._extract_numeric_source(
                ser_elem, 'c:bubbleSize'
            )
            if not values_formula:
                return
            category_formula = None
        else:
            values_formula = self._extract_formula(ser_elem, 'c:val')
            if not values_formula:
                return
            category_formula = self._extract_formula(ser_elem, 'c:cat')
            x_values_formula = None
            bubble_sizes = None

        series_name = self._extract_series_name(ser_elem)

        chart.n_series.add(
            values_formula,
            category_data=category_formula,
            name=series_name,
            chart_type=chart_type,
            x_values=x_values_formula,
            series_idx=series_idx,
            series_order=series_order,
            bubble_sizes=bubble_sizes,
        )
        series = chart.n_series[-1]

        bubble_3d = ser_elem.find('c:bubble3D', namespaces=self._c_ns)
        if bubble_3d is not None:
            series._bubble_3d = bubble_3d.get('val', '1') not in (
                '0', 'false', 'False'
            )

        # Per-series smooth (scatter/line: <c:smooth val="0/1"/>)
        smooth_elem = ser_elem.find('c:smooth', namespaces=self._c_ns)
        if smooth_elem is not None:
            series._smooth = smooth_elem.get('val', '0') in ('1', 'true', 'True')

        # Load error bars
        for errBars_elem in ser_elem.findall('c:errBars', namespaces=self._c_ns):
            eb = self._load_error_bars(errBars_elem)
            if eb is not None:
                series._error_bars.append(eb)

    def _load_error_bars(self, errBars_elem):
        """Parses a <c:errBars> element into a ChartErrorBars object."""
        eb = ChartErrorBars()

        err_dir = errBars_elem.find('c:errDir', namespaces=self._c_ns)
        if err_dir is not None and err_dir.get('val'):
            eb.direction = err_dir.get('val')

        err_bar_type = errBars_elem.find('c:errBarType', namespaces=self._c_ns)
        if err_bar_type is not None and err_bar_type.get('val'):
            eb.bar_type = err_bar_type.get('val')

        err_val_type = errBars_elem.find('c:errValType', namespaces=self._c_ns)
        if err_val_type is not None and err_val_type.get('val'):
            eb.val_type = err_val_type.get('val')

        no_end_cap = errBars_elem.find('c:noEndCap', namespaces=self._c_ns)
        if no_end_cap is not None:
            eb.no_end_cap = no_end_cap.get('val', '0') in ('1', 'true', 'True')

        # Fixed value
        val_elem = errBars_elem.find('c:val', namespaces=self._c_ns)
        if val_elem is not None and val_elem.get('val') is not None:
            try:
                eb.val = float(val_elem.get('val'))
            except ValueError:
                pass

        # Custom plus/minus formulas
        plus_elem = errBars_elem.find('c:plus/c:numRef/c:f', namespaces=self._c_ns)
        if plus_elem is not None and plus_elem.text:
            eb.plus_formula = plus_elem.text.strip()

        minus_elem = errBars_elem.find('c:minus/c:numRef/c:f', namespaces=self._c_ns)
        if minus_elem is not None and minus_elem.text:
            eb.minus_formula = minus_elem.text.strip()

        # Line styling on spPr
        sp_pr = errBars_elem.find('c:spPr', namespaces=self._c_ns)
        if sp_pr is not None:
            ln = sp_pr.find('a:ln', namespaces=self._c_ns)
            if ln is not None:
                w = ln.get('w')
                if w is not None:
                    try:
                        eb.line_width = int(w)
                    except ValueError:
                        pass
                solid = ln.find('a:solidFill/a:srgbClr', namespaces=self._c_ns)
                if solid is not None and solid.get('val'):
                    eb.line_color = solid.get('val')

        return eb

    def _load_axes(self, chart, plot_area):
        """Loads all axis elements from a plotArea element into chart.axes."""
        c_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        axis_tag_map = {
            f'{{{c_uri}}}catAx': 'cat',
            f'{{{c_uri}}}valAx': 'val',
            f'{{{c_uri}}}serAx': 'ser',
            f'{{{c_uri}}}dateAx': 'date',
        }
        for child in plot_area:
            ax_type = axis_tag_map.get(child.tag)
            if ax_type is None:
                continue
            ax = ChartAxis()
            ax.axis_type = ax_type

            ax_id_elem = child.find('c:axId', namespaces=self._c_ns)
            if ax_id_elem is not None and ax_id_elem.get('val'):
                try:
                    ax.axis_id = int(ax_id_elem.get('val'))
                except ValueError:
                    pass

            orient = child.find('c:scaling/c:orientation', namespaces=self._c_ns)
            if orient is not None and orient.get('val'):
                ax.orientation = orient.get('val')

            scaling = child.find('c:scaling', namespaces=self._c_ns)
            if scaling is not None:
                min_elem = scaling.find('c:min', namespaces=self._c_ns)
                if min_elem is not None and min_elem.get('val') is not None:
                    try:
                        ax.min_val = float(min_elem.get('val'))
                    except ValueError:
                        pass
                max_elem = scaling.find('c:max', namespaces=self._c_ns)
                if max_elem is not None and max_elem.get('val') is not None:
                    try:
                        ax.max_val = float(max_elem.get('val'))
                    except ValueError:
                        pass

            delete_elem = child.find('c:delete', namespaces=self._c_ns)
            if delete_elem is not None:
                ax.deleted = delete_elem.get('val', '0') in ('1', 'true', 'True')

            ax_pos = child.find('c:axPos', namespaces=self._c_ns)
            if ax_pos is not None and ax_pos.get('val'):
                ax.position = ax_pos.get('val')

            num_fmt = child.find('c:numFmt', namespaces=self._c_ns)
            if num_fmt is not None:
                if num_fmt.get('formatCode'):
                    ax.num_fmt = num_fmt.get('formatCode')
                ax.source_linked = num_fmt.get('sourceLinked', '1') in ('1', 'true', 'True')

            major_tick = child.find('c:majorTickMark', namespaces=self._c_ns)
            if major_tick is not None and major_tick.get('val'):
                ax.major_tick_mark = major_tick.get('val')

            minor_tick = child.find('c:minorTickMark', namespaces=self._c_ns)
            if minor_tick is not None and minor_tick.get('val'):
                ax.minor_tick_mark = minor_tick.get('val')

            tick_lbl = child.find('c:tickLblPos', namespaces=self._c_ns)
            if tick_lbl is not None and tick_lbl.get('val'):
                ax.tick_lbl_pos = tick_lbl.get('val')

            cross_ax = child.find('c:crossAx', namespaces=self._c_ns)
            if cross_ax is not None and cross_ax.get('val'):
                try:
                    ax.cross_ax = int(cross_ax.get('val'))
                except ValueError:
                    pass

            crosses = child.find('c:crosses', namespaces=self._c_ns)
            if crosses is not None and crosses.get('val'):
                ax.crosses = crosses.get('val')

            auto_elem = child.find('c:auto', namespaces=self._c_ns)
            if auto_elem is not None:
                ax.auto = auto_elem.get('val', '1') in ('1', 'true', 'True')

            lbl_algn = child.find('c:lblAlgn', namespaces=self._c_ns)
            if lbl_algn is not None and lbl_algn.get('val'):
                ax.lbl_algn = lbl_algn.get('val')

            lbl_offset = child.find('c:lblOffset', namespaces=self._c_ns)
            if lbl_offset is not None and lbl_offset.get('val'):
                try:
                    ax.lbl_offset = int(lbl_offset.get('val'))
                except ValueError:
                    pass

            cross_between = child.find('c:crossBetween', namespaces=self._c_ns)
            if cross_between is not None and cross_between.get('val'):
                ax.cross_between = cross_between.get('val')

            chart._axes.append(ax)

    def _load_view3d(self, chart, view3d_elem):
        """Loads view3D settings into a chart."""
        rot_x = view3d_elem.find('c:rotX', namespaces=self._c_ns)
        if rot_x is not None and rot_x.get('val'):
            try:
                chart.view_3d.rotation_x = int(rot_x.get('val'))
            except ValueError:
                pass

        rot_y = view3d_elem.find('c:rotY', namespaces=self._c_ns)
        if rot_y is not None and rot_y.get('val'):
            try:
                chart.view_3d.rotation_y = int(rot_y.get('val'))
            except ValueError:
                pass

        right_angle = view3d_elem.find('c:rAngAx', namespaces=self._c_ns)
        if right_angle is not None:
            chart.view_3d.right_angle_axes = right_angle.get('val', '0') in ('1', 'true', 'True')

        perspective = view3d_elem.find('c:perspective', namespaces=self._c_ns)
        if perspective is not None and perspective.get('val'):
            try:
                chart.view_3d.perspective = int(perspective.get('val'))
            except ValueError:
                pass

        height_percent = view3d_elem.find('c:hPercent', namespaces=self._c_ns)
        if height_percent is not None and height_percent.get('val'):
            try:
                chart.view_3d.height_percent = int(height_percent.get('val'))
            except ValueError:
                pass

        depth_percent = view3d_elem.find('c:depthPercent', namespaces=self._c_ns)
        if depth_percent is not None and depth_percent.get('val'):
            try:
                chart.view_3d.depth_percent = int(depth_percent.get('val'))
            except ValueError:
                pass

    def _load_chart_ex_settings(self, chart, chart_root, plot_region_elem, worksheet):
        chart.vary_colors = True
        title_val = chart_root.find('.//cx:chart/cx:title//cx:v', namespaces=self._cx_ns)
        if title_val is not None:
            chart.title = title_val.text if title_val.text is not None else ""
        else:
            title_runs = chart_root.findall(
                './/cx:chart/cx:title//a:t', namespaces=self._cx_ns
            )
            if title_runs:
                chart.title = ''.join(run.text or '' for run in title_runs)

        legend_elem = chart_root.find('.//cx:chart/cx:legend', namespaces=self._cx_ns)
        if legend_elem is not None:
            chart.show_legend = True
            if legend_elem.get('pos'):
                try:
                    chart.legend_position = legend_elem.get('pos')
                except ValueError:
                    pass
            # Support both 'visible' (our old format) and absence of element (hidden)
            visible_attr = legend_elem.get('visible')
            if visible_attr is not None:
                chart.show_legend = visible_attr not in ('0', 'false', 'False')
        else:
            chart.show_legend = False

        defined_names = self._get_defined_name_map(worksheet)
        data_map = {}
        for data_elem in chart_root.findall('.//cx:chartData/cx:data', namespaces=self._cx_ns):
            data_id = data_elem.get('id')
            if data_id is None:
                continue
            cat_formula = self._resolve_chart_ex_formula(data_elem.find('cx:strDim/cx:f', namespaces=self._cx_ns), defined_names)
            val_formula = self._resolve_chart_ex_formula(data_elem.find('cx:numDim/cx:f', namespaces=self._cx_ns), defined_names)
            data_map[str(data_id)] = (cat_formula, val_formula)

        quartile_val = None
        for ser_elem in plot_region_elem.findall('cx:series', namespaces=self._cx_ns):
            data_id_elem = ser_elem.find('cx:dataId', namespaces=self._cx_ns)
            if data_id_elem is None:
                continue
            data_id = data_id_elem.get('val')
            if data_id is None:
                continue
            cat_formula, val_formula = data_map.get(str(data_id), (None, None))
            if not val_formula:
                continue
            tx_val = ser_elem.find('cx:tx/cx:txData/cx:v', namespaces=self._cx_ns)
            name = tx_val.text if tx_val is not None else None
            chart.n_series.add(val_formula, category_data=cat_formula, name=name)
            series = chart.n_series[-1]
            series.hidden = ser_elem.get('hidden', '0') in ('1', 'true', 'True')
            has_subtotal = ser_elem.find('cx:layoutPr/cx:subtotals', namespaces=self._cx_ns) is not None
            series.is_subtotal = has_subtotal
            if has_subtotal:
                chart.has_subtotals = True
            stats = ser_elem.find('cx:layoutPr/cx:statistics', namespaces=self._cx_ns)
            if stats is not None and stats.get('quartileMethod'):
                quartile_val = stats.get('quartileMethod')
            visibility = ser_elem.find('cx:layoutPr/cx:visibility', namespaces=self._cx_ns)
            if visibility is not None:
                chart._box_show_mean_line = visibility.get('meanLine', '0') in ('1', 'true', 'True')
                chart._box_show_mean_marker = visibility.get('meanMarker', '1') in ('1', 'true', 'True')
                chart._box_show_inner_points = visibility.get('nonoutliers', '0') in ('1', 'true', 'True')
                chart._box_show_outlier_points = visibility.get('outliers', '1') in ('1', 'true', 'True')
            connector_lines = ser_elem.find('cx:layoutPr/cx:connectorLines', namespaces=self._cx_ns)
            if connector_lines is not None:
                chart.show_connector_lines = connector_lines.get('visible', '1') not in ('0', 'false', 'False')
            # Histogram binning settings (from first non-hidden series)
            binning = ser_elem.find('cx:layoutPr/cx:binning', namespaces=self._cx_ns)
            if binning is not None:
                if binning.get('intervalClosed'):
                    chart._histogram_interval_closed = binning.get('intervalClosed')
                if binning.get('count') is not None:
                    try:
                        chart._histogram_bin_count = int(binning.get('count'))
                    except ValueError:
                        pass
                if binning.get('size') is not None:
                    try:
                        chart._histogram_bin_size = float(binning.get('size'))
                    except ValueError:
                        pass
                if binning.get('overflow') is not None:
                    try:
                        chart._histogram_overflow = float(binning.get('overflow'))
                    except ValueError:
                        pass
                if binning.get('underflow') is not None:
                    try:
                        chart._histogram_underflow = float(binning.get('underflow'))
                    except ValueError:
                        pass
        if quartile_val:
            try:
                chart.quartile_method = quartile_val
            except ValueError:
                pass
        cat_axis_scaling = plot_region_elem.find('cx:axis[@id="0"]/cx:catScaling', namespaces=self._cx_ns)
        if cat_axis_scaling is not None and cat_axis_scaling.get('gapWidth') is not None:
            try:
                chart._box_gap_width = int(cat_axis_scaling.get('gapWidth'))
            except (TypeError, ValueError):
                pass

    def _detect_chart_ex_type(self, plot_region_elem):
        layout_ids = {
            (ser.get('layoutId') or '').strip()
            for ser in plot_region_elem.findall('cx:series', namespaces=self._cx_ns)
            if ser.get('layoutId')
        }
        lower_layouts = {layout.lower() for layout in layout_ids}
        if "waterfall" in lower_layouts:
            return ChartType.WATERFALL
        if "boxwhisker" in lower_layouts:
            return ChartType.BOX_WHISKER
        if "treemap" in lower_layouts:
            return ChartType.TREEMAP
        if "sunburst" in lower_layouts:
            return ChartType.SUNBURST
        if "clusteredcolumn" in lower_layouts:
            return ChartType.HISTOGRAM
        if "funnel" in lower_layouts:
            return ChartType.FUNNEL
        if "regionmap" in lower_layouts:
            return ChartType.MAP
        return ChartType.BOX_WHISKER

    def _get_defined_name_map(self, worksheet):
        name_map = {}
        workbook = getattr(worksheet, '_workbook', None)
        if workbook is None:
            return name_map
        defined_names = getattr(getattr(workbook, 'properties', None), 'defined_names', None)
        if defined_names is None:
            return name_map
        for dn in defined_names:
            if not getattr(dn, 'name', None):
                continue
            formula = getattr(dn, 'refers_to', None)
            if formula:
                name_map[dn.name] = formula
        return name_map

    def _resolve_chart_ex_formula(self, formula_elem, defined_name_map):
        if formula_elem is None or formula_elem.text is None:
            return None
        key = formula_elem.text.strip()
        if not key:
            return None
        return defined_name_map.get(key, key)

    def _extract_formula(self, series_elem, prefix):
        num_ref = series_elem.find(f'{prefix}/c:numRef/c:f', namespaces=self._c_ns)
        if num_ref is not None and num_ref.text:
            return num_ref.text.strip()
        str_ref = series_elem.find(f'{prefix}/c:strRef/c:f', namespaces=self._c_ns)
        if str_ref is not None and str_ref.text:
            return str_ref.text.strip()
        return None

    def _extract_numeric_source(self, series_elem, prefix):
        formula = self._extract_formula(series_elem, prefix)
        if formula:
            return formula
        values = []
        for point in series_elem.findall(
            f'{prefix}/c:numLit/c:pt', namespaces=self._c_ns
        ):
            value = point.find('c:v', namespaces=self._c_ns)
            if value is None or value.text is None:
                continue
            try:
                values.append(float(value.text))
            except ValueError:
                values.append(0.0)
        return tuple(values)

    def _extract_series_name(self, series_elem):
        text_val = series_elem.find('c:tx/c:v', namespaces=self._c_ns)
        if text_val is not None and text_val.text:
            return text_val.text

        cache_val = series_elem.find('c:tx/c:strRef/c:strCache/c:pt/c:v', namespaces=self._c_ns)
        if cache_val is not None and cache_val.text:
            return cache_val.text

        formula_val = series_elem.find('c:tx/c:strRef/c:f', namespaces=self._c_ns)
        if formula_val is not None and formula_val.text:
            return formula_val.text

        return None

    def _get_anchor_int(self, anchor_elem, xpath, default=None):
        node = anchor_elem.find(xpath, namespaces=self._xdr_ns)
        if node is None or node.text is None:
            return default
        try:
            return int(float(node.text))
        except ValueError:
            return default

    def _read_relationships(self, zipf, rels_path):
        rels = {}
        rels_root = self._try_read_xml(zipf, rels_path)
        if rels_root is None:
            return rels
        for rel in rels_root.findall('rel:Relationship', namespaces=self._rels_ns):
            rel_id = rel.get('Id')
            target = rel.get('Target')
            if rel_id and target:
                rels[rel_id] = target
        return rels

    def _read_relationship_types(self, zipf, rels_path):
        rel_types = {}
        rels_root = self._try_read_xml(zipf, rels_path)
        if rels_root is None:
            return rel_types
        for rel in rels_root.findall('rel:Relationship', namespaces=self._rels_ns):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type')
            if rel_id and rel_type:
                rel_types[rel_id] = rel_type
        return rel_types

    def _try_read_xml(self, zipf, zip_path):
        try:
            return ET.fromstring(zipf.read(zip_path))
        except KeyError:
            return None
        except ET.ParseError:
            return None

    def _try_read_bytes(self, zipf, zip_path):
        try:
            return zipf.read(zip_path)
        except KeyError:
            return None

    def _resolve_target(self, base_part_path, target):
        base_dir = posixpath.dirname(base_part_path)
        combined = posixpath.normpath(posixpath.join(base_dir, target))
        return combined.lstrip('/')

    def _rels_path_for_part(self, part_path):
        part_dir = posixpath.dirname(part_path)
        part_name = posixpath.basename(part_path)
        return posixpath.join(part_dir, '_rels', f'{part_name}.rels')

