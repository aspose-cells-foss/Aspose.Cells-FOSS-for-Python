"""
Aspose.Cells for Python - Chart Module

This module provides basic chart APIs for creating line charts in worksheets
and persisting chart settings to XLSX files.
"""

from enum import IntEnum


class ChartType(IntEnum):
    """Supported chart types."""
    LINE = 0
    BAR = 1
    PIE = 2
    AREA = 3
    BOX_WHISKER = 4
    WATERFALL = 5
    COMBO = 6
    SCATTER = 7
    STOCK = 8
    SURFACE = 9
    RADAR = 10
    TREEMAP = 11
    SUNBURST = 12
    HISTOGRAM = 13
    FUNNEL = 14
    MAP = 15


class ChartErrorBars:
    """
    Represents error bars attached to a chart series.

    Attributes:
        direction (str): Axis direction - 'x' or 'y'.
        bar_type (str): 'both', 'minus', or 'plus'.
        val_type (str): 'fixedVal', 'percentage', 'stdDev', 'stdErr', or 'cust'.
        no_end_cap (bool): Whether end caps are suppressed.
        val (float): Fixed/percentage/stdDev value (used when val_type != 'cust').
        plus_formula (str|None): Formula string for custom plus values.
        minus_formula (str|None): Formula string for custom minus values.
        line_width (int|None): Line width in EMU (None = default).
        line_color (str|None): Hex color string e.g. '7F7F7F' (None = default).
    """

    def __init__(self):
        self.direction = "y"
        self.bar_type = "both"
        self.val_type = "fixedVal"
        self.no_end_cap = False
        self.val = 0.0
        self.plus_formula = None
        self.minus_formula = None
        self.line_width = None
        self.line_color = None

    def copy(self):
        eb = ChartErrorBars()
        eb.direction = self.direction
        eb.bar_type = self.bar_type
        eb.val_type = self.val_type
        eb.no_end_cap = self.no_end_cap
        eb.val = self.val
        eb.plus_formula = self.plus_formula
        eb.minus_formula = self.minus_formula
        eb.line_width = self.line_width
        eb.line_color = self.line_color
        return eb


class ChartAxis:
    """
    Represents a chart axis (category, value, or series).

    Attributes:
        axis_id (int): Unique axis identifier used to link chart <axId> references.
        axis_type (str): 'cat', 'val', or 'ser'.
        orientation (str): 'minMax' (default) or 'maxMin' (reversed).
        position (str): Axis position - 'l', 'r', 't', or 'b'.
        deleted (bool): Whether the axis is hidden/deleted.
        crosses (str): 'autoZero', 'max', or 'min'.
        cross_ax (int): axis_id of the crossing axis.
        num_fmt (str): Number format code, e.g. 'General' or '#,##0.0,'.
        source_linked (bool): Whether num_fmt is linked to source data.
        tick_lbl_pos (str): 'nextTo', 'high', 'low', or 'none'.
        major_tick_mark (str): 'none', 'in', 'out', or 'cross'.
        minor_tick_mark (str): 'none', 'in', 'out', or 'cross'.
        min_val (float|None): Axis minimum scale value (None = auto).
        max_val (float|None): Axis maximum scale value (None = auto).
        auto (bool): Whether axis labels are auto (catAx only).
        lbl_algn (str): Label alignment 'ctr', 'l', or 'r' (catAx only).
        lbl_offset (int): Label offset percentage (catAx only).
        cross_between (str): 'between' or 'midCat' (valAx only).
    """

    def __init__(self):
        self.axis_id = 0
        self.axis_type = "cat"
        self.orientation = "minMax"
        self.position = "b"
        self.deleted = False
        self.crosses = "autoZero"
        self.cross_ax = 0
        self.num_fmt = "General"
        self.source_linked = True
        self.tick_lbl_pos = "nextTo"
        self.major_tick_mark = "out"
        self.minor_tick_mark = "none"
        self.min_val = None
        self.max_val = None
        self.auto = True
        self.lbl_algn = "ctr"
        self.lbl_offset = 100
        self.cross_between = "between"

    def copy(self):
        ax = ChartAxis()
        ax.axis_id = self.axis_id
        ax.axis_type = self.axis_type
        ax.orientation = self.orientation
        ax.position = self.position
        ax.deleted = self.deleted
        ax.crosses = self.crosses
        ax.cross_ax = self.cross_ax
        ax.num_fmt = self.num_fmt
        ax.source_linked = self.source_linked
        ax.tick_lbl_pos = self.tick_lbl_pos
        ax.major_tick_mark = self.major_tick_mark
        ax.minor_tick_mark = self.minor_tick_mark
        ax.min_val = self.min_val
        ax.max_val = self.max_val
        ax.auto = self.auto
        ax.lbl_algn = self.lbl_algn
        ax.lbl_offset = self.lbl_offset
        ax.cross_between = self.cross_between
        return ax


class ChartSeries:
    """Represents a single chart series."""

    def __init__(self, values, category_data=None, name=None, chart=None,
                 chart_type=None, x_values=None, series_idx=None, series_order=None,
                 bubble_sizes=None):
        self._values = values
        self._category_data = category_data
        self._name = name
        self._chart = chart
        self._hidden = False
        self._is_subtotal = False
        # Combo/scatter extensions
        self._chart_type = chart_type   # ChartType for per-series type in combo charts (None = inherit from chart)
        self._x_values = x_values       # For scatter xVal formula (None = use category_data)
        self._bubble_sizes = bubble_sizes  # Bubble size formula or cached literal values
        self._bubble_3d = False
        self._error_bars = []           # List of ChartErrorBars
        self._series_idx = series_idx   # The <c:idx val="..."> value
        self._series_order = series_order  # The <c:order val="..."> value
        self._smooth = False            # Per-series smooth for scatter/line (<c:smooth val="1"/>)

    @property
    def values(self):
        return self._values

    @values.setter
    def values(self, value):
        self._values = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def category_data(self):
        return self._category_data

    @category_data.setter
    def category_data(self, value):
        self._category_data = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def chart_type(self):
        """Per-series chart type (relevant for combo charts). None = inherit from chart."""
        return self._chart_type

    @chart_type.setter
    def chart_type(self, value):
        self._chart_type = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def x_values(self):
        """X-axis values formula for scatter series (xVal)."""
        return self._x_values

    @x_values.setter
    def x_values(self, value):
        self._x_values = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def bubble_sizes(self):
        """Bubble-size values or formula for bubble chart series."""
        return self._bubble_sizes

    @bubble_sizes.setter
    def bubble_sizes(self, value):
        self._bubble_sizes = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def error_bars(self):
        """List of ChartErrorBars attached to this series."""
        return self._error_bars

    @property
    def series_idx(self):
        """The series <c:idx val="..."> value."""
        return self._series_idx

    @series_idx.setter
    def series_idx(self, value):
        self._series_idx = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def series_order(self):
        """The series <c:order val="..."> value."""
        return self._series_order

    @series_order.setter
    def series_order(self, value):
        self._series_order = value
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def smooth(self):
        """Per-series smooth line flag for scatter/line charts (<c:smooth val='1'/>)."""
        return self._smooth

    @smooth.setter
    def smooth(self, value):
        self._smooth = bool(value)
        if self._chart is not None:
            self._chart._mark_dirty()

    def add_error_bars(self, direction="y", bar_type="both", val_type="fixedVal",
                       val=0.0, plus_formula=None, minus_formula=None,
                       no_end_cap=False, line_width=None, line_color=None):
        """
        Adds error bars to this series.

        Args:
            direction (str): 'x' or 'y'.
            bar_type (str): 'both', 'minus', or 'plus'.
            val_type (str): 'fixedVal', 'percentage', 'stdDev', 'stdErr', or 'cust'.
            val (float): Error value for non-custom types.
            plus_formula (str): Range formula for custom plus values.
            minus_formula (str): Range formula for custom minus values.
            no_end_cap (bool): Suppress end caps.
            line_width (int): Line width in EMU.
            line_color (str): Hex color string.

        Returns:
            ChartErrorBars: The created error bars object.
        """
        eb = ChartErrorBars()
        eb.direction = direction
        eb.bar_type = bar_type
        eb.val_type = val_type
        eb.val = val
        eb.plus_formula = plus_formula
        eb.minus_formula = minus_formula
        eb.no_end_cap = no_end_cap
        eb.line_width = line_width
        eb.line_color = line_color
        self._error_bars.append(eb)
        if self._chart is not None:
            self._chart._mark_dirty()
        return eb

    def copy(self, chart=None):
        cloned = ChartSeries(
            self._values, self._category_data, self._name, chart=chart,
            chart_type=self._chart_type, x_values=self._x_values,
            series_idx=self._series_idx, series_order=self._series_order,
            bubble_sizes=self._bubble_sizes,
        )
        cloned._hidden = self._hidden
        cloned._is_subtotal = self._is_subtotal
        cloned._error_bars = [eb.copy() for eb in self._error_bars]
        cloned._bubble_3d = self._bubble_3d
        return cloned

    @property
    def hidden(self):
        return self._hidden

    @hidden.setter
    def hidden(self, value):
        self._hidden = bool(value)
        if self._chart is not None:
            self._chart._mark_dirty()

    @property
    def is_subtotal(self):
        return self._is_subtotal

    @is_subtotal.setter
    def is_subtotal(self, value):
        self._is_subtotal = bool(value)
        if self._chart is not None:
            self._chart._mark_dirty()


class ChartView3D:
    """Represents chart-level 3D view settings."""

    def __init__(self, chart):
        self._chart = chart
        self._rotation_x = 15
        self._rotation_y = 20
        self._right_angle_axes = False
        self._perspective = 30
        self._height_percent = 100
        self._depth_percent = 100

    @property
    def rotation_x(self):
        return self._rotation_x

    @rotation_x.setter
    def rotation_x(self, value):
        self._rotation_x = int(value)
        self._chart._mark_dirty()

    @property
    def rotation_y(self):
        return self._rotation_y

    @rotation_y.setter
    def rotation_y(self, value):
        self._rotation_y = int(value)
        self._chart._mark_dirty()

    @property
    def right_angle_axes(self):
        return self._right_angle_axes

    @right_angle_axes.setter
    def right_angle_axes(self, value):
        self._right_angle_axes = bool(value)
        self._chart._mark_dirty()

    @property
    def perspective(self):
        return self._perspective

    @perspective.setter
    def perspective(self, value):
        self._perspective = int(value)
        self._chart._mark_dirty()

    @property
    def height_percent(self):
        return self._height_percent

    @height_percent.setter
    def height_percent(self, value):
        self._height_percent = int(value)
        self._chart._mark_dirty()

    @property
    def depth_percent(self):
        return self._depth_percent

    @depth_percent.setter
    def depth_percent(self, value):
        self._depth_percent = int(value)
        self._chart._mark_dirty()

    def copy(self, chart):
        view = ChartView3D(chart)
        view._rotation_x = self._rotation_x
        view._rotation_y = self._rotation_y
        view._right_angle_axes = self._right_angle_axes
        view._perspective = self._perspective
        view._height_percent = self._height_percent
        view._depth_percent = self._depth_percent
        return view


class NSeries:
    """Collection of series for a chart."""

    def __init__(self, chart):
        self._chart = chart
        self._series = []

    def add(self, area, is_vertical=False, category_data=None, name=None,
            chart_type=None, x_values=None, series_idx=None, series_order=None,
            bubble_sizes=None):
        """
        Adds a series to the chart.

        Args:
            area (str): A1 range for series values (or yVal for scatter).
            is_vertical (bool): Reserved for compatibility.
            category_data (str, optional): A1 range for category labels (or yVal for scatter cat).
            name (str, optional): Series name.
            chart_type (ChartType, optional): Series-level chart type for combo charts.
            x_values (str, optional): A1 range for scatter xVal.
            series_idx (int, optional): The <c:idx> value (auto-assigned if None).
            series_order (int, optional): The <c:order> value (auto-assigned if None).

        Returns:
            int: Index of the created series.
        """
        _ = is_vertical
        idx = series_idx if series_idx is not None else len(self._series)
        order = series_order if series_order is not None else len(self._series)
        ser = ChartSeries(
            area, category_data=category_data, name=name, chart=self._chart,
            chart_type=chart_type, x_values=x_values,
            series_idx=idx, series_order=order,
            bubble_sizes=bubble_sizes,
        )
        self._series.append(ser)
        self._chart._mark_dirty()
        return len(self._series) - 1

    def Add(self, area, is_vertical=False, category_data=None, name=None,
            chart_type=None, x_values=None, series_idx=None, series_order=None,
            bubble_sizes=None):
        """PascalCase alias of add()."""
        return self.add(area, is_vertical=is_vertical, category_data=category_data,
                        name=name, chart_type=chart_type, x_values=x_values,
                        series_idx=series_idx, series_order=series_order,
                        bubble_sizes=bubble_sizes)

    @property
    def count(self):
        return len(self._series)

    def __len__(self):
        return len(self._series)

    def __iter__(self):
        return iter(self._series)

    def __getitem__(self, index):
        return self._series[index]

    def copy(self, chart):
        new_collection = NSeries(chart)
        new_collection._series = [ser.copy(chart=chart) for ser in self._series]
        return new_collection


class Chart:
    """Represents a chart in a worksheet."""

    _LEGEND_MAP = {
        "right": "r",
        "left": "l",
        "top": "t",
        "bottom": "b",
        "top_right": "tr",
        "r": "r",
        "l": "l",
        "t": "t",
        "b": "b",
        "tr": "tr",
    }

    def __init__(self, chart_type, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        self._type = chart_type
        self._upper_left_row = int(upper_left_row)
        self._upper_left_column = int(upper_left_column)
        self._lower_right_row = int(lower_right_row)
        self._lower_right_column = int(lower_right_column)
        self._title = None
        self._category_data = None
        self._show_legend = True
        self._legend_position = "r"
        self._grouping = "standard"
        self._smooth = False
        self._bar_direction = "col"
        self._gap_width = 150
        self._overlap = 0
        self._vary_colors = False
        self._first_slice_angle = 0
        self._is_of_pie = False
        self._of_pie_type = "pie"
        self._second_pie_size = 75
        self._quartile_method = "exclusive"
        # Box & whisker settings (chartEx visibility + category gap width)
        self._box_show_mean_line = False
        self._box_show_mean_marker = True
        self._box_show_inner_points = False
        self._box_show_outlier_points = True
        self._box_gap_width = 1
        self._is_3d = False
        self._gap_depth = 150
        self._upper_left_row_offset = 0
        self._upper_left_column_offset = 0
        self._lower_right_row_offset = 0
        self._lower_right_column_offset = 0
        self._anchor_extent_width = None
        self._anchor_extent_height = None
        self._source_chart_xml = None
        self._source_chart_rels_xml = None
        self._source_chart_extra_parts = []
        self._source_chart_part_path = None
        self._source_chart_part_index = None
        self._source_chart_content_type = None
        self._source_is_chart_ex = False
        self._suspend_dirty_tracking = False
        self._show_connector_lines = False
        self._has_subtotals = False
        self._view_3d = ChartView3D(self)
        self._n_series = NSeries(self)
        # Combo chart support
        self._sub_charts = []   # List of dicts describing each sub-chart
        self._axes = []         # List of ChartAxis objects
        self._scatter_style = "lineMarker"
        self._is_bubble = False
        self._bubble_scale = 100.0
        self._disp_blanks_as = "gap"
        self._auto_title_deleted = False
        self._stock_style = "high_low_close"
        self._wireframe = False  # Surface chart wireframe mode
        self._radar_style = "marker"  # Radar chart style: standard, marker, filled
        # Histogram chart settings
        self._histogram_bin_count = None   # int or None (count-based binning)
        self._histogram_bin_size = None    # float or None (size-based binning)
        self._histogram_interval_closed = "r"  # "r" or "l"
        self._histogram_overflow = None    # float or None
        self._histogram_underflow = None   # float or None

    def _mark_dirty(self):
        if self._suspend_dirty_tracking:
            return
        self._source_chart_xml = None
        self._source_chart_rels_xml = None
        self._source_chart_extra_parts = []
        self._source_chart_part_path = None
        self._source_chart_part_index = None
        self._source_chart_content_type = None
        self._source_is_chart_ex = False

    @property
    def type(self):
        return self._type

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        self._title = str(value) if value is not None else None
        self._mark_dirty()

    @property
    def category_data(self):
        return self._category_data

    @category_data.setter
    def category_data(self, value):
        self._category_data = value
        self._mark_dirty()

    @property
    def show_legend(self):
        return self._show_legend

    @show_legend.setter
    def show_legend(self, value):
        self._show_legend = bool(value)
        self._mark_dirty()

    @property
    def legend_position(self):
        return self._legend_position

    @legend_position.setter
    def legend_position(self, value):
        key = str(value).lower().strip()
        if key not in self._LEGEND_MAP:
            raise ValueError("legend_position must be one of: right, left, top, bottom, top_right (or r/l/t/b/tr)")
        self._legend_position = self._LEGEND_MAP[key]
        self._mark_dirty()

    @property
    def smooth(self):
        return self._smooth

    @smooth.setter
    def smooth(self, value):
        self._smooth = bool(value)
        self._mark_dirty()

    @property
    def n_series(self):
        return self._n_series

    @property
    def NSeries(self):
        """PascalCase alias of n_series."""
        return self._n_series

    @property
    def grouping(self):
        return self._grouping

    @grouping.setter
    def grouping(self, value):
        val = str(value).strip()
        allowed = {"standard", "stacked", "percentStacked", "clustered"}
        if val not in allowed:
            raise ValueError("grouping must be one of: standard, stacked, percentStacked, clustered")
        self._grouping = val
        self._mark_dirty()

    @property
    def bar_direction(self):
        return self._bar_direction

    @bar_direction.setter
    def bar_direction(self, value):
        val = str(value).strip()
        allowed = {"bar", "col"}
        if val not in allowed:
            raise ValueError("bar_direction must be one of: bar, col")
        self._bar_direction = val
        self._mark_dirty()

    @property
    def gap_width(self):
        return self._gap_width

    @gap_width.setter
    def gap_width(self, value):
        gap = int(value)
        if gap < 0 or gap > 500:
            raise ValueError("gap_width must be in range [0, 500]")
        self._gap_width = gap
        self._mark_dirty()

    @property
    def overlap(self):
        return self._overlap

    @overlap.setter
    def overlap(self, value):
        ov = int(value)
        if ov < -100 or ov > 100:
            raise ValueError("overlap must be in range [-100, 100]")
        self._overlap = ov
        self._mark_dirty()

    @property
    def vary_colors(self):
        return self._vary_colors

    @vary_colors.setter
    def vary_colors(self, value):
        self._vary_colors = bool(value)
        self._mark_dirty()

    @property
    def first_slice_angle(self):
        return self._first_slice_angle

    @first_slice_angle.setter
    def first_slice_angle(self, value):
        angle = int(value)
        if angle < 0 or angle > 360:
            raise ValueError("first_slice_angle must be in range [0, 360]")
        self._first_slice_angle = angle
        self._mark_dirty()

    @property
    def is_of_pie(self):
        return self._is_of_pie

    @is_of_pie.setter
    def is_of_pie(self, value):
        self._is_of_pie = bool(value)
        self._mark_dirty()

    @property
    def of_pie_type(self):
        return self._of_pie_type

    @of_pie_type.setter
    def of_pie_type(self, value):
        val = str(value).strip()
        if val not in {"pie", "bar"}:
            raise ValueError("of_pie_type must be one of: pie, bar")
        self._of_pie_type = val
        self._mark_dirty()

    @property
    def second_pie_size(self):
        return self._second_pie_size

    @second_pie_size.setter
    def second_pie_size(self, value):
        size = int(value)
        if size < 5 or size > 200:
            raise ValueError("second_pie_size must be in range [5, 200]")
        self._second_pie_size = size
        self._mark_dirty()

    @property
    def quartile_method(self):
        return self._quartile_method

    @quartile_method.setter
    def quartile_method(self, value):
        val = str(value).strip()
        if val not in {"inclusive", "exclusive"}:
            raise ValueError("quartile_method must be one of: inclusive, exclusive")
        self._quartile_method = val
        self._mark_dirty()

    @property
    def box_show_mean_line(self):
        return self._box_show_mean_line

    @box_show_mean_line.setter
    def box_show_mean_line(self, value):
        self._box_show_mean_line = bool(value)
        self._mark_dirty()

    @property
    def box_show_mean_marker(self):
        return self._box_show_mean_marker

    @box_show_mean_marker.setter
    def box_show_mean_marker(self, value):
        self._box_show_mean_marker = bool(value)
        self._mark_dirty()

    @property
    def box_show_inner_points(self):
        return self._box_show_inner_points

    @box_show_inner_points.setter
    def box_show_inner_points(self, value):
        self._box_show_inner_points = bool(value)
        self._mark_dirty()

    @property
    def box_show_outlier_points(self):
        return self._box_show_outlier_points

    @box_show_outlier_points.setter
    def box_show_outlier_points(self, value):
        self._box_show_outlier_points = bool(value)
        self._mark_dirty()

    @property
    def box_gap_width(self):
        return self._box_gap_width

    @box_gap_width.setter
    def box_gap_width(self, value):
        gap = int(value)
        if gap < 0:
            raise ValueError("box_gap_width must be >= 0")
        self._box_gap_width = gap
        self._mark_dirty()

    @property
    def is_3d(self):
        return self._is_3d

    @is_3d.setter
    def is_3d(self, value):
        self._is_3d = bool(value)
        self._mark_dirty()

    @property
    def gap_depth(self):
        return self._gap_depth

    @gap_depth.setter
    def gap_depth(self, value):
        self._gap_depth = int(value)
        self._mark_dirty()

    @property
    def view_3d(self):
        return self._view_3d

    @property
    def View3D(self):
        """PascalCase alias of view_3d."""
        return self._view_3d

    @property
    def show_connector_lines(self):
        return self._show_connector_lines

    @show_connector_lines.setter
    def show_connector_lines(self, value):
        self._show_connector_lines = bool(value)
        self._mark_dirty()

    @property
    def has_subtotals(self):
        return self._has_subtotals

    @has_subtotals.setter
    def has_subtotals(self, value):
        self._has_subtotals = bool(value)
        self._mark_dirty()

    @property
    def sub_charts(self):
        """
        List of sub-chart descriptors for combo charts.

        Each entry is a dict with keys:
            'type' (ChartType): Sub-chart type (BAR, SCATTER, LINE, etc.)
            'series' (list[int]): Indices into n_series for this sub-chart.
            'bar_direction' (str): For BAR sub-charts ('bar' or 'col').
            'grouping' (str): For BAR/LINE/AREA sub-charts.
            'scatter_style' (str): For SCATTER sub-charts.
            'vary_colors' (bool): varyColors flag.
            'gap_width' (int): For BAR sub-charts.
            'ax_ids' (list[int]): Axis IDs used by this sub-chart.
        """
        return self._sub_charts

    @property
    def axes(self):
        """List of ChartAxis objects defining all axes for the chart."""
        return self._axes

    @property
    def scatter_style(self):
        return self._scatter_style

    @scatter_style.setter
    def scatter_style(self, value):
        self._scatter_style = str(value)
        self._mark_dirty()

    @property
    def wireframe(self):
        """Whether the surface chart uses wireframe display mode (<c:wireframe val='1'/>)."""
        return self._wireframe

    @wireframe.setter
    def wireframe(self, value):
        self._wireframe = bool(value)
        self._mark_dirty()

    @property
    def radar_style(self):
        """Radar chart style: 'standard', 'marker', or 'filled'."""
        return self._radar_style

    @radar_style.setter
    def radar_style(self, value):
        val = str(value).strip()
        if val not in {"standard", "marker", "filled"}:
            raise ValueError("radar_style must be one of: standard, marker, filled")
        self._radar_style = val
        self._mark_dirty()

    @property
    def histogram_bin_count(self):
        """Number of bins for count-based binning (int or None for auto/size-based)."""
        return self._histogram_bin_count

    @histogram_bin_count.setter
    def histogram_bin_count(self, value):
        self._histogram_bin_count = int(value) if value is not None else None
        self._mark_dirty()

    @property
    def histogram_bin_size(self):
        """Bin width for size-based binning (float or None for auto/count-based)."""
        return self._histogram_bin_size

    @histogram_bin_size.setter
    def histogram_bin_size(self, value):
        self._histogram_bin_size = float(value) if value is not None else None
        self._mark_dirty()

    @property
    def histogram_interval_closed(self):
        """Which side of each bin interval is closed: 'r' (right) or 'l' (left)."""
        return self._histogram_interval_closed

    @histogram_interval_closed.setter
    def histogram_interval_closed(self, value):
        val = str(value).strip()
        if val not in {"r", "l"}:
            raise ValueError("histogram_interval_closed must be 'r' or 'l'")
        self._histogram_interval_closed = val
        self._mark_dirty()

    @property
    def histogram_overflow(self):
        """Overflow bin boundary value (float or None)."""
        return self._histogram_overflow

    @histogram_overflow.setter
    def histogram_overflow(self, value):
        self._histogram_overflow = float(value) if value is not None else None
        self._mark_dirty()

    @property
    def histogram_underflow(self):
        """Underflow bin boundary value (float or None)."""
        return self._histogram_underflow

    @histogram_underflow.setter
    def histogram_underflow(self, value):
        self._histogram_underflow = float(value) if value is not None else None
        self._mark_dirty()

    @property
    def disp_blanks_as(self):
        return self._disp_blanks_as

    @disp_blanks_as.setter
    def disp_blanks_as(self, value):
        self._disp_blanks_as = str(value)
        self._mark_dirty()

    @property
    def stock_style(self):
        return self._stock_style

    @stock_style.setter
    def stock_style(self, value):
        val = str(value).strip()
        allowed = {
            "high_low_close",
            "open_high_low_close",
            "volume_high_low_close",
            "volume_open_high_low_close",
        }
        if val not in allowed:
            raise ValueError(
                "stock_style must be one of: high_low_close, open_high_low_close, "
                "volume_high_low_close, volume_open_high_low_close"
            )
        self._stock_style = val
        self._mark_dirty()

    def add_series(self, values, category_data=None, name=None,
                   chart_type=None, x_values=None):
        """Convenience method to add a series."""
        return self._n_series.add(values, category_data=category_data, name=name,
                                  chart_type=chart_type, x_values=x_values)

    def add_axis(self, axis_type="cat", axis_id=None, **kwargs):
        """
        Adds an axis to the chart and returns it.

        Args:
            axis_type (str): 'cat', 'val', or 'ser'.
            axis_id (int, optional): Unique numeric ID. Auto-generated if None.
            **kwargs: Any ChartAxis property (orientation, position, deleted, etc.)

        Returns:
            ChartAxis: The created axis.
        """
        ax = ChartAxis()
        ax.axis_type = axis_type
        if axis_id is not None:
            ax.axis_id = int(axis_id)
        else:
            existing_ids = {a.axis_id for a in self._axes}
            candidate = 70000000
            while candidate in existing_ids:
                candidate += 1
            ax.axis_id = candidate
        for k, v in kwargs.items():
            setattr(ax, k, v)
        self._axes.append(ax)
        self._mark_dirty()
        return ax

    def copy(self):
        new_chart = Chart(
            self._type,
            self._upper_left_row,
            self._upper_left_column,
            self._lower_right_row,
            self._lower_right_column,
        )
        new_chart._title = self._title
        new_chart._category_data = self._category_data
        new_chart._show_legend = self._show_legend
        new_chart._legend_position = self._legend_position
        new_chart._grouping = self._grouping
        new_chart._smooth = self._smooth
        new_chart._bar_direction = self._bar_direction
        new_chart._gap_width = self._gap_width
        new_chart._overlap = self._overlap
        new_chart._vary_colors = self._vary_colors
        new_chart._first_slice_angle = self._first_slice_angle
        new_chart._is_of_pie = self._is_of_pie
        new_chart._of_pie_type = self._of_pie_type
        new_chart._second_pie_size = self._second_pie_size
        new_chart._quartile_method = self._quartile_method
        new_chart._box_show_mean_line = self._box_show_mean_line
        new_chart._box_show_mean_marker = self._box_show_mean_marker
        new_chart._box_show_inner_points = self._box_show_inner_points
        new_chart._box_show_outlier_points = self._box_show_outlier_points
        new_chart._box_gap_width = self._box_gap_width
        new_chart._is_3d = self._is_3d
        new_chart._gap_depth = self._gap_depth
        new_chart._upper_left_row_offset = self._upper_left_row_offset
        new_chart._upper_left_column_offset = self._upper_left_column_offset
        new_chart._lower_right_row_offset = self._lower_right_row_offset
        new_chart._lower_right_column_offset = self._lower_right_column_offset
        new_chart._anchor_extent_width = self._anchor_extent_width
        new_chart._anchor_extent_height = self._anchor_extent_height
        new_chart._source_chart_xml = self._source_chart_xml
        new_chart._source_chart_rels_xml = self._source_chart_rels_xml
        new_chart._source_chart_extra_parts = list(self._source_chart_extra_parts)
        new_chart._source_chart_part_path = self._source_chart_part_path
        new_chart._source_chart_part_index = self._source_chart_part_index
        new_chart._source_chart_content_type = self._source_chart_content_type
        new_chart._source_is_chart_ex = self._source_is_chart_ex
        new_chart._show_connector_lines = self._show_connector_lines
        new_chart._has_subtotals = self._has_subtotals
        new_chart._view_3d = self._view_3d.copy(new_chart)
        new_chart._n_series = self._n_series.copy(new_chart)
        new_chart._sub_charts = [dict(sc) for sc in self._sub_charts]
        new_chart._axes = [ax.copy() for ax in self._axes]
        new_chart._scatter_style = self._scatter_style
        new_chart._is_bubble = self._is_bubble
        new_chart._bubble_scale = self._bubble_scale
        new_chart._disp_blanks_as = self._disp_blanks_as
        new_chart._auto_title_deleted = self._auto_title_deleted
        new_chart._stock_style = self._stock_style
        new_chart._wireframe = self._wireframe
        new_chart._radar_style = self._radar_style
        new_chart._histogram_bin_count = self._histogram_bin_count
        new_chart._histogram_bin_size = self._histogram_bin_size
        new_chart._histogram_interval_closed = self._histogram_interval_closed
        new_chart._histogram_overflow = self._histogram_overflow
        new_chart._histogram_underflow = self._histogram_underflow
        return new_chart


class ChartCollection:
    """Collection of charts in a worksheet."""

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._charts = []

    def add(self, chart_type, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """
        Adds a chart to the worksheet.

        Returns:
            int: Index of the created chart.
        """
        if isinstance(chart_type, int):
            chart_type = ChartType(chart_type)
        supported = (ChartType.LINE, ChartType.BAR, ChartType.PIE, ChartType.AREA,
                     ChartType.BOX_WHISKER, ChartType.WATERFALL, ChartType.COMBO, ChartType.SCATTER,
                     ChartType.STOCK, ChartType.SURFACE, ChartType.RADAR, ChartType.TREEMAP,
                     ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP)
        if chart_type not in supported:
            raise NotImplementedError("Unsupported chart type for creation")

        ul_row = int(upper_left_row)
        ul_col = int(upper_left_column)
        lr_row = int(lower_right_row)
        lr_col = int(lower_right_column)

        # Normalize anchor bounds so generated drawing anchors are valid even
        # when callers pass coordinates in reverse order.
        top_row = min(ul_row, lr_row)
        left_col = min(ul_col, lr_col)
        bottom_row = max(ul_row, lr_row)
        right_col = max(ul_col, lr_col)

        chart = Chart(chart_type, top_row, left_col, bottom_row, right_col)
        self._charts.append(chart)
        return len(self._charts) - 1

    def add_line(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a line chart and returns it."""
        idx = self.add(ChartType.LINE, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_bar(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a bar chart and returns it."""
        idx = self.add(ChartType.BAR, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_pie(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a pie chart and returns it."""
        idx = self.add(ChartType.PIE, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_area(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds an area chart and returns it."""
        idx = self.add(ChartType.AREA, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_box_whisker(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a box-whisker chart and returns it."""
        idx = self.add(ChartType.BOX_WHISKER, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._box_gap_width = 1
        return chart

    def add_waterfall(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a waterfall chart and returns it."""
        idx = self.add(ChartType.WATERFALL, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx waterfall default: legend at top
        return chart

    def add_scatter(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a scatter chart and returns it."""
        idx = self.add(ChartType.SCATTER, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_combo(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a combo chart and returns it."""
        idx = self.add(ChartType.COMBO, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_stock(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a stock chart and returns it."""
        idx = self.add(ChartType.STOCK, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        return self._charts[idx]

    def add_treemap(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a treemap chart and returns it."""
        idx = self.add(ChartType.TREEMAP, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx treemap default: legend at top
        return chart

    def add_sunburst(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a sunburst chart and returns it."""
        idx = self.add(ChartType.SUNBURST, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx sunburst default: legend at top
        return chart

    def add_histogram(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a histogram chart and returns it."""
        idx = self.add(ChartType.HISTOGRAM, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx histogram default: legend at top
        return chart

    def add_funnel(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a funnel chart and returns it."""
        idx = self.add(ChartType.FUNNEL, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx funnel default: legend at top
        return chart

    def add_map(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """Adds a map (region map) chart and returns it."""
        idx = self.add(ChartType.MAP, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._legend_position = "t"  # chartEx map default: legend at top
        return chart

    def add_radar(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column, radar_style="marker"):
        """Adds a radar chart and returns it. radar_style: 'standard', 'marker', or 'filled'."""
        idx = self.add(ChartType.RADAR, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._radar_style = radar_style
        return chart

    def add_surface(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column, is_3d=True, wireframe=False):
        """Adds a surface chart and returns it. is_3d=True for surface3DChart, False for surfaceChart."""
        idx = self.add(ChartType.SURFACE, upper_left_row, upper_left_column, lower_right_row, lower_right_column)
        chart = self._charts[idx]
        chart._is_3d = is_3d
        chart._wireframe = wireframe
        return chart

    def Add(self, chart_type, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add()."""
        return self.add(chart_type, upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddLine(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_line()."""
        return self.add_line(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddBar(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_bar()."""
        return self.add_bar(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddPie(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_pie()."""
        return self.add_pie(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddArea(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_area()."""
        return self.add_area(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddBoxWhisker(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_box_whisker()."""
        return self.add_box_whisker(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddWaterfall(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_waterfall()."""
        return self.add_waterfall(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddScatter(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_scatter()."""
        return self.add_scatter(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddCombo(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_combo()."""
        return self.add_combo(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddStock(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_stock()."""
        return self.add_stock(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddSunburst(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_sunburst()."""
        return self.add_sunburst(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddHistogram(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_histogram()."""
        return self.add_histogram(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddFunnel(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_funnel()."""
        return self.add_funnel(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def AddMap(self, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add_map()."""
        return self.add_map(upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    @property
    def count(self):
        return len(self._charts)

    def __len__(self):
        return len(self._charts)

    def __iter__(self):
        return iter(self._charts)

    def __getitem__(self, index):
        return self._charts[index]

    def copy(self, worksheet):
        new_collection = ChartCollection(worksheet)
        new_collection._charts = [chart.copy() for chart in self._charts]
        return new_collection
