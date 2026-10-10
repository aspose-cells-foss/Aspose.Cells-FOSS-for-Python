"""
Aspose.Cells for Python - XML Conditional Formatting Loader Module

This module provides the ConditionalFormatXMLLoader class which handles loading
conditional formatting data from XML format according to ECMA-376 specification.

ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element structure.
"""


_X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
_XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main'


class ConditionalFormatXMLLoader:
    """
    Handles loading conditional formatting data from XML format for .xlsx files.

    The ConditionalFormatXMLLoader class is responsible for parsing the XML
    representation of conditional formatting rules including cell value rules,
    text rules, date rules, formula rules, color scales, data bars, icon sets,
    and more.

    Examples:
        >>> loader = ConditionalFormatXMLLoader(namespaces, workbook)
        >>> loader.load_conditional_formatting(worksheet, worksheet_root)
    """

    def __init__(self, namespaces, workbook):
        """
        Initializes a new instance of the ConditionalFormatXMLLoader class.

        Args:
            namespaces: XML namespaces dictionary for parsing.
            workbook: The workbook instance for accessing dxf styles.
        """
        self.ns = namespaces
        self.workbook = workbook

    def load_conditional_formatting(self, worksheet, worksheet_root):
        """
        Loads conditional formatting from worksheet XML according to ECMA-376 specification.

        ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element and its children:
        - conditionalFormatting: Main element with sqref attribute
        - cfRule: Individual conditional formatting rule

        Args:
            worksheet (Worksheet): The worksheet object to load data into.
            worksheet_root: The XML root element of the worksheet.
        """
        from .conditional_format import ConditionalFormat

        # Find all conditionalFormatting elements
        cf_elements = worksheet_root.findall('.//main:conditionalFormatting', namespaces=self.ns)

        for cf_elem in cf_elements:
            sqref = cf_elem.get('sqref')
            if not sqref:
                continue

            # Load each cfRule within this conditionalFormatting element
            for rule_elem in cf_elem.findall('main:cfRule', namespaces=self.ns):
                cf = ConditionalFormat()
                cf._range = sqref

                # Parse rule attributes
                rule_type = rule_elem.get('type')
                cf._type = rule_type
                extension_id = rule_elem.find(f'.//{{{_X14_NS}}}id')
                if extension_id is not None and extension_id.text:
                    cf._data_bar_extension_id = extension_id.text.strip()

                # Parse priority
                priority = rule_elem.get('priority')
                if priority:
                    cf._priority = int(priority)

                # Parse stopIfTrue
                stop_if_true = rule_elem.get('stopIfTrue')
                cf._stop_if_true = stop_if_true == '1'

                # Parse dxfId for loading formatting later
                dxf_id = rule_elem.get('dxfId')
                if dxf_id:
                    cf._dxf_id = int(dxf_id)

                # Parse operator for cellIs type
                operator = rule_elem.get('operator')
                if operator:
                    cf._operator = operator

                # Parse text attribute for text-based rules
                text = rule_elem.get('text')
                if text:
                    cf._formula1 = text

                # Parse timePeriod attribute
                time_period = rule_elem.get('timePeriod')
                if time_period:
                    cf._operator = time_period

                # Parse top10 attributes
                if rule_type == 'top10':
                    bottom = rule_elem.get('bottom')
                    cf._top = bottom != '1'
                    percent = rule_elem.get('percent')
                    cf._percent = percent == '1'
                    rank = rule_elem.get('rank')
                    if rank:
                        cf._rank = int(rank)

                # Parse aboveAverage attributes
                if rule_type == 'aboveAverage':
                    above_average = rule_elem.get('aboveAverage', '1')
                    cf._above = above_average != '0'
                    std_dev = rule_elem.get('stdDev')
                    if std_dev:
                        cf._std_dev = int(std_dev)

                # Parse formula elements based on rule type
                formula_elems = rule_elem.findall('main:formula', namespaces=self.ns)
                if rule_type == 'expression':
                    # Expression rules store formula in _formula property
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula = formula_elems[0].text
                elif rule_type == 'cellIs':
                    # Cell value rules use _formula1 and _formula2
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula1 = formula_elems[0].text
                    if len(formula_elems) > 1 and formula_elems[1].text:
                        cf._formula2 = formula_elems[1].text
                elif rule_type in ('containsText', 'notContainsText', 'beginsWith', 'endsWith'):
                    # Text rules: the text attribute is already parsed above
                    # The formula element contains the Excel formula, not the text value
                    # We use the text attribute value which was set earlier
                    pass
                else:
                    # Default: store in _formula1
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula1 = formula_elems[0].text
                    if len(formula_elems) > 1 and formula_elems[1].text:
                        cf._formula2 = formula_elems[1].text

                # Parse colorScale element
                color_scale_elem = rule_elem.find('main:colorScale', namespaces=self.ns)
                if color_scale_elem is not None:
                    self._load_color_scale(cf, color_scale_elem)

                # Parse dataBar element
                data_bar_elem = rule_elem.find('main:dataBar', namespaces=self.ns)
                if data_bar_elem is not None:
                    self._load_data_bar(cf, data_bar_elem)

                # Parse iconSet element
                icon_set_elem = rule_elem.find('main:iconSet', namespaces=self.ns)
                if icon_set_elem is not None:
                    self._load_icon_set(cf, icon_set_elem)

                # Add to worksheet's conditional formats
                worksheet.conditional_formats._formats.append(cf)

        self._apply_extended_data_bars(worksheet, worksheet_root)

        # Load dxf formatting and apply to conditional formats
        self._apply_dxf_to_conditional_formats(worksheet)

    def _load_color_scale(self, cf, color_scale_elem):
        """Loads colorScale element data into conditional format."""
        # Get cfvo elements to determine if 2-color or 3-color
        cfvo_elems = color_scale_elem.findall('main:cfvo', namespaces=self.ns)
        cf._color_scale_type = '3-color' if len(cfvo_elems) >= 3 else '2-color'
        cf._color_scale_thresholds = [
            dict(cfvo_elem.attrib) for cfvo_elem in cfvo_elems
        ]

        # Get color elements
        color_elems = color_scale_elem.findall('main:color', namespaces=self.ns)
        if len(color_elems) >= 1:
            cf._min_color = color_elems[0].get('rgb')
        if len(color_elems) >= 3:
            cf._mid_color = color_elems[1].get('rgb')
            cf._max_color = color_elems[2].get('rgb')
        elif len(color_elems) >= 2:
            cf._max_color = color_elems[1].get('rgb')

    def _load_data_bar(self, cf, data_bar_elem):
        """Loads dataBar element data into conditional format."""
        cf._data_bar_thresholds = [
            dict(cfvo_elem.attrib)
            for cfvo_elem in data_bar_elem.findall(
                'main:cfvo', namespaces=self.ns
            )
        ]
        try:
            cf._data_bar_min_length = int(
                data_bar_elem.get('minLength', '10')
            )
        except (TypeError, ValueError):
            cf._data_bar_min_length = 10
        try:
            cf._data_bar_max_length = int(
                data_bar_elem.get('maxLength', '90')
            )
        except (TypeError, ValueError):
            cf._data_bar_max_length = 90
        cf._data_bar_show_value = data_bar_elem.get(
            'showValue', '1'
        ) not in ('0', 'false', 'False')
        # The legacy dataBar schema used by this loader renders a gradient.
        # Solid fills and other extended properties live in the x14 extension.
        cf._data_bar_gradient = True

        # Get color element
        color_elem = data_bar_elem.find('main:color', namespaces=self.ns)
        if color_elem is not None:
            cf._bar_color = color_elem.get('rgb')

    def _apply_extended_data_bars(self, worksheet, worksheet_root):
        """Merges x14 data-bar properties into their legacy cfRule."""
        data_bars = [
            rule for rule in worksheet.conditional_formats
            if getattr(rule, '_type', None) == 'dataBar'
        ]
        by_id = {
            self._normalize_extension_id(
                getattr(rule, '_data_bar_extension_id', None)
            ): rule
            for rule in data_bars
            if getattr(rule, '_data_bar_extension_id', None)
        }

        for container in worksheet_root.findall(
            f'.//{{{_X14_NS}}}conditionalFormatting'
        ):
            sqref_elem = container.find(f'{{{_XM_NS}}}sqref')
            sqref = (
                sqref_elem.text.strip()
                if sqref_elem is not None and sqref_elem.text
                else None
            )
            for extended_rule in container.findall(
                f'{{{_X14_NS}}}cfRule'
            ):
                if extended_rule.get('type') != 'dataBar':
                    continue
                extension_id = self._normalize_extension_id(
                    extended_rule.get('id')
                )
                target = by_id.get(extension_id)
                if target is None and sqref is not None:
                    target = next((
                        rule for rule in data_bars
                        if str(getattr(rule, 'range', '')).strip() == sqref
                    ), None)
                if target is None:
                    continue
                data_bar = extended_rule.find(f'{{{_X14_NS}}}dataBar')
                if data_bar is not None:
                    self._load_extended_data_bar(target, data_bar)

    def _load_extended_data_bar(self, cf, data_bar_elem):
        cf._data_bar_min_length = self._integer_attribute(
            data_bar_elem, 'minLength', cf._data_bar_min_length
        )
        cf._data_bar_max_length = self._integer_attribute(
            data_bar_elem, 'maxLength', cf._data_bar_max_length
        )
        cf._data_bar_show_value = self._boolean_attribute(
            data_bar_elem, 'showValue', cf._data_bar_show_value
        )
        cf._data_bar_gradient = self._boolean_attribute(
            data_bar_elem, 'gradient', cf._data_bar_gradient
        )
        cf._show_border = self._boolean_attribute(
            data_bar_elem, 'border', cf._show_border
        )

        direction = data_bar_elem.get('direction')
        if direction == 'leftToRight':
            cf._direction = 'left-to-right'
        elif direction == 'rightToLeft':
            cf._direction = 'right-to-left'
        cf._data_bar_axis_position = data_bar_elem.get('axisPosition')

        thresholds = []
        for threshold_elem in data_bar_elem.findall(
            f'{{{_X14_NS}}}cfvo'
        ):
            threshold = {'type': threshold_elem.get('type', '')}
            formula = threshold_elem.find(f'{{{_XM_NS}}}f')
            if formula is not None and formula.text is not None:
                threshold['val'] = formula.text
            thresholds.append(threshold)
        if len(thresholds) == 2:
            cf._data_bar_thresholds = thresholds

        color_properties = (
            ('negativeFillColor', '_negative_color'),
            ('borderColor', '_data_bar_border_color'),
            ('negativeBorderColor', '_data_bar_negative_border_color'),
            ('axisColor', '_data_bar_axis_color'),
        )
        for element_name, property_name in color_properties:
            color = data_bar_elem.find(f'{{{_X14_NS}}}{element_name}')
            if color is not None and color.get('rgb'):
                setattr(cf, property_name, color.get('rgb'))

    @staticmethod
    def _normalize_extension_id(value):
        return str(value or '').strip().upper()

    @staticmethod
    def _integer_attribute(element, name, default):
        try:
            return int(element.get(name, default))
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _boolean_attribute(element, name, default):
        value = element.get(name)
        if value is None:
            return default
        return value not in ('0', 'false', 'False')

    def _load_icon_set(self, cf, icon_set_elem):
        """Loads iconSet element data into conditional format."""
        cf._icon_set_type = icon_set_elem.get('iconSet', '3TrafficLights1')
        cf._reverse_icons = icon_set_elem.get('reverse') == '1'
        cf._show_icon_only = icon_set_elem.get('showValue') == '0'
        cf._icon_set_thresholds = [
            dict(threshold.attrib)
            for threshold in icon_set_elem.findall(
                'main:cfvo', namespaces=self.ns
            )
        ]

    def _apply_dxf_to_conditional_formats(self, worksheet):
        """Applies differential formatting (dxf) to conditional formats."""
        # dxf styles are stored in workbook._dxf_styles after loading styles.xml
        if not hasattr(self.workbook, '_dxf_styles') or not self.workbook._dxf_styles:
            return

        for cf in worksheet.conditional_formats:
            if hasattr(cf, '_dxf_id') and cf._dxf_id is not None:
                if cf._dxf_id < len(self.workbook._dxf_styles):
                    dxf_data = self.workbook._dxf_styles[cf._dxf_id]
                    self._apply_dxf_data_to_cf(cf, dxf_data)

    def _apply_dxf_data_to_cf(self, cf, dxf_data):
        """Applies dxf data to a conditional format."""
        if 'font' in dxf_data:
            font = dxf_data['font']
            cf._font.bold = font.get('bold', False)
            cf._font.italic = font.get('italic', False)
            cf._font.underline = font.get('underline', False)
            cf._font.strikethrough = font.get('strikethrough', False)
            if 'color' in font:
                cf._font.color = font['color']

        if 'fill' in dxf_data:
            fill = dxf_data['fill']
            foreground_color = fill.get(
                'fg_color', fill.get('bg_color', 'FFFFFFFF')
            )
            background_color = fill.get('bg_color', foreground_color)
            cf._fill.pattern_type = fill.get('pattern_type', 'solid')
            cf._fill.foreground_color = foreground_color
            cf._fill.background_color = background_color
