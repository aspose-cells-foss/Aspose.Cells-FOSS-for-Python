"""Style resolution helpers for layout and rendering."""

import math
import re
from datetime import date, datetime, timedelta

from .style import Style, NumberFormat


_UNSUPPORTED_CONDITIONAL_EXPRESSION = object()
_CONDITIONAL_CELL_REFERENCE = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<absolute_column>\$?)(?P<column>[A-Za-z]{1,3})"
    r"(?P<absolute_row>\$?)(?P<row>[1-9][0-9]*)$"
)
_CONDITIONAL_NUMBER = re.compile(
    r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$"
)
_SUPPORTED_PDF_ICON_SETS = frozenset(('3TrafficLights1',))


class StyleResolver:
    """
    Resolves the effective visual style for worksheet cells.

    Precedence:
    1. workbook default style
    2. table style
    3. column style xf
    4. cell style
    5. matching conditional-format differential styles, when requested
    """

    @classmethod
    def resolve_cell_style(
        cls,
        worksheet,
        row,
        column,
        include_conditional_formats=False,
    ):
        if worksheet is None:
            raise ValueError("worksheet is required")
        if row is None or int(row) < 1:
            raise ValueError("row must be >= 1")
        if column is None:
            raise ValueError("column is required")

        row = int(row)
        column = cls._normalize_column(column)
        workbook = getattr(worksheet, "_workbook", None)

        resolved = cls._copy_style(cls._get_workbook_default_style(workbook))
        has_table_style = cls._overlay_table_style(
            resolved, worksheet, row, column
        )

        column_style_idx = (getattr(worksheet, "_column_styles", None) or {}).get(column)
        if column_style_idx is not None:
            cls._overlay_style(
                resolved,
                cls.style_from_index(workbook, column_style_idx),
                full=not has_table_style,
                default_style=cls._blank_style(),
            )

        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        if cell is not None:
            if cls._has_explicit_style_index(cell):
                cls._overlay_style(
                    resolved,
                    cell.style,
                    full=not has_table_style,
                    default_style=cls._blank_style(),
                )
            else:
                cls._overlay_style(
                    resolved,
                    cell.style,
                    full=False,
                    default_style=cls._blank_style(),
                )

        if include_conditional_formats:
            cls._overlay_conditional_formats(
                resolved, worksheet, row, column
            )

        return resolved

    @classmethod
    def resolve_data_bar(cls, worksheet, row, column):
        """Resolve the highest-priority supported data bar for a cell."""
        if worksheet is None:
            raise ValueError("worksheet is required")
        row = int(row)
        column = cls._normalize_column(column)

        indexed_rules = list(enumerate(worksheet.conditional_formats))
        indexed_rules.sort(key=lambda item: (
            int(getattr(item[1], "priority", 0) or 0), item[0]
        ))
        for _, rule in indexed_rules:
            areas = cls._conditional_areas(worksheet, rule.range)
            if not any(
                min_row <= row <= max_row
                and min_col <= column <= max_col
                for min_row, min_col, max_row, max_col in areas
            ):
                continue
            if not cls._conditional_rule_matches(
                worksheet, rule, row, column, areas
            ):
                continue
            if getattr(rule, '_type', None) == 'dataBar':
                return cls._data_bar_resolution(
                    worksheet, rule, row, column, areas
                )
            if rule.stop_if_true:
                return None
        return None

    @classmethod
    def resolve_icon_set(cls, worksheet, row, column):
        """Resolve the highest-priority supported icon for a cell."""
        if worksheet is None:
            raise ValueError("worksheet is required")
        row = int(row)
        column = cls._normalize_column(column)

        indexed_rules = list(enumerate(worksheet.conditional_formats))
        indexed_rules.sort(key=lambda item: (
            int(getattr(item[1], "priority", 0) or 0), item[0]
        ))
        for _, rule in indexed_rules:
            areas = cls._conditional_areas(worksheet, rule.range)
            if not any(
                min_row <= row <= max_row
                and min_col <= column <= max_col
                for min_row, min_col, max_row, max_col in areas
            ):
                continue
            if not cls._conditional_rule_matches(
                worksheet, rule, row, column, areas
            ):
                continue
            if getattr(rule, '_type', None) == 'iconSet':
                return cls._icon_set_resolution(
                    worksheet, rule, row, column, areas
                )
            if rule.stop_if_true:
                return None
        return None

    @classmethod
    def _overlay_conditional_formats(
        cls, target, worksheet, row, column
    ):
        indexed_rules = list(enumerate(worksheet.conditional_formats))
        indexed_rules.sort(key=lambda item: (
            int(getattr(item[1], "priority", 0) or 0), item[0]
        ))

        matching_rules = []
        for _, rule in indexed_rules:
            areas = cls._conditional_areas(worksheet, rule.range)
            if not any(
                min_row <= row <= max_row
                and min_col <= column <= max_col
                for min_row, min_col, max_row, max_col in areas
            ):
                continue
            if not cls._conditional_rule_matches(
                worksheet, rule, row, column, areas
            ):
                continue
            matching_rules.append((rule, areas))
            if rule.stop_if_true:
                break

        # Apply low-priority rules first so higher-priority properties win.
        for rule, areas in reversed(matching_rules):
            rule_type = getattr(rule, '_type', None)
            if rule_type == 'colorScale':
                cls._overlay_color_scale(
                    target, worksheet, rule, row, column, areas
                )
            elif rule_type not in ('dataBar', 'iconSet'):
                cls._overlay_conditional_dxf(target, worksheet, rule)

    @staticmethod
    def _conditional_areas(worksheet, range_text):
        if not range_text:
            return ()

        areas = []
        for token in str(range_text).replace(',', ' ').split():
            token = token.rsplit('!', 1)[-1].replace('$', '')
            start_ref, separator, end_ref = token.partition(':')
            if not separator:
                end_ref = start_ref
            try:
                start_row, start_col = (
                    worksheet.cells.coordinate_from_string(start_ref)
                )
                end_row, end_col = (
                    worksheet.cells.coordinate_from_string(end_ref)
                )
            except ValueError:
                continue
            areas.append((
                min(start_row, end_row),
                min(start_col, end_col),
                max(start_row, end_row),
                max(start_col, end_col),
            ))
        return tuple(areas)

    @classmethod
    def _conditional_rule_matches(
        cls, worksheet, rule, row, column, areas
    ):
        rule_type = getattr(rule, '_type', None)
        if rule_type in ('cellIs', 'cellValue'):
            return cls._cell_value_rule_matches(
                worksheet, rule, row, column
            )
        if rule_type in (
            'text', 'containsText', 'notContainsText',
            'beginsWith', 'endsWith',
        ):
            return cls._text_rule_matches(worksheet, rule, row, column)
        if rule_type in ('duplicateValues', 'uniqueValues'):
            return cls._duplicate_rule_matches(
                worksheet, rule, row, column, areas
            )
        if rule_type in ('top10', 'bottom10'):
            return cls._top_bottom_rule_matches(
                worksheet, rule, row, column, areas
            )
        if rule_type in ('date', 'timePeriod'):
            return cls._time_period_rule_matches(
                worksheet, rule, row, column
            )
        if rule_type in ('expression', 'formula'):
            return cls._expression_rule_matches(
                worksheet, rule, row, column, areas
            )
        if rule_type in ('colorScale', 'dataBar', 'iconSet'):
            ref = worksheet.cells.coordinate_to_string(row, column)
            cell = worksheet.cells.get_all_cells().get(ref)
            return cls._conditional_numeric_value(
                getattr(cell, 'value', None)
            ) is not None
        if rule_type not in ('aboveAverage', 'belowAverage'):
            return False

        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        if value is None:
            return False

        values = cls._conditional_numeric_values(worksheet, areas)
        if not values:
            return False

        average = sum(values) / len(values)
        above = getattr(rule, 'above', None)
        if above is None:
            above = rule_type != 'belowAverage'
        standard_deviations = abs(int(
            getattr(rule, 'std_dev', 0) or 0
        ))
        if standard_deviations:
            variance = sum(
                (candidate - average) ** 2 for candidate in values
            ) / len(values)
            distance = math.sqrt(variance) * standard_deviations
            threshold = average + distance if above else average - distance
        else:
            threshold = average
        return value > threshold if above else value < threshold

    @classmethod
    def _expression_rule_matches(
        cls, worksheet, rule, row, column, areas
    ):
        formula = getattr(rule, 'formula', None)
        if not isinstance(formula, str) or not formula.strip():
            return False

        anchor = next((
            (min_row, min_col)
            for min_row, min_col, max_row, max_col in areas
            if min_row <= row <= max_row and min_col <= column <= max_col
        ), None)
        if anchor is None:
            return False

        expression = formula.strip()
        if expression.startswith('='):
            expression = expression[1:].strip()
        result = cls._evaluate_conditional_expression(
            worksheet,
            expression,
            row,
            column,
            anchor[0],
            anchor[1],
        )
        if result is _UNSUPPORTED_CONDITIONAL_EXPRESSION:
            return False
        return cls._conditional_truthy(result)

    @classmethod
    def _evaluate_conditional_expression(
        cls, worksheet, expression, row, column, anchor_row, anchor_column
    ):
        expression = cls._strip_conditional_parentheses(expression.strip())
        if not expression:
            return _UNSUPPORTED_CONDITIONAL_EXPRESSION

        function = cls._conditional_function_call(expression)
        if function is not None:
            name, argument_text = function
            arguments = cls._split_conditional_arguments(argument_text)
            if arguments is None:
                return _UNSUPPORTED_CONDITIONAL_EXPRESSION
            values = [
                cls._evaluate_conditional_expression(
                    worksheet,
                    argument,
                    row,
                    column,
                    anchor_row,
                    anchor_column,
                )
                for argument in arguments
            ]
            if any(
                value is _UNSUPPORTED_CONDITIONAL_EXPRESSION
                for value in values
            ):
                return _UNSUPPORTED_CONDITIONAL_EXPRESSION
            if name == 'AND' and values:
                return all(cls._conditional_truthy(value) for value in values)
            if name == 'OR' and values:
                return any(cls._conditional_truthy(value) for value in values)
            if name == 'NOT' and len(values) == 1:
                return not cls._conditional_truthy(values[0])
            return _UNSUPPORTED_CONDITIONAL_EXPRESSION

        comparison = cls._find_conditional_comparison(expression)
        if comparison is not None:
            index, operator = comparison
            left = cls._evaluate_conditional_operand(
                worksheet,
                expression[:index],
                row,
                column,
                anchor_row,
                anchor_column,
            )
            right = cls._evaluate_conditional_operand(
                worksheet,
                expression[index + len(operator):],
                row,
                column,
                anchor_row,
                anchor_column,
            )
            if (
                left is _UNSUPPORTED_CONDITIONAL_EXPRESSION
                or right is _UNSUPPORTED_CONDITIONAL_EXPRESSION
            ):
                return _UNSUPPORTED_CONDITIONAL_EXPRESSION
            return cls._compare_conditional_values(left, right, operator)

        return cls._evaluate_conditional_operand(
            worksheet,
            expression,
            row,
            column,
            anchor_row,
            anchor_column,
        )

    @classmethod
    def _evaluate_conditional_operand(
        cls, worksheet, operand, row, column, anchor_row, anchor_column
    ):
        operand = cls._strip_conditional_parentheses(operand.strip())
        if not operand:
            return _UNSUPPORTED_CONDITIONAL_EXPRESSION
        if cls._find_conditional_comparison(operand):
            return cls._evaluate_conditional_expression(
                worksheet,
                operand,
                row,
                column,
                anchor_row,
                anchor_column,
            )

        function = cls._conditional_function_call(operand)
        if function is not None:
            return cls._evaluate_conditional_expression(
                worksheet,
                operand,
                row,
                column,
                anchor_row,
                anchor_column,
            )

        upper = operand.upper()
        if upper == 'TRUE':
            return True
        if upper == 'FALSE':
            return False
        if len(operand) >= 2 and operand[0] == operand[-1] == '"':
            return operand[1:-1].replace('""', '"')
        if _CONDITIONAL_NUMBER.fullmatch(operand):
            try:
                return float(operand)
            except (ValueError, OverflowError):
                return _UNSUPPORTED_CONDITIONAL_EXPRESSION

        reference = _CONDITIONAL_CELL_REFERENCE.fullmatch(operand)
        if reference is None:
            return _UNSUPPORTED_CONDITIONAL_EXPRESSION
        sheet_name = reference.group('sheet')
        if sheet_name is not None:
            if sheet_name.startswith("'"):
                sheet_name = sheet_name[1:-1].replace("''", "'")
            if sheet_name.casefold() != str(
                getattr(worksheet, 'name', '')
            ).casefold():
                return _UNSUPPORTED_CONDITIONAL_EXPRESSION

        source_column = worksheet.cells.column_index_from_string(
            reference.group('column')
        )
        source_row = int(reference.group('row'))
        if not reference.group('absolute_column'):
            source_column += column - anchor_column
        if not reference.group('absolute_row'):
            source_row += row - anchor_row
        if source_row < 1 or source_column < 1:
            return _UNSUPPORTED_CONDITIONAL_EXPRESSION

        source_ref = worksheet.cells.coordinate_to_string(
            source_row, source_column
        )
        cell = worksheet.cells.get_all_cells().get(source_ref)
        return getattr(cell, 'value', None)

    @staticmethod
    def _conditional_function_call(expression):
        open_index = expression.find('(')
        if open_index <= 0 or not expression.endswith(')'):
            return None
        name = expression[:open_index].strip().upper()
        if not re.fullmatch(r'[A-Z][A-Z0-9_.]*', name):
            return None
        if not StyleResolver._conditional_parentheses_wrap(
            expression[open_index:]
        ):
            return None
        return name, expression[open_index + 1:-1]

    @staticmethod
    def _split_conditional_arguments(argument_text):
        if not argument_text.strip():
            return []
        arguments = []
        start = 0
        depth = 0
        quoted = False
        index = 0
        while index < len(argument_text):
            character = argument_text[index]
            if character == '"':
                if quoted and index + 1 < len(argument_text) and (
                    argument_text[index + 1] == '"'
                ):
                    index += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if character == '(':
                    depth += 1
                elif character == ')':
                    depth -= 1
                    if depth < 0:
                        return None
                elif character in ',;' and depth == 0:
                    argument = argument_text[start:index].strip()
                    if not argument:
                        return None
                    arguments.append(argument)
                    start = index + 1
            index += 1
        if quoted or depth != 0:
            return None
        argument = argument_text[start:].strip()
        if not argument:
            return None
        arguments.append(argument)
        return arguments

    @staticmethod
    def _find_conditional_comparison(expression):
        depth = 0
        quoted = False
        index = 0
        while index < len(expression):
            character = expression[index]
            if character == '"':
                if quoted and index + 1 < len(expression) and (
                    expression[index + 1] == '"'
                ):
                    index += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if character == '(':
                    depth += 1
                elif character == ')':
                    depth -= 1
                    if depth < 0:
                        return None
                elif depth == 0:
                    for operator in ('>=', '<=', '<>', '=', '>', '<'):
                        if expression.startswith(operator, index):
                            return index, operator
            index += 1
        return None

    @classmethod
    def _strip_conditional_parentheses(cls, expression):
        while cls._conditional_parentheses_wrap(expression):
            expression = expression[1:-1].strip()
        return expression

    @staticmethod
    def _conditional_parentheses_wrap(expression):
        if not expression.startswith('(') or not expression.endswith(')'):
            return False
        depth = 0
        quoted = False
        index = 0
        while index < len(expression):
            character = expression[index]
            if character == '"':
                if quoted and index + 1 < len(expression) and (
                    expression[index + 1] == '"'
                ):
                    index += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if character == '(':
                    depth += 1
                elif character == ')':
                    depth -= 1
                    if depth == 0 and index != len(expression) - 1:
                        return False
                    if depth < 0:
                        return False
            index += 1
        return not quoted and depth == 0

    @staticmethod
    def _compare_conditional_values(left, right, operator):
        if left is None:
            left = '' if isinstance(right, str) else 0.0
        if right is None:
            right = '' if isinstance(left, str) else 0.0

        if isinstance(left, str) and isinstance(right, str):
            left = left.casefold()
            right = right.casefold()
        elif (
            isinstance(left, (int, float))
            and not isinstance(left, bool)
            and isinstance(right, (int, float))
            and not isinstance(right, bool)
        ):
            left = float(left)
            right = float(right)
        elif type(left) is not type(right):
            return operator == '<>'

        try:
            return {
                '=': left == right,
                '<>': left != right,
                '>': left > right,
                '<': left < right,
                '>=': left >= right,
                '<=': left <= right,
            }[operator]
        except (KeyError, TypeError, ValueError):
            return False

    @staticmethod
    def _conditional_truthy(value):
        if value is None or value == '':
            return False
        if isinstance(value, str):
            return value.casefold() == 'true'
        return bool(value)

    @classmethod
    def _cell_value_rule_matches(cls, worksheet, rule, row, column):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        if value is None:
            return False

        first = cls._conditional_formula_number(
            getattr(rule, 'formula1', None)
        )
        operator = getattr(rule, 'operator', None)
        if first is None or operator is None:
            return False

        if operator == 'equal':
            return value == first
        if operator == 'notEqual':
            return value != first
        if operator == 'greaterThan':
            return value > first
        if operator == 'lessThan':
            return value < first
        if operator == 'greaterThanOrEqual':
            return value >= first
        if operator == 'lessThanOrEqual':
            return value <= first

        second = cls._conditional_formula_number(
            getattr(rule, 'formula2', None)
        )
        if second is None:
            return False
        if operator == 'between':
            return first <= value <= second
        if operator == 'notBetween':
            return value < first or value > second
        return False

    @classmethod
    def _text_rule_matches(cls, worksheet, rule, row, column):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = getattr(cell, 'value', None)
        text = '' if value is None else str(value)
        criterion = str(getattr(rule, 'text_formula', None) or '')
        text = text.casefold()
        criterion = criterion.casefold()

        rule_type = getattr(rule, '_type', None)
        operator = {
            'containsText': 'contains',
            'notContainsText': 'notContains',
            'beginsWith': 'beginsWith',
            'endsWith': 'endsWith',
        }.get(rule_type, getattr(rule, 'text_operator', None))
        if operator in ('contains', 'containsText'):
            return criterion in text
        if operator in ('notContains', 'notContainsText'):
            return criterion not in text
        if operator == 'beginsWith':
            return text.startswith(criterion)
        if operator == 'endsWith':
            return text.endswith(criterion)
        return False

    @classmethod
    def _time_period_rule_matches(cls, worksheet, rule, row, column):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_date_value(
            worksheet, getattr(cell, 'value', None)
        )
        if value is None:
            return False

        today = cls._today()
        operator = getattr(rule, 'date_operator', None)
        if operator == 'yesterday':
            return value == today - timedelta(days=1)
        if operator == 'today':
            return value == today
        if operator == 'tomorrow':
            return value == today + timedelta(days=1)
        if operator == 'last7Days':
            return today - timedelta(days=6) <= value <= today
        if operator == 'yearToDate':
            return date(today.year, 1, 1) <= value <= today

        if operator in ('lastWeek', 'thisWeek', 'nextWeek'):
            this_week = today - timedelta(
                days=(today.weekday() + 1) % 7
            )
            offset = {
                'lastWeek': -7,
                'thisWeek': 0,
                'nextWeek': 7,
            }[operator]
            start = this_week + timedelta(days=offset)
            end = start + timedelta(days=7)
            return start <= value < end

        if operator in ('lastMonth', 'thisMonth', 'nextMonth'):
            offset = {
                'lastMonth': -1,
                'thisMonth': 0,
                'nextMonth': 1,
            }[operator]
            start = cls._month_start(today, offset)
            end = cls._month_start(today, offset + 1)
            return start <= value < end

        if operator in ('lastQuarter', 'thisQuarter', 'nextQuarter'):
            offset = {
                'lastQuarter': -3,
                'thisQuarter': 0,
                'nextQuarter': 3,
            }[operator]
            quarter_month = ((today.month - 1) // 3) * 3 + 1
            quarter_start = date(today.year, quarter_month, 1)
            start = cls._month_start(quarter_start, offset)
            end = cls._month_start(quarter_start, offset + 3)
            return start <= value < end

        if operator in ('lastYear', 'thisYear', 'nextYear'):
            year = today.year + {
                'lastYear': -1,
                'thisYear': 0,
                'nextYear': 1,
            }[operator]
            return date(year, 1, 1) <= value < date(year + 1, 1, 1)
        return False

    @classmethod
    def _conditional_date_value(cls, worksheet, value):
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        number = cls._conditional_numeric_value(value)
        if number is None:
            return None

        workbook = getattr(worksheet, '_workbook', None)
        properties = getattr(workbook, 'properties', None)
        workbook_pr = getattr(properties, 'workbook_pr', None)
        use_1904_system = bool(getattr(workbook_pr, 'date1904', False))
        epoch = date(1904, 1, 1) if use_1904_system else date(1899, 12, 30)
        try:
            return epoch + timedelta(days=math.floor(number))
        except OverflowError:
            return None

    @staticmethod
    def _month_start(value, month_offset):
        month_index = value.year * 12 + value.month - 1 + month_offset
        year, month = divmod(month_index, 12)
        return date(year, month + 1, 1)

    @staticmethod
    def _today():
        return date.today()

    @classmethod
    def _duplicate_rule_matches(
        cls, worksheet, rule, row, column, areas
    ):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        key = cls._conditional_value_key(getattr(cell, 'value', None))
        if key is None:
            return False

        count = 0
        cells = worksheet.cells.get_all_cells()
        visited = set()
        for min_row, min_col, max_row, max_col in areas:
            for candidate_row in range(min_row, max_row + 1):
                for candidate_column in range(min_col, max_col + 1):
                    coordinate = (candidate_row, candidate_column)
                    if coordinate in visited:
                        continue
                    visited.add(coordinate)
                    candidate_ref = worksheet.cells.coordinate_to_string(
                        candidate_row, candidate_column
                    )
                    candidate = cells.get(candidate_ref)
                    if cls._conditional_value_key(
                        getattr(candidate, 'value', None)
                    ) == key:
                        count += 1

        duplicate = getattr(rule, '_type', None) == 'duplicateValues'
        return count > 1 if duplicate else count == 1

    @classmethod
    def _top_bottom_rule_matches(
        cls, worksheet, rule, row, column, areas
    ):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        values = cls._conditional_numeric_values(worksheet, areas)
        if value is None or not values:
            return False

        try:
            rank = max(1, int(getattr(rule, 'rank', None) or 10))
        except (TypeError, ValueError, OverflowError):
            rank = 10
        if getattr(rule, 'percent', False):
            rank = max(1, math.ceil(len(values) * min(rank, 100) / 100))
        rank = min(rank, len(values))

        top = getattr(rule, 'top', None)
        if top is None:
            top = getattr(rule, '_type', None) != 'bottom10'
        ordered = sorted(values, reverse=bool(top))
        threshold = ordered[rank - 1]
        return value >= threshold if top else value <= threshold

    @classmethod
    def _conditional_value_key(cls, value):
        if value is None or value == '':
            return None
        if isinstance(value, bool):
            return ('bool', value)
        if isinstance(value, str):
            return ('text', value.casefold())
        number = cls._conditional_numeric_value(value)
        if number is not None:
            return ('number', number)
        try:
            hash(value)
            return ('value', value)
        except TypeError:
            return ('value', repr(value))

    @staticmethod
    def _conditional_formula_number(formula):
        if formula is None or isinstance(formula, bool):
            return None
        if isinstance(formula, str):
            formula = formula.strip()
            if formula.startswith('='):
                formula = formula[1:].strip()
            if not formula:
                return None
        try:
            number = float(formula)
        except (TypeError, ValueError, OverflowError):
            return None
        return number if math.isfinite(number) else None

    @classmethod
    def _conditional_numeric_values(cls, worksheet, areas):
        cells = worksheet.cells.get_all_cells()
        values = []
        visited = set()
        for min_row, min_col, max_row, max_col in areas:
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    coordinate = (row, column)
                    if coordinate in visited:
                        continue
                    visited.add(coordinate)
                    ref = worksheet.cells.coordinate_to_string(row, column)
                    cell = cells.get(ref)
                    value = cls._conditional_numeric_value(
                        getattr(cell, 'value', None)
                    )
                    if value is not None:
                        values.append(value)
        return values

    @staticmethod
    def _conditional_numeric_value(value):
        if value is None or isinstance(value, (bool, str)):
            return None
        try:
            number = float(value)
        except (TypeError, ValueError, OverflowError):
            return None
        return number if math.isfinite(number) else None

    @classmethod
    def _overlay_color_scale(
        cls, target, worksheet, rule, row, column, areas
    ):
        color = cls._color_scale_color(
            worksheet, rule, row, column, areas
        )
        if color is None:
            return
        target.fill.set_solid_fill(color)
        target.fill._fg_color_type = 'rgb'
        target.fill._fg_color_value = color
        target.fill._bg_color_type = 'rgb'
        target.fill._bg_color_value = color

    @classmethod
    def _data_bar_resolution(
        cls, worksheet, rule, row, column, areas
    ):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        values = cls._conditional_numeric_values(worksheet, areas)
        if value is None or not values or min(values) < 0:
            # Negative ranges require an axis and separate positive/negative
            # geometry, which is intentionally not approximated here.
            return None

        defaults = [{'type': 'min'}, {'type': 'max'}]
        specs = list(
            getattr(rule, '_data_bar_thresholds', ()) or ()
        )
        if len(specs) != 2:
            specs = defaults
        thresholds = []
        for index, spec in enumerate(specs):
            threshold = cls._data_bar_threshold(spec, values)
            if threshold is None:
                threshold = cls._data_bar_threshold(
                    defaults[index], values
                )
            thresholds.append(threshold)

        lower, upper = thresholds
        if upper <= lower:
            normalized = 1.0
        else:
            normalized = min(max(
                (value - lower) / (upper - lower), 0.0
            ), 1.0)

        min_length = min(max(float(getattr(
            rule, '_data_bar_min_length', 10
        )), 0.0), 100.0)
        max_length = min(max(float(getattr(
            rule, '_data_bar_max_length', 90
        )), 0.0), 100.0)
        if max_length < min_length:
            min_length, max_length = max_length, min_length
        length_ratio = (
            min_length + normalized * (max_length - min_length)
        ) / 100.0

        color = cls._normalize_argb(
            getattr(rule, 'bar_color', None) or 'FF638EC6'
        )
        if color is None:
            return None
        return {
            'length_ratio': length_ratio,
            'color': color,
            'gradient': bool(getattr(rule, '_data_bar_gradient', True)),
            'direction': getattr(rule, 'direction', None) or 'left-to-right',
            'show_value': bool(getattr(
                rule, '_data_bar_show_value', True
            )),
        }

    @classmethod
    def _data_bar_threshold(cls, spec, values):
        threshold_type = str(spec.get('type', '')).lower()
        if threshold_type == 'automin':
            return min(0.0, min(values))
        if threshold_type == 'automax':
            return max(0.0, max(values))
        return cls._color_scale_threshold(spec, values)

    @classmethod
    def _icon_set_resolution(
        cls, worksheet, rule, row, column, areas
    ):
        icon_set_type = (
            getattr(rule, 'icon_set_type', None) or '3TrafficLights1'
        )
        if icon_set_type not in _SUPPORTED_PDF_ICON_SETS:
            return None

        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        values = cls._conditional_numeric_values(worksheet, areas)
        if value is None or not values:
            return None

        icon_count = cls._icon_set_count(icon_set_type)
        specs = list(
            getattr(rule, '_icon_set_thresholds', ()) or ()
        )
        defaults = cls._default_icon_set_thresholds(icon_count)
        if len(specs) != icon_count:
            specs = defaults

        thresholds = []
        for index, spec in enumerate(specs):
            threshold = cls._color_scale_threshold(spec, values)
            if threshold is None:
                threshold = cls._color_scale_threshold(
                    defaults[index], values
                )
            thresholds.append(threshold)

        icon_index = 0
        for index in range(1, icon_count):
            inclusive = str(
                specs[index].get('gte', '1')
            ).lower() not in ('0', 'false')
            if (
                value >= thresholds[index]
                if inclusive
                else value > thresholds[index]
            ):
                icon_index = index
        if getattr(rule, 'reverse_icons', False):
            icon_index = icon_count - icon_index - 1

        return {
            'icon_set_type': icon_set_type,
            'icon_index': icon_index,
            'show_value': not bool(getattr(rule, 'show_icon_only', False)),
        }

    @staticmethod
    def _icon_set_count(icon_set_type):
        if str(icon_set_type).startswith('4'):
            return 4
        if str(icon_set_type).startswith('5'):
            return 5
        return 3

    @staticmethod
    def _default_icon_set_thresholds(icon_count):
        return [
            {
                'type': 'percent',
                'val': str(int(100 * index / icon_count)),
            }
            for index in range(icon_count)
        ]

    @classmethod
    def _color_scale_color(
        cls, worksheet, rule, row, column, areas
    ):
        ref = worksheet.cells.coordinate_to_string(row, column)
        cell = worksheet.cells.get_all_cells().get(ref)
        value = cls._conditional_numeric_value(
            getattr(cell, 'value', None)
        )
        values = cls._conditional_numeric_values(worksheet, areas)
        if value is None or not values:
            return None

        is_three_color = (
            getattr(rule, 'color_scale_type', None) == '3-color'
            or getattr(rule, 'mid_color', None) is not None
        )
        colors = [getattr(rule, 'min_color', None) or 'FFF8696B']
        if is_three_color:
            colors.append(
                getattr(rule, 'mid_color', None) or 'FFFFEB84'
            )
        colors.append(getattr(rule, 'max_color', None) or 'FF63BE7B')

        defaults = [{'type': 'min'}]
        if is_three_color:
            defaults.append({'type': 'percentile', 'val': '50'})
        defaults.append({'type': 'max'})
        specs = list(
            getattr(rule, '_color_scale_thresholds', ()) or ()
        )
        if len(specs) != len(colors):
            specs = defaults

        thresholds = []
        for index, spec in enumerate(specs):
            threshold = cls._color_scale_threshold(spec, values)
            if threshold is None:
                threshold = cls._color_scale_threshold(
                    defaults[index], values
                )
            thresholds.append(threshold)

        if value <= thresholds[0]:
            return cls._normalize_argb(colors[0])
        for index in range(len(thresholds) - 1):
            lower = thresholds[index]
            upper = thresholds[index + 1]
            if value > upper:
                continue
            if upper <= lower:
                return cls._normalize_argb(colors[index + 1])
            interpolation_lower = lower
            if (
                index > 0
                and str(specs[index].get('type', '')).lower()
                == 'percentile'
            ):
                # Excel evaluates the upper half from the high side of an
                # interpolated percentile boundary before truncating RGB.
                interpolation_lower = math.nextafter(lower, math.inf)
            ratio = (value - interpolation_lower) / (
                upper - interpolation_lower
            )
            return cls._interpolate_argb(
                colors[index], colors[index + 1], ratio
            )
        return cls._normalize_argb(colors[-1])

    @classmethod
    def _color_scale_threshold(cls, spec, values):
        threshold_type = str(spec.get('type', '')).lower()
        if threshold_type in ('min', 'automin'):
            return min(values)
        if threshold_type in ('max', 'automax'):
            return max(values)

        operand = cls._conditional_formula_number(spec.get('val'))
        if operand is None:
            return None
        if threshold_type in ('num', 'formula'):
            return operand
        if threshold_type == 'percent':
            ratio = min(max(operand, 0.0), 100.0) / 100.0
            return min(values) + (max(values) - min(values)) * ratio
        if threshold_type == 'percentile':
            return cls._percentile(values, operand)
        return None

    @staticmethod
    def _percentile(values, percent):
        ordered = sorted(values)
        if len(ordered) == 1:
            return ordered[0]
        ratio = min(max(percent, 0.0), 100.0) / 100.0
        position = (len(ordered) - 1) * ratio
        lower_index = int(math.floor(position))
        upper_index = int(math.ceil(position))
        if lower_index == upper_index:
            return ordered[lower_index]
        fraction = position - lower_index
        return ordered[lower_index] + (
            ordered[upper_index] - ordered[lower_index]
        ) * fraction

    @classmethod
    def _interpolate_argb(cls, start_color, end_color, ratio):
        start = cls._argb_channels(start_color)
        end = cls._argb_channels(end_color)
        if start is None or end is None:
            return None
        ratio = min(max(float(ratio), 0.0), 1.0)
        channels = (
            first + int((second - first) * ratio)
            for first, second in zip(start, end)
        )
        return ''.join(f'{channel:02X}' for channel in channels)

    @classmethod
    def _normalize_argb(cls, color):
        channels = cls._argb_channels(color)
        if channels is None:
            return None
        return ''.join(f'{channel:02X}' for channel in channels)

    @staticmethod
    def _argb_channels(color):
        text = str(color or '').strip().lstrip('#')
        if len(text) == 6:
            text = 'FF' + text
        if len(text) != 8:
            return None
        try:
            return tuple(
                int(text[index:index + 2], 16)
                for index in range(0, 8, 2)
            )
        except ValueError:
            return None

    @classmethod
    def _overlay_conditional_dxf(cls, target, worksheet, rule):
        workbook = getattr(worksheet, '_workbook', None)
        dxf_styles = getattr(workbook, '_dxf_styles', None) or []
        dxf_id = getattr(rule, '_dxf_id', None)
        if dxf_id is not None and 0 <= dxf_id < len(dxf_styles):
            dxf = dxf_styles[dxf_id]
        else:
            dxf = cls._conditional_format_data(rule)

        font_data = dxf.get('font', {})
        for attr in ('bold', 'italic', 'strikethrough', 'color'):
            if attr in font_data:
                setattr(target.font, attr, font_data[attr])
        if 'underline' in font_data:
            target.font.underline = bool(font_data['underline'])
            target.font.underline_type = (
                'single' if target.font.underline else 'none'
            )

        fill_data = dxf.get('fill', {})
        if 'pattern_type' in fill_data:
            target.fill.pattern_type = fill_data['pattern_type']
        if 'fg_color' in fill_data:
            target.fill.foreground_color = fill_data['fg_color']
            target.fill._fg_color_type = 'rgb'
            target.fill._fg_color_value = fill_data['fg_color']
        if 'bg_color' in fill_data:
            target.fill.background_color = fill_data['bg_color']
            target.fill._bg_color_type = 'rgb'
            target.fill._bg_color_value = fill_data['bg_color']

        border_data = dxf.get('border')
        if border_data:
            for side in ('top', 'bottom', 'left', 'right'):
                border = getattr(target.borders, side)
                border.line_style = border_data.get('style', 'thin')
                border.color = border_data.get('color', 'FF000000')

        blank = cls._blank_style()
        for attr in (
            'horizontal', 'vertical', 'wrap_text', 'indent',
            'text_rotation', 'shrink_to_fit', 'reading_order',
            'relative_indent',
        ):
            value = getattr(rule.alignment, attr)
            if value != getattr(blank.alignment, attr):
                setattr(target.alignment, attr, value)
        if rule.number_format != blank.number_format:
            target.number_format = rule.number_format

    @staticmethod
    def _conditional_format_data(rule):
        data = {}
        font = {}
        for attr in ('bold', 'italic', 'underline', 'strikethrough'):
            if getattr(rule.font, attr):
                font[attr] = True
        if rule.font.color != 'FF000000':
            font['color'] = rule.font.color
        if font:
            data['font'] = font
        if rule.fill.pattern_type != 'none':
            data['fill'] = {
                'pattern_type': rule.fill.pattern_type,
                'fg_color': rule.fill.foreground_color,
                'bg_color': rule.fill.background_color,
            }
        if rule.border.line_style != 'none':
            data['border'] = {
                'style': rule.border.line_style,
                'color': rule.border.color,
            }
        return data

    @classmethod
    def _overlay_table_style(cls, target, worksheet, row, column):
        from .table import _parse_cell_range

        for table in worksheet.tables:
            bounds = _parse_cell_range(table.ref)
            if bounds is None:
                continue
            min_row, min_col, max_row, max_col = (
                bounds[0] + 1,
                bounds[1] + 1,
                bounds[2] + 1,
                bounds[3] + 1,
            )
            if not (min_row <= row <= max_row and min_col <= column <= max_col):
                continue
            if table.table_style_info.name == "TableStyleMedium2":
                cls._apply_table_style_medium2(
                    target, table, row, column,
                    min_row, min_col, max_row, max_col,
                )
                return True
            workbook = getattr(worksheet, '_workbook', None)
            custom_style = (
                getattr(workbook, '_table_styles', {}) or {}
            ).get(table.table_style_info.name)
            if custom_style:
                cls._apply_custom_table_style(
                    target, workbook, custom_style, table, row, column,
                    min_row, min_col, max_row, max_col,
                )
                return True
            return False
        return False

    @classmethod
    def _apply_custom_table_style(
        cls, target, workbook, elements, table, row, column,
        min_row, min_col, max_row, max_col,
    ):
        dxf_styles = getattr(workbook, '_dxf_styles', ()) or ()

        def apply(element_type):
            element = elements.get(element_type)
            if element is None:
                return
            dxf_id = element.get('dxf_id')
            if dxf_id is None or not 0 <= dxf_id < len(dxf_styles):
                return
            cls._overlay_table_dxf(
                target, dxf_styles[dxf_id], row, column,
                min_row, min_col, max_row, max_col,
            )

        apply('wholeTable')
        header_row = min_row if table.has_headers else None
        totals_row = max_row if table.show_totals_row else None
        body_start = min_row + (1 if table.has_headers else 0)
        body_end = max_row - (1 if table.show_totals_row else 0)
        style_info = table.table_style_info

        if body_start <= row <= body_end:
            body_row = row - body_start
            body_column = column - min_col
            if style_info.show_row_stripes:
                first = elements.get('firstRowStripe')
                second = elements.get('secondRowStripe')
                first_size = max(1, int((first or {}).get('size', 1)))
                second_size = max(1, int((second or {}).get('size', 1)))
                cycle = first_size + second_size
                apply(
                    'firstRowStripe'
                    if body_row % cycle < first_size
                    else 'secondRowStripe'
                )
            if style_info.show_column_stripes:
                first = elements.get('firstColumnStripe')
                second = elements.get('secondColumnStripe')
                first_size = max(1, int((first or {}).get('size', 1)))
                second_size = max(1, int((second or {}).get('size', 1)))
                cycle = first_size + second_size
                apply(
                    'firstColumnStripe'
                    if body_column % cycle < first_size
                    else 'secondColumnStripe'
                )
        if style_info.show_first_column and column == min_col:
            apply('firstColumn')
        if style_info.show_last_column and column == max_col:
            apply('lastColumn')
        if row == header_row:
            apply('headerRow')
        elif row == totals_row:
            apply('totalRow')

    @staticmethod
    def _overlay_table_dxf(
        target, dxf, row, column, min_row, min_col, max_row, max_col
    ):
        font_data = dxf.get('font', {})
        for attr in ('bold', 'italic', 'strikethrough', 'color'):
            if attr in font_data:
                setattr(target.font, attr, font_data[attr])
        if 'underline' in font_data:
            target.font.underline = bool(font_data['underline'])

        fill_data = dxf.get('fill', {})
        if 'pattern_type' in fill_data:
            target.fill.pattern_type = fill_data['pattern_type']
        if 'fg_color' in fill_data:
            target.fill.foreground_color = fill_data['fg_color']
        if 'bg_color' in fill_data:
            target.fill.background_color = fill_data['bg_color']

        borders = dxf.get('borders', {})
        applicable = {}
        if row == min_row and 'top' in borders:
            applicable['top'] = borders['top']
        if row == max_row and 'bottom' in borders:
            applicable['bottom'] = borders['bottom']
        if column == min_col and 'left' in borders:
            applicable['left'] = borders['left']
        if column == max_col and 'right' in borders:
            applicable['right'] = borders['right']
        if row < max_row and 'horizontal' in borders:
            applicable['bottom'] = borders['horizontal']
        for side, border_data in applicable.items():
            border = getattr(target.borders, side)
            border.line_style = border_data.get('style', 'thin')
            border.color = border_data.get('color', 'FF000000')

    @staticmethod
    def _apply_table_style_medium2(
        target, table, row, column, min_row, min_col, max_row, max_col
    ):
        style_info = table.table_style_info
        border_color = "FF44B3E1"
        header_row = min_row if table.has_headers else None
        totals_row = max_row if table.show_totals_row else None

        if row == header_row:
            target.fill.set_solid_fill("FF156082")
            target.font.bold = True
            target.font.color = "FFFFFFFF"
        elif row == totals_row:
            target.font.bold = True
        else:
            body_row = row - min_row - (1 if table.has_headers else 0)
            if style_info.show_row_stripes and body_row % 2 == 0:
                target.fill.set_solid_fill("FFC0E6F5")

        if style_info.show_column_stripes:
            body_column = column - min_col
            if body_column % 2 == 0:
                target.fill.set_solid_fill("FFC0E6F5")
        if style_info.show_first_column and column == min_col:
            target.font.bold = True
        if style_info.show_last_column and column == max_col:
            target.font.bold = True

        target.borders.bottom.line_style = "medium"
        target.borders.bottom.color = border_color
        if row == min_row:
            target.borders.top.line_style = "medium"
            target.borders.top.color = border_color
        if column == min_col:
            target.borders.left.line_style = "medium"
            target.borders.left.color = border_color
        if column == max_col:
            target.borders.right.line_style = "medium"
            target.borders.right.color = border_color

    @classmethod
    def style_from_index(cls, workbook, style_idx):
        style = cls._get_workbook_default_style(workbook)
        resolved = cls._copy_style(style)
        if workbook is None or style_idx is None:
            return resolved

        try:
            style_idx = int(style_idx)
        except (TypeError, ValueError):
            return resolved

        cell_style_key = getattr(workbook, "_cell_xf_by_index", {}).get(style_idx)
        if cell_style_key is None:
            for candidate_key, candidate_idx in getattr(workbook, "_cell_styles", {}).items():
                if candidate_idx == style_idx:
                    cell_style_key = candidate_key
                    break
        if cell_style_key is None:
            return resolved

        cls._apply_style_key(resolved, workbook, cell_style_key)
        return resolved

    @staticmethod
    def _normalize_column(column):
        if isinstance(column, str):
            from .cells import Cells
            return Cells.column_index_from_string(column)
        column = int(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        return column

    @staticmethod
    def _blank_style():
        return Style()

    @staticmethod
    def _get_workbook_default_style(workbook):
        styles = getattr(workbook, "_styles", None) or []
        if styles:
            return styles[0]
        return Style()

    @staticmethod
    def _has_explicit_style_index(cell):
        return (
            getattr(cell, "_source_style_idx", None) is not None
            or int(getattr(cell, "_style_index", 0) or 0) > 0
        )

    @classmethod
    def _apply_style_key(cls, target, workbook, cell_style_key):
        font_key, fill_key, border_key, num_fmt_key, alignment_key, protection_key = cell_style_key

        font_data = (getattr(workbook, "_font_styles", None) or {}).get(font_key)
        if font_data is not None:
            target.font.name = font_data["name"]
            target.font.size = font_data["size"]
            target.font.color = font_data["color"]
            target.font.bold = font_data["bold"]
            target.font.italic = font_data["italic"]
            target.font.underline = font_data["underline"]
            target.font.underline_type = font_data.get(
                "underline_type",
                "single" if font_data["underline"] else "none",
            )
            target.font.vertical_alignment = font_data.get(
                "vertical_alignment", "baseline"
            )
            target.font.strikethrough = font_data["strikethrough"]

        fill_data = (getattr(workbook, "_fill_styles", None) or {}).get(fill_key)
        if fill_data is not None:
            target.fill.pattern_type = fill_data["pattern_type"]
            target.fill.foreground_color = fill_data["fg_color"]
            target.fill.background_color = fill_data["bg_color"]
            target.fill._fg_color_type = fill_data.get("fg_color_type", "rgb")
            target.fill._fg_color_value = fill_data.get(
                "fg_color_value",
                fill_data.get("fg_color"),
            )
            target.fill._fg_color_tint = fill_data.get("fg_color_tint")
            target.fill._bg_color_type = fill_data.get("bg_color_type", "rgb")
            target.fill._bg_color_value = fill_data.get(
                "bg_color_value",
                fill_data.get("bg_color"),
            )
            target.fill._bg_color_tint = fill_data.get("bg_color_tint")

        border_data = (getattr(workbook, "_border_styles", None) or {}).get(border_key)
        if border_data is not None:
            target.borders.top.line_style = border_data["top"]["style"]
            target.borders.top.color = border_data["top"]["color"]
            target.borders.top.automatic_color = border_data["top"].get(
                "automatic_color", False
            )
            target.borders.bottom.line_style = border_data["bottom"]["style"]
            target.borders.bottom.color = border_data["bottom"]["color"]
            target.borders.bottom.automatic_color = border_data["bottom"].get(
                "automatic_color", False
            )
            target.borders.left.line_style = border_data["left"]["style"]
            target.borders.left.color = border_data["left"]["color"]
            target.borders.left.automatic_color = border_data["left"].get(
                "automatic_color", False
            )
            target.borders.right.line_style = border_data["right"]["style"]
            target.borders.right.color = border_data["right"]["color"]
            target.borders.right.automatic_color = border_data["right"].get(
                "automatic_color", False
            )
            diagonal = border_data.get("diagonal", {})
            target.borders.diagonal.line_style = diagonal.get("style", "none")
            target.borders.diagonal.color = diagonal.get("color", "FF000000")
            target.borders.diagonal_up = border_data.get("diagonal_up", False)
            target.borders.diagonal_down = border_data.get(
                "diagonal_down", False
            )

        num_format = (getattr(workbook, "_num_formats", None) or {}).get(num_fmt_key)
        if num_format is None:
            num_format = NumberFormat.get_builtin_format(num_fmt_key)
        if num_format is not None:
            target.number_format = num_format

        alignment_data = (getattr(workbook, "_alignment_styles", None) or {}).get(alignment_key)
        if alignment_data is not None:
            target.alignment.horizontal = alignment_data["horizontal"]
            target.alignment.vertical = alignment_data["vertical"]
            target.alignment.wrap_text = alignment_data["wrap_text"]
            target.alignment.indent = alignment_data["indent"]
            target.alignment.text_rotation = alignment_data["text_rotation"]
            target.alignment.shrink_to_fit = alignment_data["shrink_to_fit"]
            target.alignment.reading_order = alignment_data["reading_order"]
            target.alignment.relative_indent = alignment_data["relative_indent"]

        protection_data = (getattr(workbook, "_protection_styles", None) or {}).get(protection_key)
        if protection_data is not None:
            target.protection.locked = protection_data["locked"]
            target.protection.hidden = protection_data["hidden"]

    @classmethod
    def _copy_style(cls, source):
        result = Style()
        cls._overlay_style(result, source, full=True)
        return result

    @classmethod
    def _overlay_style(cls, target, source, full, default_style=None):
        if source is None:
            return
        if default_style is None:
            default_style = cls._blank_style()

        cls._overlay_font(target, source, default_style, full)
        cls._overlay_fill(target, source, default_style, full)
        cls._overlay_borders(target, source, default_style, full)
        cls._overlay_alignment(target, source, default_style, full)
        cls._overlay_protection(target, source, default_style, full)
        if full or source.number_format != default_style.number_format:
            target.number_format = source.number_format

    @staticmethod
    def _overlay_font(target, source, default_style, full):
        for attr in (
            "name", "size", "color", "bold", "italic", "underline",
            "underline_type", "vertical_alignment", "strikethrough",
        ):
            if full or getattr(source.font, attr) != getattr(default_style.font, attr):
                setattr(target.font, attr, getattr(source.font, attr))

    @staticmethod
    def _overlay_fill(target, source, default_style, full):
        for attr in ("pattern_type", "foreground_color", "background_color"):
            if full or getattr(source.fill, attr) != getattr(default_style.fill, attr):
                setattr(target.fill, attr, getattr(source.fill, attr))
        for attr in (
            "_fg_color_type",
            "_fg_color_value",
            "_fg_color_tint",
            "_bg_color_type",
            "_bg_color_value",
            "_bg_color_tint",
        ):
            if full or hasattr(source.fill, attr):
                value = getattr(source.fill, attr, None)
                if full or value is not None:
                    setattr(target.fill, attr, value)

    @staticmethod
    def _overlay_borders(target, source, default_style, full):
        for side in ("top", "bottom", "left", "right", "diagonal"):
            source_border = getattr(source.borders, side)
            target_border = getattr(target.borders, side)
            default_border = getattr(default_style.borders, side)
            for attr in (
                "line_style", "color", "weight", "automatic_color"
            ):
                if full or getattr(source_border, attr) != getattr(default_border, attr):
                    setattr(target_border, attr, getattr(source_border, attr))

        for attr in ("diagonal_up", "diagonal_down"):
            if full or getattr(source.borders, attr) != getattr(default_style.borders, attr):
                setattr(target.borders, attr, getattr(source.borders, attr))

    @staticmethod
    def _overlay_alignment(target, source, default_style, full):
        for attr in (
            "horizontal",
            "vertical",
            "wrap_text",
            "indent",
            "text_rotation",
            "shrink_to_fit",
            "reading_order",
            "relative_indent",
        ):
            if full or getattr(source.alignment, attr) != getattr(default_style.alignment, attr):
                setattr(target.alignment, attr, getattr(source.alignment, attr))

    @staticmethod
    def _overlay_protection(target, source, default_style, full):
        for attr in ("locked", "hidden"):
            if full or getattr(source.protection, attr) != getattr(default_style.protection, attr):
                setattr(target.protection, attr, getattr(source.protection, attr))
