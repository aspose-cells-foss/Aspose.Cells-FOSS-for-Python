"""
Aspose.Cells for Python - Display Value Formatter

This module centralizes Excel-like display-value formatting so layout/render
code and text-based exports can share one formatting path.
"""

import re
from fractions import Fraction
from datetime import datetime, date, time, timedelta
from decimal import Decimal, ROUND_HALF_UP


class DisplayValueOptions:
    """Options controlling display-text formatting."""

    def __init__(self):
        self.empty_text = ''
        self.boolean_true_text = 'TRUE'
        self.boolean_false_text = 'FALSE'
        self.date_format = '%Y-%m-%d'
        self.datetime_format = '%Y-%m-%d %H:%M:%S'
        self.time_format = '%H:%M:%S'
        self.evaluate_formulas = True

    @classmethod
    def from_csv_options(cls, options):
        display_options = cls()
        if options is None:
            return display_options
        display_options.date_format = getattr(options, 'date_format', display_options.date_format)
        display_options.datetime_format = getattr(options, 'datetime_format', display_options.datetime_format)
        display_options.time_format = getattr(options, 'time_format', display_options.time_format)
        return display_options


class ValueFormatter:
    """Formats cells or raw values into display text."""

    @staticmethod
    def format_cell(cell, workbook=None, worksheet=None, options=None):
        """Return the visible display text for a cell."""
        if cell is None:
            return ValueFormatter.format_value(None, None, options)

        if options is None:
            options = DisplayValueOptions()

        value = cell.value
        if value is None and options.evaluate_formulas and getattr(cell, 'formula', None) and workbook is not None:
            try:
                from .formula_evaluator import FormulaEvaluator
                evaluator = FormulaEvaluator(workbook)
                value = evaluator.evaluate(cell.formula, worksheet)
            except Exception:
                value = None

        number_format = getattr(getattr(cell, 'style', None), 'number_format', None)
        return ValueFormatter.format_value(value, number_format, options)

    @staticmethod
    def format_value(value, number_format=None, options=None):
        """Return the visible display text for a raw value."""
        if options is None:
            options = DisplayValueOptions()

        if value is None:
            return options.empty_text

        if isinstance(value, bool):
            return options.boolean_true_text if value else options.boolean_false_text

        if isinstance(value, datetime):
            return ValueFormatter._format_datetime_like_excel(value, number_format, options.datetime_format)

        if isinstance(value, date):
            dt_value = datetime(value.year, value.month, value.day)
            return ValueFormatter._format_datetime_like_excel(dt_value, number_format, options.date_format)

        if isinstance(value, time):
            if number_format and number_format not in ('', 'General', 'general', '@'):
                dt_value = datetime(1899, 12, 30, value.hour, value.minute, value.second, value.microsecond)
                return ValueFormatter._format_datetime_like_excel(dt_value, number_format, options.time_format)
            return value.strftime(options.time_format)

        if isinstance(value, (int, float)):
            formatted_number = ValueFormatter._format_number_with_format(value, number_format)
            if formatted_number is not None:
                return formatted_number
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return ValueFormatter._format_float_like_excel(value)

        return str(value)

    @staticmethod
    def _format_datetime_like_excel(value, number_format, default_format):
        if number_format and number_format not in ('', 'General', 'general', '@'):
            serial_date = ValueFormatter._datetime_to_excel_serial(value)
            return ValueFormatter._format_date_with_excel_format(serial_date, number_format)
        return value.strftime(default_format)

    @staticmethod
    def _format_float_like_excel(value):
        formatted = f'{value:.10g}'
        if 'e' in formatted.lower() and abs(value) < 1e-10:
            return '0'
        return formatted

    @staticmethod
    def _format_number_with_format(value, format_code):
        if format_code is None:
            return None

        if format_code == '' or format_code.lower() == 'general' or format_code == '@':
            return None

        sections = format_code.split(';')
        section = ValueFormatter._select_number_format_section(value, sections)
        value_to_format = value

        has_conditions = any(ValueFormatter._section_condition(item) for item in sections)
        if len(sections) > 1 and not has_conditions:
            if value < 0:
                section = sections[1]
                value_to_format = abs(value)
            elif value == 0 and len(sections) > 2:
                section = sections[2]
            else:
                section = sections[0]

        date_section = section
        if ValueFormatter._is_date_format(date_section):
            return ValueFormatter._format_date_with_excel_format(value, date_section, format_code)

        section = ValueFormatter._normalize_format_section(section)

        fraction_text = ValueFormatter._format_fraction_with_format(value_to_format, section)
        if fraction_text is not None:
            return fraction_text

        if not re.search(r'[0#?]', section):
            return ValueFormatter._clean_format_literal(section)

        first_idx = None
        last_idx = None
        for idx, ch in enumerate(section):
            if ch in '0#?':
                if first_idx is None:
                    first_idx = idx
                last_idx = idx

        if first_idx is None:
            return ValueFormatter._clean_format_literal(section)

        prefix_raw = section[:first_idx]
        suffix_raw = section[last_idx + 1:]
        prefix = ValueFormatter._clean_format_literal(prefix_raw)
        suffix = ValueFormatter._clean_format_literal(suffix_raw)

        if '%' in section:
            value_to_format *= 100

        if 'E' in section or 'e' in section:
            decimals = 0
            match = re.search(r'\.(?P<frac>[0#?]+)[eE]', section)
            if match:
                decimals = len(match.group('frac'))
            formatted = f"{value_to_format:.{decimals}E}"
            return f"{prefix}{formatted}{suffix}"

        number_pattern = section[first_idx:last_idx + 1]
        pattern_clean = re.sub(r'[^0#.,]', '', number_pattern)
        if '.' in pattern_clean:
            int_part, frac_part = pattern_clean.split('.', 1)
        else:
            int_part, frac_part = pattern_clean, ''

        use_grouping = ',' in int_part
        min_decimals = frac_part.count('0')
        max_decimals = sum(1 for ch in frac_part if ch in '0#')

        if max_decimals == 0:
            rounded = int(Decimal(str(value_to_format)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            integer_template = ValueFormatter._format_integer_template(rounded, section)
            if integer_template is not None:
                return integer_template
            formatted = f'{rounded:,}' if use_grouping else str(rounded)
        else:
            quant = Decimal('0.' + '0' * max_decimals)
            rounded_d = Decimal(str(value_to_format)).quantize(quant, rounding=ROUND_HALF_UP)
            rounded_f = float(rounded_d)
            formatted = f'{rounded_f:,.{max_decimals}f}' if use_grouping else f'{rounded_f:.{max_decimals}f}'
            if max_decimals > min_decimals and '.' in formatted:
                int_text, frac_text = formatted.split('.', 1)
                frac_text = frac_text.rstrip('0')
                if len(frac_text) < min_decimals:
                    frac_text = frac_text.ljust(min_decimals, '0')
                formatted = int_text if frac_text == '' else f"{int_text}.{frac_text}"

        return f"{prefix}{formatted}{suffix}"

    @staticmethod
    def _section_condition(section):
        for token in re.findall(r'\[([^\]]+)\]', section):
            match = re.fullmatch(r'\s*(<=|>=|<>|=|<|>)\s*(-?\d+(?:\.\d+)?)\s*', token)
            if match:
                return match.group(1), Decimal(match.group(2))
        return None

    @staticmethod
    def _select_number_format_section(value, sections):
        conditions = [ValueFormatter._section_condition(section) for section in sections]
        if not any(conditions):
            return sections[0] if sections else ''

        numeric_value = Decimal(str(value))
        fallback = None
        comparisons = {
            '<': lambda left, right: left < right,
            '<=': lambda left, right: left <= right,
            '>': lambda left, right: left > right,
            '>=': lambda left, right: left >= right,
            '=': lambda left, right: left == right,
            '<>': lambda left, right: left != right,
        }
        for section, condition in zip(sections, conditions):
            if condition is None:
                if fallback is None:
                    fallback = section
                continue
            operator, boundary = condition
            if comparisons[operator](numeric_value, boundary):
                return section
        return fallback if fallback is not None else sections[0]

    @staticmethod
    def _format_integer_template(value, section):
        """Render integer placeholders while preserving embedded literals."""
        if any(ch in section for ch in '.,%Ee'):
            return None

        tokens = []
        idx = 0
        while idx < len(section):
            ch = section[idx]
            if ch == '[':
                end = section.find(']', idx + 1)
                if end < 0:
                    return None
                idx = end + 1
                continue
            if ch == '"':
                end = section.find('"', idx + 1)
                if end < 0:
                    return None
                tokens.extend(('literal', item) for item in section[idx + 1:end])
                idx = end + 1
                continue
            if ch == '\\' and idx + 1 < len(section):
                tokens.append(('literal', section[idx + 1]))
                idx += 2
                continue
            if ch == '_':
                tokens.append(('literal', ' '))
                idx += 2
                continue
            if ch == '*':
                idx += 2
                continue
            if ch in '0#?':
                tokens.append(('placeholder', ch))
            else:
                tokens.append(('literal', ch))
            idx += 1

        placeholder_indexes = [
            index for index, token in enumerate(tokens) if token[0] == 'placeholder'
        ]
        if not placeholder_indexes:
            return None

        negative = value < 0
        digits = list(str(abs(value)))
        rendered = [''] * len(tokens)
        for index in reversed(placeholder_indexes):
            placeholder = tokens[index][1]
            if digits:
                rendered[index] = digits.pop()
            elif placeholder == '0':
                rendered[index] = '0'
            elif placeholder == '?':
                rendered[index] = ' '
        for index, token in enumerate(tokens):
            if token[0] == 'literal':
                rendered[index] = token[1]

        prefix = ''.join(digits)
        if negative:
            prefix = '-' + prefix
        return prefix + ''.join(rendered)

    @staticmethod
    def _clean_format_literal(text):
        result = []
        idx = 0
        while idx < len(text):
            ch = text[idx]
            if ch == '"':
                idx += 1
                while idx < len(text) and text[idx] != '"':
                    result.append(text[idx])
                    idx += 1
                idx += 1
                continue
            if ch == '_':
                result.append(' ')
                idx += 2
                continue
            if ch == '*':
                idx += 2
                continue
            if ch == '\\':
                if idx + 1 < len(text):
                    result.append(text[idx + 1])
                    idx += 2
                else:
                    idx += 1
                continue
            result.append(ch)
            idx += 1
        return ''.join(result)

    @staticmethod
    def _normalize_format_section(section):
        def replace_token(match):
            token = match.group(1)
            if token.startswith('$'):
                currency = token[1:]
                if '-' in currency:
                    currency = currency.split('-', 1)[0]
                return currency
            return ''

        return re.sub(r'\[([^\]]+)\]', replace_token, section)

    @staticmethod
    def _strip_literals_for_analysis(format_code):
        result = []
        idx = 0
        while idx < len(format_code):
            ch = format_code[idx]
            if ch == '"':
                idx += 1
                while idx < len(format_code) and format_code[idx] != '"':
                    idx += 1
                idx += 1
                continue
            if ch == '[':
                while idx < len(format_code) and format_code[idx] != ']':
                    idx += 1
                idx += 1
                continue
            if ch == '\\':
                idx += 2
                continue
            if ch == '_':
                idx += 2
                continue
            if ch == '*':
                idx += 2
                continue
            result.append(ch)
            idx += 1
        return ''.join(result)

    @staticmethod
    def _is_date_format(format_code):
        if not format_code:
            return False

        _, _, cleaned_format = ValueFormatter._extract_date_format_metadata(format_code)
        for token in ValueFormatter._tokenize_excel_date_format(cleaned_format):
            if token['kind'] != 'literal':
                return True
        return False

    @staticmethod
    def _format_date_with_excel_format(serial_date, format_code, original_format_code=None):
        excel_epoch = datetime(1899, 12, 30)

        try:
            serial_decimal = Decimal(str(serial_date))
            days = int(serial_decimal)
            fraction = serial_decimal - Decimal(days)
            rounded_seconds = int((fraction * Decimal('86400')).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            dt = excel_epoch + timedelta(days=days, seconds=rounded_seconds)
        except (ValueError, OverflowError):
            return str(serial_date)

        try:
            return ValueFormatter._render_excel_date_format(dt, serial_decimal, format_code, original_format_code)
        except Exception:
            python_format = ValueFormatter._excel_date_format_to_python(format_code)
            try:
                return dt.strftime(python_format)
            except ValueError:
                return dt.strftime('%Y-%m-%d')

    @staticmethod
    def _render_excel_date_format(dt, serial_date, format_code, original_format_code=None):
        locale_code, dbnum, cleaned_format = ValueFormatter._extract_date_format_metadata(format_code)

        normalized_original = (original_format_code or format_code or '').lower()
        if locale_code == 'F400':
            cleaned_format = 'h:mm:ss'
        elif cleaned_format.lower() == 'm/d/yyyy' and normalized_original == 'm/d/yyyy':
            cleaned_format = 'yyyy/m/d'

        if locale_code == 'F800':
            cleaned_format = 'yyyy"年"m"月"d"日"'

        special_time = ValueFormatter._render_special_time_format(serial_date, cleaned_format)
        if special_time is not None:
            return special_time

        tokens = ValueFormatter._tokenize_excel_date_format(cleaned_format)
        ValueFormatter._resolve_month_minute_tokens(tokens)
        has_ampm = any(token['kind'] in ('ampm', 'ampm_local') for token in tokens)

        rendered = []
        for token in tokens:
            if token['kind'] == 'literal':
                rendered.append(token['value'])
            else:
                rendered.append(
                    ValueFormatter._render_excel_date_token(
                        dt,
                        token['kind'],
                        token['value'],
                        locale_code,
                        dbnum,
                        has_ampm,
                    )
                )
        return ''.join(rendered)

    @staticmethod
    def _render_special_time_format(serial_date, cleaned_format):
        format_lower = cleaned_format.lower()
        total_seconds_exact = Decimal(str(serial_date)) * Decimal('86400')

        if format_lower == 'mm:ss.0':
            seconds_rounded = total_seconds_exact.quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            minute = int(seconds_rounded // Decimal('60')) % 60
            second = float(seconds_rounded % Decimal('60'))
            return f'{minute:02d}:{second:04.1f}'

        if format_lower == '[h]:mm:ss':
            seconds_rounded = int(total_seconds_exact.quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            hours = seconds_rounded // 3600
            minute = (seconds_rounded // 60) % 60
            second = seconds_rounded % 60
            return f'{hours}:{minute:02d}:{second:02d}'

        return None

    @staticmethod
    def _format_fraction_with_format(value, section):
        match = re.fullmatch(r'(?P<whole>[#0?]+)(?:\\ | )(?P<num>[?]+)\/(?P<den>[?]+|\d+)', section)
        if not match:
            return None

        sign = '-' if value < 0 else ''
        abs_value = abs(value)
        whole_value = int(abs_value)
        fraction_value = abs_value - whole_value

        denominator_pattern = match.group('den')
        numerator_width = len(match.group('num'))

        if set(denominator_pattern) == {'?'}:
            max_denominator = (10 ** len(denominator_pattern)) - 1
            fraction = Fraction(fraction_value).limit_denominator(max_denominator)
            numerator = fraction.numerator
            denominator = fraction.denominator
        else:
            denominator = int(denominator_pattern)
            numerator = int(Decimal(str(fraction_value * denominator)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

        if numerator >= denominator:
            whole_value += numerator // denominator
            numerator = numerator % denominator

        if numerator == 0:
            return f'{sign}{whole_value}'

        numerator_text = str(numerator).rjust(numerator_width)
        fraction_text = f'{numerator_text}/{denominator}'
        if whole_value == 0:
            return f'{sign} {fraction_text}'
        return f'{sign}{whole_value} {fraction_text}'

    @staticmethod
    def _extract_date_format_metadata(format_code):
        locale_code = None
        dbnum = None
        cleaned = []
        idx = 0

        while idx < len(format_code):
            ch = format_code[idx]
            if ch == '[':
                end = format_code.find(']', idx + 1)
                if end == -1:
                    cleaned.append(ch)
                    idx += 1
                    continue
                token = format_code[idx + 1:end]
                token_lower = token.lower()
                if token_lower.startswith('dbnum'):
                    dbnum = token
                elif token.startswith('$-'):
                    locale_code = token[2:].upper()
                elif token_lower in ('h', 'hh', 'm', 'mm', 's', 'ss'):
                    cleaned.append(format_code[idx:end + 1])
                idx = end + 1
                continue
            cleaned.append(ch)
            idx += 1

        return locale_code, dbnum, ''.join(cleaned)

    @staticmethod
    def _tokenize_excel_date_format(format_code):
        tokens = []
        idx = 0

        while idx < len(format_code):
            if format_code[idx] == '[':
                end = format_code.find(']', idx + 1)
                if end != -1:
                    token = format_code[idx + 1:end].lower()
                    if token in ('h', 'hh'):
                        tokens.append({'kind': 'elapsed_hour', 'value': format_code[idx:end + 1]})
                        idx = end + 1
                        continue
                    if token in ('m', 'mm'):
                        tokens.append({'kind': 'elapsed_minute', 'value': format_code[idx:end + 1]})
                        idx = end + 1
                        continue
                    if token in ('s', 'ss'):
                        tokens.append({'kind': 'elapsed_second', 'value': format_code[idx:end + 1]})
                        idx = end + 1
                        continue
            if format_code[idx:idx + 5] == '上午/下午':
                tokens.append({'kind': 'ampm_local', 'value': '上午/下午'})
                idx += 5
                continue

            if format_code[idx:idx + 5].lower() == 'am/pm':
                tokens.append({'kind': 'ampm', 'value': format_code[idx:idx + 5]})
                idx += 5
                continue

            if format_code[idx:idx + 3].lower() == 'a/p':
                tokens.append({'kind': 'ampm', 'value': format_code[idx:idx + 3]})
                idx += 3
                continue

            ch = format_code[idx]
            lower = ch.lower()

            if ch == '"':
                idx += 1
                literal = []
                while idx < len(format_code) and format_code[idx] != '"':
                    literal.append(format_code[idx])
                    idx += 1
                idx += 1
                tokens.append({'kind': 'literal', 'value': ''.join(literal)})
                continue

            if ch == '\\':
                if idx + 1 < len(format_code):
                    tokens.append({'kind': 'literal', 'value': format_code[idx + 1]})
                    idx += 2
                else:
                    idx += 1
                continue

            if ch == '_':
                tokens.append({'kind': 'literal', 'value': ' '})
                idx += 2
                continue

            if ch == '*':
                idx += 2
                continue

            prev_char = format_code[idx - 1] if idx > 0 else ''
            next_char = format_code[idx + 1] if idx + 1 < len(format_code) else ''
            if lower in ('y', 'm', 'd', 'h', 's', 'a') and not ValueFormatter._is_ascii_alpha(prev_char):
                count = 1
                while idx + count < len(format_code) and format_code[idx + count].lower() == lower:
                    count += 1
                after_token = format_code[idx + count] if idx + count < len(format_code) else ''
                if ValueFormatter._is_ascii_alpha(next_char) and next_char.lower() != lower:
                    tokens.append({'kind': 'literal', 'value': ch})
                    idx += 1
                    continue
                if ValueFormatter._is_ascii_alpha(after_token):
                    tokens.append({'kind': 'literal', 'value': format_code[idx:idx + count]})
                    idx += count
                    continue
                token = format_code[idx:idx + count]
                token_kind = {
                    'y': 'year',
                    'm': 'month_or_minute',
                    'd': 'day',
                    'h': 'hour',
                    's': 'second',
                    'a': 'weekday_local',
                }[lower]
                tokens.append({'kind': token_kind, 'value': token})
                idx += count
                continue

            tokens.append({'kind': 'literal', 'value': ch})
            idx += 1

        return tokens

    @staticmethod
    def _resolve_month_minute_tokens(tokens):
        semantic_tokens = [token for token in tokens if token['kind'] != 'literal']
        for index, token in enumerate(semantic_tokens):
            if token['kind'] != 'month_or_minute':
                continue
            previous_kind = semantic_tokens[index - 1]['kind'] if index > 0 else None
            next_kind = semantic_tokens[index + 1]['kind'] if index + 1 < len(semantic_tokens) else None
            if previous_kind == 'hour' or next_kind == 'second':
                token['kind'] = 'minute'
            else:
                token['kind'] = 'month'

    @staticmethod
    def _render_excel_date_token(dt, kind, token, locale_code, dbnum, has_ampm=False):
        token_lower = token.lower()

        if kind == 'year':
            if dbnum and token_lower in ('yy', 'yyyy'):
                year_text = f'{dt.year % 100:02d}' if token_lower == 'yy' else str(dt.year)
                return ''.join(ValueFormatter._to_chinese_digit(char) for char in year_text)
            return f'{dt.year % 100:02d}' if token_lower == 'yy' else str(dt.year)

        if kind == 'month':
            if token_lower == 'mmmmm':
                return ValueFormatter._month_narrow(dt.month)
            if token_lower == 'mmmm':
                return ValueFormatter._month_name(dt.month)
            if token_lower == 'mmm':
                return ValueFormatter._month_abbrev(dt.month)
            if dbnum:
                return ValueFormatter._to_chinese_number(dt.month)
            return f'{dt.month:02d}' if len(token) >= 2 else str(dt.month)

        if kind == 'day':
            if token_lower == 'dddd':
                return ValueFormatter._weekday_name(dt.weekday())
            if token_lower == 'ddd':
                return ValueFormatter._weekday_abbrev(dt.weekday())
            if dbnum:
                return ValueFormatter._to_chinese_number(dt.day)
            return f'{dt.day:02d}' if len(token) >= 2 else str(dt.day)

        if kind == 'weekday_local':
            return ValueFormatter._weekday_local(dt.weekday(), token_lower, locale_code)

        if kind == 'hour':
            hour = dt.hour
            if has_ampm:
                hour = hour % 12 or 12
            if dbnum:
                return ValueFormatter._to_chinese_number(hour)
            return f'{hour:02d}' if len(token) >= 2 else str(hour)

        if kind == 'elapsed_hour':
            total_seconds = int((serial_date * Decimal('86400')).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            return str(total_seconds // 3600)

        if kind == 'elapsed_minute':
            total_seconds = int((serial_date * Decimal('86400')).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            total_minutes = total_seconds // 60
            return f'{total_minutes:02d}' if 'mm' in token.lower() else str(total_minutes)

        if kind == 'elapsed_second':
            total_seconds = int((serial_date * Decimal('86400')).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            return f'{total_seconds:02d}' if 'ss' in token.lower() else str(total_seconds)

        if kind == 'minute':
            if dbnum:
                return ValueFormatter._to_chinese_number(dt.minute)
            return f'{dt.minute:02d}' if len(token) >= 2 else str(dt.minute)

        if kind == 'second':
            if dbnum:
                return ValueFormatter._to_chinese_number(dt.second)
            return f'{dt.second:02d}' if len(token) >= 2 else str(dt.second)

        if kind == 'ampm':
            return 'AM' if dt.hour < 12 else 'PM'

        if kind == 'ampm_local':
            return '上午' if dt.hour < 12 else '下午'

        return token

    @staticmethod
    def _month_name(month):
        return [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December',
        ][month - 1]

    @staticmethod
    def _month_abbrev(month):
        return [
            'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
        ][month - 1]

    @staticmethod
    def _month_narrow(month):
        return ValueFormatter._month_name(month)[0]

    @staticmethod
    def _weekday_name(weekday):
        return [
            'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday',
        ][weekday]

    @staticmethod
    def _weekday_abbrev(weekday):
        return ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'][weekday]

    @staticmethod
    def _is_ascii_alpha(char):
        return ('a' <= char <= 'z') or ('A' <= char <= 'Z')

    @staticmethod
    def _weekday_local(weekday, token_lower, locale_code):
        if locale_code == '804':
            if len(token_lower) >= 4:
                return ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'][weekday]
            return ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][weekday]
        return ValueFormatter._weekday_abbrev(weekday)

    @staticmethod
    def _to_chinese_digit(char):
        return {
            '0': '零',
            '1': '一',
            '2': '二',
            '3': '三',
            '4': '四',
            '5': '五',
            '6': '六',
            '7': '七',
            '8': '八',
            '9': '九',
        }.get(char, char)

    @staticmethod
    def _to_chinese_number(number):
        digits = {
            0: '零',
            1: '一',
            2: '二',
            3: '三',
            4: '四',
            5: '五',
            6: '六',
            7: '七',
            8: '八',
            9: '九',
            10: '十',
        }
        if number <= 10:
            return digits[number]
        if number < 20:
            return '十' + digits[number % 10]
        tens, ones = divmod(number, 10)
        text = digits[tens] + '十'
        if ones:
            text += digits[ones]
        return text

    @staticmethod
    def _excel_date_format_to_python(excel_format):
        result = []
        i = 0
        fmt = excel_format

        while i < len(fmt):
            ch = fmt[i]
            ch_lower = ch.lower()

            if ch_lower == 'y':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'y':
                    count += 1
                result.append('%Y' if count >= 4 else '%y')
                i += count
                continue

            if ch_lower == 'm':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'm':
                    count += 1

                is_minute = False
                for j in range(i - 1, -1, -1):
                    if fmt[j].lower() == 'h':
                        is_minute = True
                        break
                    if fmt[j].lower() in 'yds':
                        break

                if is_minute:
                    result.append('%M')
                elif count >= 4:
                    result.append('%B')
                elif count == 3:
                    result.append('%b')
                else:
                    result.append('%m')
                i += count
                continue

            if ch_lower == 'd':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'd':
                    count += 1
                if count >= 4:
                    result.append('%A')
                elif count == 3:
                    result.append('%a')
                else:
                    result.append('%d')
                i += count
                continue

            if ch_lower == 'h':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'h':
                    count += 1
                if 'am' in fmt.lower() or 'pm' in fmt.lower():
                    result.append('%I')
                else:
                    result.append('%H')
                i += count
                continue

            if ch_lower == 's':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 's':
                    count += 1
                result.append('%S')
                i += count
                continue

            if ch_lower == 'a' and i + 1 < len(fmt) and fmt[i + 1].lower() == 'm':
                if i + 4 < len(fmt) and fmt[i:i + 5].lower() == 'am/pm':
                    result.append('%p')
                    i += 5
                elif i + 1 < len(fmt) and fmt[i:i + 2].lower() == 'am':
                    result.append('%p')
                    i += 2
                else:
                    result.append(ch)
                    i += 1
                continue

            if ch == '\\' and i + 1 < len(fmt):
                result.append(fmt[i + 1])
                i += 2
                continue

            if ch == '"':
                i += 1
                while i < len(fmt) and fmt[i] != '"':
                    result.append(fmt[i])
                    i += 1
                i += 1
                continue

            result.append(ch)
            i += 1

        return ''.join(result)

    @staticmethod
    def _datetime_to_excel_serial(value):
        excel_epoch = datetime(1899, 12, 30)
        delta = value - excel_epoch
        return delta.days + (delta.seconds + delta.microseconds / 1000000.0) / 86400.0
