"""
Shared text measurement and font resolution for layout/rendering.
"""

from collections import namedtuple
from functools import lru_cache
import os
import re
import unicodedata

try:
    from PIL import ImageFont
except ImportError:  # pragma: no cover - optional dependency path
    ImageFont = None


FontResolution = namedtuple(
    "FontResolution",
    [
        "requested_family",
        "resolved_family",
        "font_path",
        "size_points",
        "size_pixels",
        "bold",
        "italic",
        "fallback_chain",
        "backend",
    ],
)

MeasuredText = namedtuple(
    "MeasuredText",
    [
        "text",
        "width_points",
        "height_points",
        "line_count",
        "line_widths",
        "line_height_points",
        "mdw_points",
        "lines",
        "runs",
        "font_resolution",
    ],
)

TextRun = namedtuple("TextRun", ["text", "family", "width_points"])


class FontStrategy:
    """
    Resolves effective font families and fallback chains for text measurement.
    """

    DEFAULT_LATIN_FALLBACKS = (
        "Calibri",
        "Arial",
        "Segoe UI",
        "DejaVu Sans",
    )
    DEFAULT_CJK_FALLBACKS = (
        "Microsoft YaHei",
        "SimSun",
        "Noto Sans CJK SC",
        "WenQuanYi Zen Hei",
        "Arial Unicode MS",
        "DejaVu Sans",
    )
    DEFAULT_SYMBOL_FALLBACKS = (
        "Segoe UI Emoji",
        "Segoe UI Symbol",
        "Noto Color Emoji",
        "Arial Unicode MS",
        "DejaVu Sans",
    )

    _FONT_FILE_CANDIDATES = {
        "等线": ("Deng.ttf", "Dengb.ttf", "Dengl.ttf"),
        "dengxian": ("Deng.ttf", "Dengb.ttf", "Dengl.ttf"),
        "aptos narrow": (
            "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"
        ),
        "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
        "arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
        "tahoma": ("tahoma.ttf", "tahomabd.ttf"),
        "times new roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
        "courier new": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
        "segoe ui": ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
        "segoe ui light": ("segoeuil.ttf",),
        "segoe ui emoji": ("seguiemj.ttf",),
        "segoe ui symbol": ("seguisym.ttf",),
        "microsoft yahei": (
            "msyh.ttc", "msyhbd.ttc", "msyhl.ttc", "msyh.ttf", "msyhbd.ttf"
        ),
        "simsun": ("simsun.ttc", "simsun.ttf"),
        "noto sans cjk sc": ("NotoSansCJKsc-Regular.otf", "NotoSansCJK-Regular.ttc"),
        "arial unicode ms": ("ARIALUNI.TTF",),
        "dejavu sans": ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf"),
        "noto color emoji": ("NotoColorEmoji.ttf",),
        "wenquanyi zen hei": ("wqy-zenhei.ttc", "wqy-zenhei.ttf"),
    }

    _FONT_STYLE_FILE_CANDIDATES = {
        "等线": {
            (False, False): ("Deng.ttf",),
            (True, False): ("Dengb.ttf",),
            (False, True): ("Deng.ttf",),
            (True, True): ("Dengb.ttf",),
        },
        "dengxian": {
            (False, False): ("Deng.ttf",),
            (True, False): ("Dengb.ttf",),
            (False, True): ("Deng.ttf",),
            (True, True): ("Dengb.ttf",),
        },
        "aptos narrow": {
            (False, False): ("calibri.ttf",),
            (True, False): ("calibrib.ttf",),
            (False, True): ("calibrii.ttf",),
            (True, True): ("calibriz.ttf",),
        },
        "calibri": {
            (False, False): ("calibri.ttf",),
            (True, False): ("calibrib.ttf",),
            (False, True): ("calibrii.ttf",),
            (True, True): ("calibriz.ttf",),
        },
        "arial": {
            (False, False): ("arial.ttf",),
            (True, False): ("arialbd.ttf",),
            (False, True): ("ariali.ttf",),
            (True, True): ("arialbi.ttf",),
        },
        "tahoma": {
            (False, False): ("tahoma.ttf",),
            (True, False): ("tahomabd.ttf",),
            (False, True): ("tahoma.ttf",),
            (True, True): ("tahomabd.ttf",),
        },
        "times new roman": {
            (False, False): ("times.ttf",),
            (True, False): ("timesbd.ttf",),
            (False, True): ("timesi.ttf",),
            (True, True): ("timesbi.ttf",),
        },
        "courier new": {
            (False, False): ("cour.ttf",),
            (True, False): ("courbd.ttf",),
            (False, True): ("couri.ttf",),
            (True, True): ("courbi.ttf",),
        },
        "segoe ui": {
            (False, False): ("segoeui.ttf",),
            (True, False): ("segoeuib.ttf",),
            (False, True): ("segoeuii.ttf",),
            (True, True): ("segoeuiz.ttf",),
        },
        "segoe ui light": {
            (False, False): ("segoeuil.ttf",),
            (True, False): ("segoeuil.ttf",),
            (False, True): ("segoeuil.ttf",),
            (True, True): ("segoeuil.ttf",),
        },
        "dejavu sans": {
            (False, False): ("DejaVuSans.ttf",),
            (True, False): ("DejaVuSans-Bold.ttf",),
            (False, True): ("DejaVuSans-Oblique.ttf",),
            (True, True): ("DejaVuSans-BoldOblique.ttf",),
        },
    }

    def __init__(self, default_fallbacks=None):
        self._default_fallbacks = tuple(default_fallbacks or ())

    @property
    def default_fallbacks(self):
        return self._default_fallbacks

    @default_fallbacks.setter
    def default_fallbacks(self, families):
        self._default_fallbacks = tuple(families or ())

    def resolve(self, font):
        family = getattr(font, "name", None) or "Calibri"
        size_points = float(getattr(font, "size", 11) or 11)
        bold = bool(getattr(font, "bold", False))
        italic = bool(getattr(font, "italic", False))

        fallback_chain = self._build_fallback_chain(family)
        font_path = None
        resolved_family = family
        backend = "heuristic"

        for candidate in fallback_chain:
            candidate_path = self._find_font_path(candidate, bold=bold, italic=italic)
            if candidate_path:
                font_path = candidate_path
                resolved_family = candidate
                backend = "pillow"
                break

        return FontResolution(
            requested_family=family,
            resolved_family=resolved_family,
            font_path=font_path,
            size_points=size_points,
            size_pixels=max(1, int(round(size_points * 96.0 / 72.0))),
            bold=bold,
            italic=italic,
            fallback_chain=fallback_chain,
            backend=backend,
        )

    def choose_family_for_text(self, text, resolution):
        if self._contains_emoji(text):
            for family in resolution.fallback_chain:
                if (
                    family in self.DEFAULT_SYMBOL_FALLBACKS
                    and self._find_font_path(family)
                ):
                    return family
        if self._contains_cjk(text):
            for family in resolution.fallback_chain:
                if (
                    family in self.DEFAULT_CJK_FALLBACKS
                    and self._find_font_path(family)
                ):
                    return family
        return resolution.resolved_family

    def _build_fallback_chain(self, family):
        chain = []
        for candidate in (
            family,
            *self._default_fallbacks,
            *self.DEFAULT_LATIN_FALLBACKS,
            *self.DEFAULT_CJK_FALLBACKS,
            *self.DEFAULT_SYMBOL_FALLBACKS,
        ):
            if candidate and candidate not in chain:
                chain.append(candidate)
        return tuple(chain)

    @classmethod
    def _contains_cjk(cls, text):
        for char in text:
            code = ord(char)
            if (
                0x3400 <= code <= 0x4DBF
                or 0x4E00 <= code <= 0x9FFF
                or 0x3040 <= code <= 0x30FF
                or 0xAC00 <= code <= 0xD7AF
            ):
                return True
        return False

    @classmethod
    def _contains_emoji(cls, text):
        for char in text:
            code = ord(char)
            if (
                0x1F300 <= code <= 0x1FAFF
                or 0x2600 <= code <= 0x27BF
            ):
                return True
        return False

    @classmethod
    @lru_cache(maxsize=256)
    def _find_font_path(cls, family, bold=False, italic=False):
        family_key = (family or "").strip().lower()
        styled = cls._FONT_STYLE_FILE_CANDIDATES.get(family_key, {}).get(
            (bool(bold), bool(italic)), ()
        )
        candidates = tuple(dict.fromkeys(
            styled + cls._FONT_FILE_CANDIDATES.get(family_key, ())
        ))
        search_dirs = [
            os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"),
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            os.path.expanduser("~/.fonts"),
        ]
        for search_dir in search_dirs:
            if not os.path.isdir(search_dir):
                continue
            for file_name in candidates:
                candidate_path = os.path.join(search_dir, file_name)
                if os.path.exists(candidate_path):
                    return candidate_path
        return None


class TextMeasurer:
    """
    Backend-neutral text measurement service used by layout and rendering.
    """

    def __init__(self, font_strategy=None):
        self.font_strategy = font_strategy or FontStrategy()

    def measure_text(self, text, font, width_points=None, wrap_text=False):
        resolution = self.font_strategy.resolve(font)
        mdw_points = self.measure_mdw(font)
        line_height_points = self._measure_line_height_points(resolution)

        lines = []
        line_widths = []
        runs = []

        paragraphs = str(text or "").splitlines() or [""]
        if not paragraphs:
            paragraphs = [""]

        for paragraph in paragraphs:
            paragraph_lines = self._wrap_paragraph(
                paragraph,
                resolution,
                width_points=width_points,
                wrap_text=wrap_text,
            )
            if not paragraph_lines:
                paragraph_lines = [""]
            for line in paragraph_lines:
                width = self._measure_text_width_points(line, resolution)
                line_widths.append(width)
                lines.append(line)
                line_runs = self._split_runs(line, resolution)
                runs.extend(line_runs)

        width = max(line_widths) if line_widths else 0.0
        height = line_height_points * max(1, len(lines))
        return MeasuredText(
            text=str(text or ""),
            width_points=width,
            height_points=height,
            line_count=max(1, len(lines)),
            line_widths=tuple(line_widths),
            line_height_points=line_height_points,
            mdw_points=mdw_points,
            lines=tuple(lines),
            runs=tuple(runs),
            font_resolution=resolution,
        )

    def measure_mdw(self, font):
        resolution = self.font_strategy.resolve(font)
        return self._measure_text_width_points("0", resolution)

    def measure_unhinted_mdw(self, font):
        """Measure the Normal-font digit advance without screen hinting."""
        resolution = self.font_strategy.resolve(font)
        if ImageFont is None or resolution.font_path is None:
            return None
        target_pixels = max(
            resolution.size_points * 96.0 / 72.0, 0.01
        )
        sample_size = max(256, int(round(target_pixels * 100.0)))
        sample_resolution = resolution._replace(size_pixels=sample_size)
        sample_font = self._load_font(sample_resolution)
        if sample_font is None:
            return None
        try:
            if hasattr(sample_font, 'getlength'):
                sample_width = float(sample_font.getlength("0"))
            else:
                bbox = sample_font.getbbox("0")
                sample_width = float(max(0, bbox[2] - bbox[0]))
        except Exception:
            return None
        width_pixels = sample_width * target_pixels / sample_size
        return width_pixels * 72.0 / 96.0

    def derive_default_row_height(self, font, stored_height=None, is_custom=False):
        if is_custom and stored_height:
            return float(stored_height)

        resolution = self.font_strategy.resolve(font)
        font_height = self._measure_line_height_points(resolution)
        if font_height > 0:
            return round(font_height + 0.1, 2)

        if stored_height:
            return float(stored_height)
        return 15.0

    def _wrap_paragraph(self, text, resolution, width_points, wrap_text):
        if not wrap_text or width_points is None or width_points <= 0:
            return [text]

        tokens = self._tokenize(text)
        if not tokens:
            return [""]

        lines = []
        current = ""
        current_width = 0.0
        for token in tokens:
            token_width = self._measure_text_width_points(token, resolution)
            if current and current_width + token_width > width_points:
                lines.append(current.rstrip())
                current = token.lstrip()
                current_width = self._measure_text_width_points(current, resolution)
                if current_width > width_points and len(current) > 1:
                    chunked = self._break_long_token(current, resolution, width_points)
                    lines.extend(chunked[:-1])
                    current = chunked[-1]
                    current_width = self._measure_text_width_points(current, resolution)
                continue

            if not current and token_width > width_points and len(token) > 1:
                chunked = self._break_long_token(token, resolution, width_points)
                lines.extend(chunked[:-1])
                current = chunked[-1]
                current_width = self._measure_text_width_points(current, resolution)
                continue

            current += token
            current_width += token_width

        lines.append(current.rstrip())
        return lines or [""]

    def _break_long_token(self, token, resolution, width_points):
        parts = []
        current = ""
        for char in token:
            tentative = current + char
            if current and self._measure_text_width_points(tentative, resolution) > width_points:
                parts.append(current)
                current = char
            else:
                current = tentative
        if current or not parts:
            parts.append(current)
        return parts

    def _split_runs(self, text, resolution):
        if not text:
            return []
        runs = []
        current = text[0]
        current_family = self.font_strategy.choose_family_for_text(text[0], resolution)
        for char in text[1:]:
            family = self.font_strategy.choose_family_for_text(char, resolution)
            if family == current_family:
                current += char
                continue
            runs.append(
                TextRun(
                    text=current,
                    family=current_family,
                    width_points=self._measure_text_width_points(
                        current,
                        resolution._replace(resolved_family=current_family),
                    ),
                )
            )
            current = char
            current_family = family
        runs.append(
            TextRun(
                text=current,
                family=current_family,
                width_points=self._measure_text_width_points(
                    current,
                    resolution._replace(resolved_family=current_family),
                ),
            )
        )
        return runs

    def _tokenize(self, text):
        tokens = []
        for chunk in re.findall(r"\S+\s*|\s+", text):
            if self.font_strategy._contains_cjk(chunk) and " " not in chunk:
                tokens.extend(list(chunk))
            else:
                tokens.append(chunk)
        return tokens

    def _measure_text_width_points(self, text, resolution):
        if not text:
            return 0.0

        font = self._load_font(resolution)
        if font is not None:
            try:
                bbox = font.getbbox(text)
                width_px = max(0, bbox[2] - bbox[0])
                return width_px * 72.0 / 96.0
            except Exception:
                pass

        width_factor = 0.0
        for char in text:
            width_factor += self._heuristic_width_factor(char)
        if resolution.bold:
            width_factor *= 1.04
        if resolution.italic:
            width_factor *= 1.02
        return round(width_factor * resolution.size_points, 4)

    def _measure_line_height_points(self, resolution):
        excel_line_height = resolution.size_points * 1.32
        font = self._load_font(resolution)
        if font is not None:
            try:
                ascent, descent = font.getmetrics()
                metric_height = (ascent + descent) * 72.0 / 96.0
                if resolution.resolved_family.lower() in ('等线', 'dengxian'):
                    # Excel's DengXian grid uses compact external leading.
                    return round(resolution.size_points + 1.38, 4)
                return round(max(metric_height, excel_line_height), 4)
            except Exception:
                pass
        return round(excel_line_height, 4)

    @staticmethod
    @lru_cache(maxsize=256)
    def _load_font(resolution):
        if ImageFont is None or resolution.font_path is None:
            return None
        try:
            return ImageFont.truetype(resolution.font_path, resolution.size_pixels)
        except Exception:
            return None

    @staticmethod
    def _heuristic_width_factor(char):
        if char == "\t":
            return 1.32
        if char.isspace():
            return 0.33
        if unicodedata.category(char).startswith("P"):
            return 0.35
        code = ord(char)
        if (
            0x3400 <= code <= 0x4DBF
            or 0x4E00 <= code <= 0x9FFF
            or 0x3040 <= code <= 0x30FF
            or 0xAC00 <= code <= 0xD7AF
        ):
            return 1.0
        if 0x1F300 <= code <= 0x1FAFF:
            return 1.1
        if char.isdigit():
            return 0.56
        if char.isupper():
            return 0.62
        return 0.55
