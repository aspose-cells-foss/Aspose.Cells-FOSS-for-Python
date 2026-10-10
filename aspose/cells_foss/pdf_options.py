"""Options for workbook PDF export."""


class PdfSaveOptions:
    """Controls worksheet selection and basic PDF rendering behavior."""

    def __init__(self):
        self._worksheet_indices = None
        self._page_size = None
        self.render_gridlines = None
        self.export_hidden_worksheets = False
        self._min_scale = 0.1

    @property
    def worksheet_indices(self):
        """Optional sequence of zero-based worksheet indexes to export."""
        return self._worksheet_indices

    @worksheet_indices.setter
    def worksheet_indices(self, value):
        if value is None:
            self._worksheet_indices = None
            return
        indexes = tuple(int(index) for index in value)
        if any(index < 0 for index in indexes):
            raise ValueError("worksheet indexes must be >= 0")
        self._worksheet_indices = indexes

    @property
    def page_size(self):
        """Optional ``(width, height)`` override in PDF points."""
        return self._page_size

    @page_size.setter
    def page_size(self, value):
        if value is None:
            self._page_size = None
            return
        if len(value) != 2 or float(value[0]) <= 0 or float(value[1]) <= 0:
            raise ValueError("page_size must contain two positive point values")
        self._page_size = (float(value[0]), float(value[1]))

    @property
    def min_scale(self):
        """Smallest allowed page scale as a fraction, defaulting to 0.1."""
        return self._min_scale

    @min_scale.setter
    def min_scale(self, value):
        value = float(value)
        if value <= 0 or value > 1:
            raise ValueError("min_scale must be greater than 0 and at most 1")
        self._min_scale = value
