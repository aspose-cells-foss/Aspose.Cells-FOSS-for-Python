"""Workbook-level PDF export orchestration."""

import os

from .pdf_layout import PdfLayoutEngine
from .pdf_options import PdfSaveOptions
from .pdf_renderer import SkiaPdfRenderer


class PdfExporter:
    """Connects workbook models to prepared pages and a PDF renderer."""

    @classmethod
    def build_pages(cls, workbook, options=None):
        options = cls._options(options)
        pages = []
        for worksheet in cls._select_worksheets(workbook, options):
            pages.extend(PdfLayoutEngine.build_pages(worksheet, options))
        return tuple(pages)

    @classmethod
    def save(cls, workbook, file_path, options=None, renderer=None):
        options = cls._options(options)
        pages = cls.build_pages(workbook, options)
        output_dir = os.path.dirname(os.path.abspath(file_path))
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        (renderer or SkiaPdfRenderer()).render(pages, file_path)

    @staticmethod
    def _options(options):
        if options is None:
            return PdfSaveOptions()
        if not isinstance(options, PdfSaveOptions):
            raise TypeError("PDF options must be a PdfSaveOptions instance")
        return options

    @staticmethod
    def _select_worksheets(workbook, options):
        if options.worksheet_indices is None:
            selected = list(enumerate(workbook.worksheets))
        else:
            selected = []
            for index in options.worksheet_indices:
                if index >= len(workbook.worksheets):
                    raise IndexError(f"worksheet index out of range: {index}")
                selected.append((index, workbook.worksheets[index]))
        return [
            worksheet for _, worksheet in selected
            if options.export_hidden_worksheets or worksheet.visible is True
        ]
