"""
Aspose.Cells for Python - XML Saver Module

This module provides the XMLSaver class which handles saving workbook data to XML format.
The XMLSaver class creates all the necessary XML files for an Excel .xlsx file.

Compatible with Aspose.Cells for .NET API structure.
ECMA-376 Compliant cell value export.
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from .cell_value_handler import CellValueHandler
from .shared_strings import SharedStringTable
from .comment_xml import CommentXMLWriter
from .xml_autofilter_saver import AutoFilterXMLWriter
from .xml_conditional_format_saver import ConditionalFormatXMLWriter
from .xml_properties_saver import WorkbookPropertiesXMLWriter, WorksheetPropertiesXMLWriter
from .xml_hyperlink_handler import HyperlinkXMLSaver, HyperlinkRelationshipWriter
from .xml_datavalidation_saver import DataValidationXmlSaver
from .xml_chart_saver import ChartXmlSaver
from .xml_table_saver import TableXmlSaver
from .xml_sparkline_saver import SparklineXmlSaver
from .chart import ChartType
from .workbook_properties import DefinedName


# Minimal Office Theme XML injected into programmatically-created workbooks that
# contain chartEx charts (treemap / waterfall / sunburst / histogram / funnel / map). The companion
# style{n}.xml files
# reference scheme colours (accent1, bg1, tx1, …) that Excel can only resolve when
# a theme is present in the package.  Without this file Excel repairs the file and
# removes the drawing that references the chart.
_DEFAULT_THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements>'
    '<a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
    '<a:accent3><a:srgbClr val="A9D18E"/></a:accent3>'
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst>'
    '<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
    '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>'
    '</a:effectLst></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme>'
    '</a:themeElements>'
    '<a:objectDefaults/>'
    '<a:extraClrSchemeLst/>'
    '</a:theme>'
)


class XMLSaver:
    """
    Handles saving workbook data to XML format for .xlsx files.
    
    The XMLSaver class is responsible for creating all of the XML files that make up
    an Excel .xlsx file, including content types, relationships, workbook, styles,
    shared strings, and worksheet files.
    
    Examples:
        >>> saver = XMLSaver(workbook)
        >>> saver.save('output.xlsx')
    """
    
    def __init__(self, workbook):
        """
        Initializes a new instance of the XMLSaver class.
        
        Args:
            workbook (Workbook): The workbook to save.
        """
        self._workbook = workbook
        
        # Initialize style dictionaries if they don't exist
        if not hasattr(workbook, '_font_styles'):
            workbook._font_styles = {}
        if not hasattr(workbook, '_fill_styles'):
            workbook._fill_styles = {}
        if not hasattr(workbook, '_border_styles'):
            workbook._border_styles = {}
        if not hasattr(workbook, '_alignment_styles'):
            workbook._alignment_styles = {}
        if not hasattr(workbook, '_cell_styles'):
            workbook._cell_styles = {}
        if not hasattr(workbook, '_num_formats'):
            workbook._num_formats = {}
        
        # Initialize shared string table
        self._shared_string_table = SharedStringTable()

        # Initialize comment writer
        self._comment_writer = CommentXMLWriter()

        # Initialize autofilter writer
        self._autofilter_writer = AutoFilterXMLWriter(self._escape_xml)

        # Initialize conditional formatting writer
        self._cf_writer = ConditionalFormatXMLWriter(self._escape_xml)

        # Initialize hyperlink writer
        self._hyperlink_writer = HyperlinkXMLSaver()

        # Initialize data validation writer
        self._dv_writer = DataValidationXmlSaver()
        self._chart_writer = ChartXmlSaver(self._escape_xml)

        # Initialize properties writers
        self._wb_props_writer = WorkbookPropertiesXMLWriter(self._escape_xml)
        self._ws_props_writer = WorksheetPropertiesXMLWriter(self._escape_xml)

        # Initialize differential formatting (dxf) collection for conditional formatting
        self._dxf_styles = []
        self._sheet_drawing_rel_ids = {}
        self._sheet_drawing_paths = {}  # sheet_num -> actual "xl/drawings/drawingN.xml" path
        self._table_writer = TableXmlSaver(self._escape_xml)
        self._sheet_table_rel_ids = {}     # sheet_num -> {table_idx -> rel_id_str}
        self._table_global_indices = {}    # (sheet_num, table_idx) -> global_table_num
        self._sparkline_writer = SparklineXmlSaver(self._escape_xml)

    @staticmethod
    def _parse_attr_string(attrs_text):
        """Parse a raw XML attribute fragment into an ordered list of (name, value)."""
        if not attrs_text:
            return []
        return re.findall(r'([A-Za-z_][\w:.-]*)="([^"]*)"', attrs_text)

    def _merge_root_attrs(self, base_attrs, source_attrs_text):
        """Merge source root attributes without duplicating native declarations."""
        merged = []
        seen = {}
        ignorable_tokens = []

        for attr_text in base_attrs:
            if '=' not in attr_text:
                continue
            name, value = attr_text.split('=', 1)
            value = value.strip().strip('"')
            seen[name] = len(merged)
            merged.append((name, value))
            if name == 'mc:Ignorable':
                ignorable_tokens.extend(value.split())

        for name, value in self._parse_attr_string(source_attrs_text):
            if name == 'mc:Ignorable':
                for token in value.split():
                    if token not in ignorable_tokens:
                        ignorable_tokens.append(token)
                continue
            if name in seen:
                continue
            seen[name] = len(merged)
            merged.append((name, value))

        if ignorable_tokens:
            ignorable_value = ' '.join(ignorable_tokens)
            if 'mc:Ignorable' in seen:
                merged[seen['mc:Ignorable']] = ('mc:Ignorable', ignorable_value)
            else:
                merged.append(('mc:Ignorable', ignorable_value))

        return [f'{name}="{value}"' for name, value in merged]

    def _can_preserve_calc_chain(self):
        """Return True when every calcChain reference still points to a formula cell."""
        calc_chain_bytes = getattr(self._workbook, '_source_calc_chain_bytes', None)
        if not calc_chain_bytes:
            return False

        try:
            root = ET.fromstring(calc_chain_bytes)
        except ET.ParseError:
            return False

        ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        current_sheet_idx = 1
        worksheets = getattr(self._workbook, 'worksheets', [])

        for c_elem in root.findall('m:c', ns):
            sheet_attr = c_elem.get('i')
            if sheet_attr is not None:
                try:
                    current_sheet_idx = int(sheet_attr)
                except ValueError:
                    return False
            if current_sheet_idx < 1 or current_sheet_idx > len(worksheets):
                return False
            cell_ref = c_elem.get('r')
            if not cell_ref:
                return False
            cell = worksheets[current_sheet_idx - 1].cells.get_all_cells().get(cell_ref)
            if cell is None or not getattr(cell, 'formula', None):
                return False
        return True

    def _compute_dimension_ref(self, worksheet):
        """Compute worksheet dimension from populated cells, merged ranges, and tables."""
        min_row = min_col = None
        max_row = max_col = None

        def include_ref(ref):
            nonlocal min_row, min_col, max_row, max_col
            if not ref:
                return
            if isinstance(ref, tuple) and len(ref) == 2:
                row_idx, col_idx = ref
                try:
                    ref = worksheet.cells.coordinate_to_string(int(row_idx), int(col_idx))
                except (TypeError, ValueError):
                    return
            elif not isinstance(ref, str):
                return
            if ':' in ref:
                start_ref, end_ref = ref.split(':', 1)
            else:
                start_ref = end_ref = ref
            start_row, start_col = worksheet.cells.coordinate_from_string(start_ref)
            end_row, end_col = worksheet.cells.coordinate_from_string(end_ref)
            if min_row is None or start_row < min_row:
                min_row = start_row
            if min_col is None or start_col < min_col:
                min_col = start_col
            if max_row is None or end_row > max_row:
                max_row = end_row
            if max_col is None or end_col > max_col:
                max_col = end_col

        for cell_ref in worksheet.cells.get_all_cells():
            include_ref(cell_ref)
        for merge_ref in getattr(worksheet, '_merged_cells', []) or []:
            include_ref(merge_ref)
        ws_tables = getattr(worksheet, 'tables', None)
        if ws_tables:
            for table in ws_tables:
                include_ref(getattr(table, 'ref', None))

        if min_row is None:
            return getattr(worksheet, '_source_dimension_ref', None)

        start_ref = worksheet.cells.coordinate_to_string(min_row, min_col)
        end_ref = worksheet.cells.coordinate_to_string(max_row, max_col)
        if start_ref == end_ref:
            return start_ref
        return f'{start_ref}:{end_ref}'

    def _max_used_style_index(self):
        """Return the highest style index referenced by workbook cells or columns."""
        max_style = -1

        for worksheet in getattr(self._workbook, 'worksheets', []):
            for cell in worksheet.cells.get_all_cells().values():
                style_idx = getattr(cell, '_source_style_idx', None)
                if style_idx is None:
                    continue
                try:
                    max_style = max(max_style, int(style_idx))
                except (TypeError, ValueError):
                    pass

            for style_idx in (getattr(worksheet, '_column_styles', None) or {}).values():
                try:
                    max_style = max(max_style, int(style_idx))
                except (TypeError, ValueError):
                    pass

        for style_idx in (getattr(self._workbook, '_cell_styles', None) or {}).values():
            try:
                max_style = max(max_style, int(style_idx))
            except (TypeError, ValueError):
                pass

        for style_idx in (getattr(self._workbook, '_cell_xf_by_index', None) or {}).keys():
            try:
                max_style = max(max_style, int(style_idx))
            except (TypeError, ValueError):
                pass

        return max_style

    def _can_preserve_source_styles(self):
        """Return True when the original styles.xml still covers all referenced style ids."""
        source_styles = getattr(self._workbook, '_source_styles_xml', None)
        source_cell_xfs_count = getattr(self._workbook, '_source_cell_xfs_count', 0)
        if source_styles is None or source_cell_xfs_count <= 0:
            return False
        return self._max_used_style_index() < source_cell_xfs_count

    def _register_conditional_format_dxfs(self):
        """
        Registers differential formatting (dxf) styles for all conditional formats.

        This method assigns dxfId to each conditional format that has formatting applied.
        The dxf styles are stored in _dxf_styles for later writing to styles.xml.
        """
        # Preserve DXFs loaded from source styles.xml (e.g. table column dataDxfId
        # references) so existing IDs remain valid after save.
        self._dxf_styles = list(getattr(self._workbook, '_dxf_styles', []) or [])

        # When the source styles.xml is copied through verbatim, only the dxfs it
        # already contains exist in the output, so appending new ones would leave
        # cfRules pointing past the end of <dxfs> and Excel refuses to open the file.
        can_append_dxfs = not self._can_preserve_source_styles()

        for worksheet in self._workbook.worksheets:
            for cf in worksheet.conditional_formats:
                # dxfId the rule was loaded with, if any.
                source_dxf_id = getattr(cf, '_dxf_id', None)
                if source_dxf_id is not None and not 0 <= source_dxf_id < len(self._dxf_styles):
                    source_dxf_id = None

                # Skip rules that don't use dxf (colorScale, dataBar, iconSet)
                if cf._type in ('colorScale', 'dataBar', 'iconSet'):
                    cf._dxf_id = None
                    continue

                # Check if this conditional format has any formatting applied
                has_formatting = self._cf_has_formatting(cf)

                if has_formatting and can_append_dxfs:
                    # Create dxf entry and assign ID
                    dxf_data = self._create_dxf_data(cf)
                    cf._dxf_id = len(self._dxf_styles)
                    self._dxf_styles.append(dxf_data)
                else:
                    # Fall back to the rule's original dxf, which survives the save
                    # either way: it is in the preserved styles.xml, and it is also
                    # re-emitted when styles.xml is regenerated.
                    cf._dxf_id = source_dxf_id

    def _cf_has_formatting(self, cf):
        """Checks if a conditional format has any formatting applied."""
        # Check font
        if cf._font:
            if (cf._font.bold or cf._font.italic or cf._font.underline or
                cf._font.strikethrough or cf._font.color != 'FF000000'):
                return True

        # Check fill
        if cf._fill:
            if cf._fill.pattern_type != 'none' and cf._fill.foreground_color != 'FFFFFFFF':
                return True

        # Check border
        if cf._border:
            if cf._border.line_style != 'none':
                return True

        return False

    def _create_dxf_data(self, cf):
        """Creates a dxf data dictionary from a conditional format."""
        dxf_data = {}

        # Add font data if modified
        if cf._font:
            font_data = {}
            if cf._font.bold:
                font_data['bold'] = True
            if cf._font.italic:
                font_data['italic'] = True
            if cf._font.underline:
                font_data['underline'] = True
            if cf._font.strikethrough:
                font_data['strikethrough'] = True
            if cf._font.color != 'FF000000':
                font_data['color'] = cf._font.color
            if font_data:
                dxf_data['font'] = font_data

        # Add fill data if modified
        if cf._fill and cf._fill.pattern_type != 'none':
            dxf_data['fill'] = {
                'pattern_type': cf._fill.pattern_type,
                'fg_color': cf._fill.foreground_color,
                'bg_color': cf._fill.background_color
            }

        # Add border data if modified
        if cf._border and cf._border.line_style != 'none':
            dxf_data['border'] = {
                'style': cf._border.line_style,
                'color': cf._border.color
            }

        return dxf_data

    def save(self, file_path):
        """
        Saves the workbook to an Excel file (.xlsx format).
        
        Args:
            file_path (str): Path where the Excel file should be saved.
            
        Examples:
            >>> saver = XMLSaver(workbook)
            >>> saver.save('output.xlsx')
        """
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create a ZIP file (XLSX is a ZIP archive)
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            self._sheet_drawing_rel_ids = {}
            self._sheet_drawing_paths = {}
            self._sheet_table_rel_ids = {}
            self._table_global_indices = {}
            written_child_parts = set()
            # Pre-pass: assign globally-unique table numbers across all worksheets.
            _global_table_num = 1
            for _i, _ws in enumerate(self._workbook.worksheets):
                _ws_tables = getattr(_ws, 'tables', None)
                if _ws_tables:
                    for _j in range(_ws_tables.count):
                        self._table_global_indices[(_i + 1, _j)] = _global_table_num
                        _global_table_num += 1
            # Sync worksheet print settings before workbook.xml is generated.
            self._sync_print_settings_to_defined_names()
            # Inject _xlchart.v1.x defined names for new chartEx charts before workbook.xml is written.
            self._inject_chartex_defined_names()
            # Apply compatibility defaults for new chartEx workbooks.
            self._apply_chartex_compat_defaults()

            # Pre-populate drawing paths so _write_content_types can use them.
            # Prefer the original stored path to avoid collisions with chart extra parts.
            for _i, _ws in enumerate(self._workbook.worksheets):
                _has_drawing = (
                    (getattr(_ws, 'charts', None) is not None and _ws.charts.count > 0)
                    or (getattr(_ws, 'pictures', None) is not None and _ws.pictures.count > 0)
                    or (getattr(_ws, 'shapes', None) is not None and _ws.shapes.count > 0)
                    or getattr(_ws, '_source_drawing_xml', None) is not None
                )
                if _has_drawing:
                    _stored = getattr(_ws, '_source_drawing_part_path', None)
                    self._sheet_drawing_paths[_i + 1] = _stored or f'xl/drawings/drawing{_i + 1}.xml'

            # Write [Content_Types].xml
            self._write_content_types(zipf)
            
            # Write _rels/.rels
            self._write_root_relationships(zipf)
            
            # Write xl/_rels/workbook.xml.rels
            self._write_workbook_relationships(zipf)
            
            # Write xl/workbook.xml
            self._write_workbook_xml(zipf)
            
            # Register differential formats (dxf) for conditional formatting
            # This must be done BEFORE writing worksheets so cfRules can reference dxfId
            self._register_conditional_format_dxfs()

            # Process worksheets first to populate shared string table and style collections
            # This must be done BEFORE writing shared strings and styles XML
            next_chart_index = 1
            for i, worksheet in enumerate(self._workbook.worksheets):
                # Chartsheets have no worksheet object model: re-emit their preserved
                # part subtree (chartsheet xml, drawing, charts, ...) verbatim.
                if getattr(worksheet, '_is_chartsheet', False):
                    for part in getattr(worksheet, '_chartsheet_parts', []):
                        path = part.get('part_path')
                        data = part.get('part_bytes')
                        if not path or data is None or path in written_child_parts:
                            continue
                        written_child_parts.add(path)
                        zipf.writestr(path, data)
                    continue

                # Write xl/worksheets/sheet{i+1}.xml
                self._write_worksheet_xml(zipf, worksheet, i+1)

                # Write xl/worksheets/_rels/sheet{i+1}.xml.rels
                self._write_worksheet_relationships(zipf, i+1)

                # Write chart and drawing parts if worksheet has charts
                next_chart_index = self._chart_writer.write_chart_parts(zipf, worksheet, i + 1, next_chart_index)

                # Write xl/tables/tableN.xml parts for this worksheet's tables
                self._write_table_parts_for_worksheet(zipf, worksheet, i + 1)

                # Write xl/comments{i+1}.xml if worksheet has comments
                self._comment_writer.write_comments_xml(zipf, worksheet, i+1)

                # Write xl/drawings/vmlDrawing{i+1}.vml if worksheet has comments
                self._comment_writer.write_vml_drawing_xml(zipf, worksheet, i+1)

                # Write source extra sheet rel parts (vmlDrawing, comments, etc.) for round-trip
                if not self._comment_writer.worksheet_has_comments(worksheet):
                    for r in getattr(worksheet, '_source_extra_sheet_rels', []):
                        if r.get('part_bytes') and r.get('part_path'):
                            zipf.writestr(r['part_path'], r['part_bytes'])
                        if r.get('part_rels_bytes') and r.get('part_rels_path'):
                            zipf.writestr(r['part_rels_path'], r['part_rels_bytes'])
                        # Sheets commonly share one image across their vmlDrawings,
                        # so only the first sheet to claim a media part writes it.
                        for child in r.get('child_parts', []):
                            if not (child.get('part_bytes') and child.get('part_path')):
                                continue
                            if child['part_path'] in written_child_parts:
                                continue
                            written_child_parts.add(child['part_path'])
                            zipf.writestr(child['part_path'], child['part_bytes'])

            # Write extra workbook rel parts (external links, etc.) for round-trip
            for r in getattr(self._workbook, '_source_extra_workbook_rels', []):
                if r.get('part_bytes') and r.get('part_path'):
                    zipf.writestr(r['part_path'], r['part_bytes'])
                if r.get('part_rels_bytes') and r.get('part_rels_path'):
                    zipf.writestr(r['part_rels_path'], r['part_rels_bytes'])
                for child in r.get('child_parts', []):
                    if child.get('part_bytes') and child.get('part_path'):
                        zipf.writestr(child['part_path'], child['part_bytes'])
            calc_chain_bytes = getattr(self._workbook, '_source_calc_chain_bytes', None)
            calc_chain_rel = getattr(self._workbook, '_source_calc_chain_rel', None)
            if self._can_preserve_calc_chain() and calc_chain_bytes and calc_chain_rel and calc_chain_rel.get('part_path'):
                zipf.writestr(calc_chain_rel['part_path'], calc_chain_bytes)

            # Write xl/styles.xml (AFTER processing worksheets to ensure styles are registered)
            self._write_styles_xml(zipf)
            
            # Write xl/sharedStrings.xml (AFTER processing worksheets)
            self._write_shared_strings_xml(zipf)

            # Write xl/theme/theme1.xml when preserving a loaded theme
            self._write_theme_xml(zipf)

            # Write docProps/core.xml (document properties)
            self._write_core_properties_xml(zipf)

            # Write docProps/app.xml (extended properties)
            self._write_app_properties_xml(zipf)

    def _inject_chartex_defined_names(self):
        """
        Pre-computes _xlchart.v1.x defined names for new (non-source) chartEx charts
        (TREEMAP / WATERFALL / BOX_WHISKER / SUNBURST / HISTOGRAM / FUNNEL / MAP) and stores a
        per-series name-map on each
        chart object.

        Excel requires chartEx cx:f elements to reference workbook-level defined names
        rather than raw cell formulas.  This must be called before _write_workbook_xml()
        so the names appear in workbook.xml <definedNames>.

        Names used by source (loaded) charts are preserved; only names from previous
        programmatic injections are removed before re-injecting.
        """
        import re as _re

        defined_names = self._workbook.properties.defined_names

        def _a1_col_to_index(col_letters):
            idx = 0
            for ch in col_letters:
                idx = idx * 26 + (ord(ch) - ord('A') + 1)
            return idx

        def _a1_index_to_col(idx):
            letters = []
            n = int(idx)
            while n > 0:
                n, rem = divmod(n - 1, 26)
                letters.append(chr(ord('A') + rem))
            return ''.join(reversed(letters))

        def _infer_box_tx_formula(values_ref, worksheet_name):
            """
            For box-whisker horizontal series refs like B2:F2, infer tx from the
            immediate left cell (A2), matching Excel-generated chartEx behavior.
            """
            if not values_ref:
                return None
            src = str(values_ref).strip()
            if src.startswith('='):
                src = src[1:].strip()
            m = _re.match(
                r"^(?:(?:'[^']+'|[^'!]+)!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$",
                src,
            )
            if not m:
                return None
            c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
            c2 = c2 or c1
            r2 = r2 or r1
            if r1 != r2:
                return None  # only infer for single-row ranges
            start_col_idx = _a1_col_to_index(c1.upper())
            if start_col_idx <= 1:
                return None
            left_col = _a1_index_to_col(start_col_idx - 1)
            return self._chart_writer._normalize_chart_range_formula(f"{left_col}{r1}", worksheet_name)

        # Collect _xlchart.v1.x names referenced in source chart XMLs so we don't remove them.
        source_chartex_refs = set()
        for ws in self._workbook.worksheets:
            if getattr(ws, 'charts', None) is None:
                continue
            for chart in ws.charts:
                src = getattr(chart, '_source_chart_xml', None)
                if src:
                    text = src if isinstance(src, str) else src.decode('utf-8', errors='replace')
                    for m in _re.findall(r'_xlchart\.v1\.\d+', text):
                        source_chartex_refs.add(m)

        # Remove _xlchart.v1.x names NOT referenced by any source chart (i.e. from a
        # previous programmatic injection).
        defined_names._names = [
            dn for dn in defined_names._names
            if not dn.name.startswith('_xlchart.v1.') or dn.name in source_chartex_refs
        ]

        # Start counter after the highest existing _xlchart.v1.x index.
        counter = 0
        for dn in defined_names._names:
            if dn.name.startswith('_xlchart.v1.'):
                try:
                    n = int(dn.name[len('_xlchart.v1.'):])
                    counter = max(counter, n + 1)
                except (ValueError, IndexError):
                    pass

        for worksheet in self._workbook.worksheets:
            if getattr(worksheet, 'charts', None) is None:
                continue
            for chart in worksheet.charts:
                # Only new (non-source) chartEx charts need injected defined names.
                if getattr(chart, '_source_chart_xml', None) is not None:
                    continue
                if chart.type not in (ChartType.TREEMAP, ChartType.WATERFALL, ChartType.BOX_WHISKER, ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP):
                    continue

                # Build a list of per-series defined-name refs used by chartEx.
                series_name_map = []
                for series in chart.n_series:
                    categories_source = (
                        series.category_data if series.category_data
                        else getattr(chart, 'category_data', None)
                    )
                    cat_name = None
                    if categories_source:
                        cat_formula = self._chart_writer._normalize_chart_range_formula(
                            categories_source, worksheet.name
                        )
                        cat_dn_name = f'_xlchart.v1.{counter}'
                        _cat_dn = DefinedName(cat_dn_name, cat_formula)
                        _cat_dn.hidden = True
                        defined_names.add(_cat_dn)
                        counter += 1
                        cat_name = cat_dn_name

                    val_formula = self._chart_writer._normalize_chart_range_formula(
                        series.values, worksheet.name
                    )
                    val_dn_name = f'_xlchart.v1.{counter}'
                    _val_dn = DefinedName(val_dn_name, val_formula)
                    _val_dn.hidden = True
                    defined_names.add(_val_dn)
                    counter += 1

                    tx_name = None
                    if series.name:
                        # Only emit tx defined names when series.name is a real cell/range ref.
                        # Literal strings (e.g. "Sales") must stay as cx:v only; emitting a
                        # synthetic defined name for literals can trigger Excel repair.
                        tx_source = str(series.name).strip()
                        if tx_source.startswith('='):
                            tx_source = tx_source[1:].strip()

                        # Optional sheet-name prefix must include '!' when present.
                        ref_like = _re.match(
                            r"^(?:(?:'[^']+'|[^'!]+)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$",
                            tx_source,
                        )
                        if ref_like:
                            tx_formula = self._chart_writer._normalize_chart_range_formula(
                                tx_source, worksheet.name
                            )
                            tx_dn_name = f'_xlchart.v1.{counter}'
                            _tx_dn = DefinedName(tx_dn_name, tx_formula)
                            _tx_dn.hidden = True
                            defined_names.add(_tx_dn)
                            counter += 1
                            tx_name = tx_dn_name
                        elif chart.type == ChartType.BOX_WHISKER:
                            inferred_tx_formula = _infer_box_tx_formula(series.values, worksheet.name)
                            if inferred_tx_formula:
                                tx_dn_name = f'_xlchart.v1.{counter}'
                                _tx_dn = DefinedName(tx_dn_name, inferred_tx_formula)
                                _tx_dn.hidden = True
                                defined_names.add(_tx_dn)
                                counter += 1
                                tx_name = tx_dn_name

                    series_name_map.append({'cat': cat_name, 'val': val_dn_name, 'tx': tx_name})

                chart._chartex_series_name_map = series_name_map

    def _has_new_chartex(self):
        """True if workbook contains a programmatically-created chartEx chart."""
        for ws in self._workbook.worksheets:
            for chart in getattr(ws, 'charts', []):
                if chart.type in (ChartType.TREEMAP, ChartType.WATERFALL, ChartType.BOX_WHISKER, ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP):
                    if getattr(chart, '_source_chart_xml', None) is None:
                        return True
        return False

    def _apply_chartex_compat_defaults(self):
        """
        Normalizes workbook/worksheet metadata for new chartEx files.

        Some Excel builds are strict about metadata accompanying chartEx drawings.
        These defaults align programmatic output with files that Excel opens
        without repair dialogs.
        """
        if not self._has_new_chartex():
            return

        props = self._workbook.properties

        # Match modern Excel build metadata seen in valid chartEx packages.
        props.file_version.last_edited = "7"
        props.file_version.lowest_edited = "7"
        props.file_version.rup_build = "29628"
        props.workbook_pr.default_theme_version = 202300
        if props.calculation.calc_id is None or int(props.calculation.calc_id) == 0:
            # For programmatic chartEx compatibility, keep calcId at 0.
            props.calculation.calc_id = 0

        # Ensure sheetFormatPr uses x14ac dyDescent when writing chartEx drawings.
        for ws in self._workbook.worksheets:
            if ws.properties.format.dy_descent is None:
                ws.properties.format.dy_descent = 0.3

    def _sync_print_settings_to_defined_names(self):
        """
        Synchronizes worksheet print areas/titles with workbook defined names.

        Excel stores these as sheet-local built-in names in workbook.xml.
        """
        defined_names = self._workbook.properties.defined_names
        print_names = {'_xlnm.Print_Area', '_xlnm.Print_Titles'}
        existing = [dn for dn in defined_names if dn.name not in print_names]
        defined_names._names = existing

        for sheet_idx, worksheet in enumerate(self._workbook.worksheets):
            print_area = getattr(worksheet, '_print_area', None)
            sheet_name = worksheet.name.replace("'", "''")
            refs = []
            if print_area:
                for token in str(print_area).split(','):
                    part = token.strip().upper()
                    if not part:
                        continue
                    if ':' in part:
                        start_ref, end_ref = part.split(':', 1)
                    else:
                        start_ref, end_ref = part, part
                    abs_start = self._to_absolute_a1_ref(start_ref)
                    abs_end = self._to_absolute_a1_ref(end_ref)
                    abs_ref = f"{abs_start}:{abs_end}" if abs_start != abs_end else abs_start
                    refs.append(f"'{sheet_name}'!{abs_ref}")

            if refs:
                defined_names.add('_xlnm.Print_Area', ','.join(refs), local_sheet_id=sheet_idx)

            title_refs = []
            title_rows = worksheet.page_setup.print_title_rows
            title_columns = worksheet.page_setup.print_title_columns
            if title_rows:
                start, end = title_rows.split(':', 1)
                title_refs.append(f"'{sheet_name}'!${start}:${end}")
            if title_columns:
                start, end = title_columns.split(':', 1)
                title_refs.append(f"'{sheet_name}'!${start}:${end}")
            if title_refs:
                defined_names.add(
                    '_xlnm.Print_Titles',
                    ','.join(title_refs),
                    local_sheet_id=sheet_idx,
                )

    def _to_absolute_a1_ref(self, ref):
        """
        Converts A1 reference (e.g. A1) to absolute form (e.g. $A$1).
        """
        ref = str(ref).replace('$', '').upper()
        col = ''.join(ch for ch in ref if ch.isalpha())
        row = ''.join(ch for ch in ref if ch.isdigit())
        if not col or not row:
            return ref
        return f"${col}${row}"
    
    def _write_content_types(self, zipf):
        """Writes [Content_Types].xml file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        override_seen = set()

        def append_override(part_name, content_type):
            key = (part_name, content_type)
            if key in override_seen:
                return ""
            override_seen.add(key)
            return f'    <Override PartName="{part_name}" ContentType="{self._escape_xml(content_type)}"/>\n'

        default_types = {
            "rels": "application/vnd.openxmlformats-package.relationships+xml",
            "xml": "application/xml",
        }
        for ext, ctype in getattr(self._workbook, '_content_type_defaults', {}).items():
            if ext and ctype:
                default_types[str(ext).lower()] = ctype
        for ext in sorted(default_types.keys()):
            content += f'    <Default Extension="{self._escape_xml(ext)}" ContentType="{self._escape_xml(default_types[ext])}"/>\n'

        # Ensure image defaults are available when workbook has pictures.
        image_ext_content_types = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
            "tif": "image/tiff",
            "tiff": "image/tiff",
            "webp": "image/webp",
        }
        image_exts = set()
        for worksheet in self._workbook.worksheets:
            pics = getattr(worksheet, "pictures", None)
            if pics is None:
                continue
            for pic in pics:
                ext = getattr(pic, "_image_extension", None)
                if ext:
                    image_exts.add(str(ext).lower().lstrip("."))
        for ext in sorted(image_exts):
            if ext in default_types:
                continue
            if ext in image_ext_content_types:
                content += f'    <Default Extension="{self._escape_xml(ext)}" ContentType="{self._escape_xml(image_ext_content_types[ext])}"/>\n'

        content += append_override("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        content += append_override("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        content += append_override("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        
        # Add worksheet content types (chartsheets declare their own below)
        for i, worksheet in enumerate(self._workbook.worksheets):
            if getattr(worksheet, '_is_chartsheet', False):
                continue
            content += append_override(f"/xl/worksheets/sheet{i+1}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

        # Add chartsheet content types and those of their preserved sub-parts
        # (drawing, charts, printerSettings, ...).
        for worksheet in self._workbook.worksheets:
            if not getattr(worksheet, '_is_chartsheet', False):
                continue
            for part_name, content_type in getattr(worksheet, '_chartsheet_content_types', []):
                content += append_override(part_name, content_type)
        
        # Add comments and VML drawing content types for worksheets that have comments
        for i, worksheet in enumerate(self._workbook.worksheets):
            if self._comment_writer.worksheet_has_comments(worksheet):
                content += append_override(f"/xl/comments{i+1}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
                content += append_override(f"/xl/drawings/vmlDrawing{i+1}.vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")
            else:
                # Register content types for comments files preserved via source extra sheet rels
                # (files may use non-standard numbering, e.g. comments1.xml on sheet2)
                for r in getattr(worksheet, '_source_extra_sheet_rels', []):
                    if 'comments' in r.get('rel_type', '').lower() and r.get('part_path'):
                        part_name = f'/{r["part_path"].lstrip("/")}'
                        content += append_override(part_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")

        # Add chart drawing/chart part content types
        chart_index = 1
        for i, worksheet in enumerate(self._workbook.worksheets):
            has_charts = getattr(worksheet, 'charts', None) is not None and worksheet.charts.count > 0
            has_pictures = getattr(worksheet, 'pictures', None) is not None and worksheet.pictures.count > 0
            has_shapes = getattr(worksheet, 'shapes', None) is not None and worksheet.shapes.count > 0
            has_source_drawing = getattr(worksheet, '_source_drawing_xml', None) is not None
            if has_charts or has_pictures or has_shapes or has_source_drawing:
                drawing_path = self._sheet_drawing_paths.get(i + 1, f'xl/drawings/drawing{i+1}.xml')
                content += append_override(f'/{drawing_path}', "application/vnd.openxmlformats-officedocument.drawing+xml")
                for chart in worksheet.charts if has_charts else []:
                    source_part_path = getattr(chart, '_source_chart_part_path', None)
                    source_content_type = getattr(chart, '_source_chart_content_type', None)
                    if getattr(chart, '_source_chart_xml', None) is not None and source_part_path and source_content_type:
                        part_name = f'/{source_part_path.lstrip("/")}'
                        content += append_override(part_name, source_content_type)
                    else:
                        is_chart_ex = bool(getattr(chart, '_source_is_chart_ex', False)) or chart.type in (ChartType.WATERFALL, ChartType.TREEMAP, ChartType.BOX_WHISKER, ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP)
                        if is_chart_ex:
                            content += append_override(
                                f"/xl/charts/chartEx{chart_index}.xml",
                                "application/vnd.ms-office.chartex+xml",
                            )
                            # Excel requires companion style/colors files for every chartEx chart.
                            content += append_override(
                                f"/xl/charts/style{chart_index}.xml",
                                "application/vnd.ms-office.chartstyle+xml",
                            )
                            content += append_override(
                                f"/xl/charts/colors{chart_index}.xml",
                                "application/vnd.ms-office.chartcolorstyle+xml",
                            )
                        else:
                            content += append_override(
                                f"/xl/charts/chart{chart_index}.xml",
                                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                            )
                    chart_index += 1

        # Add extra chart-related content types preserved from loaded files (e.g. style/colors parts)
        for part_name, content_type in self._chart_writer.get_extra_content_type_overrides(self._workbook):
            content += append_override(part_name, content_type)

        # Add content types for extra workbook rel parts (external links, etc.)
        for r in getattr(self._workbook, '_source_extra_workbook_rels', []):
            if r.get('part_path') and r.get('content_type'):
                content += append_override(f'/{r["part_path"]}', r['content_type'])
            for child in r.get('child_parts', []):
                if child.get('part_path') and child.get('content_type'):
                    content += append_override(f'/{child["part_path"]}', child['content_type'])
        calc_chain_rel = getattr(self._workbook, '_source_calc_chain_rel', None)
        preserve_calc_chain = self._can_preserve_calc_chain()
        if preserve_calc_chain and calc_chain_rel and calc_chain_rel.get('part_path'):
            content += append_override(
                f'/{calc_chain_rel["part_path"]}',
                calc_chain_rel.get('content_type') or "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
            )

        # Add table content types (xl/tables/tableN.xml)
        for _i, _ws in enumerate(self._workbook.worksheets):
            _ws_tables = getattr(_ws, 'tables', None)
            if _ws_tables and _ws_tables.count > 0:
                for _j in range(_ws_tables.count):
                    _gidx = self._table_global_indices.get((_i + 1, _j), _j + 1)
                    _table = _ws_tables[_j]
                    _src_path = getattr(_table, '_source_part_path', None)
                    _part_name = f'/{_src_path}' if _src_path else f'/xl/tables/table{_gidx}.xml'
                    content += append_override(_part_name,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")

        # Add theme content type when theme part is available (loaded or default)
        if getattr(self._workbook, '_theme_xml', None) is not None or self._needs_default_theme():
            content += append_override("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")

        # Add docProps content types
        content += append_override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
        content += append_override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")

        content += '</Types>\n'
        zipf.writestr('[Content_Types].xml', content)
    
    def _write_root_relationships(self, zipf):
        """Writes _rels/.rels file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        content += '    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        content += '    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n'
        content += '    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n'
        content += '</Relationships>\n'
        zipf.writestr('_rels/.rels', content)
    
    def _write_workbook_relationships(self, zipf):
        """Writes xl/_rels/workbook.xml.rels file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        
        # Add sheet relationships (worksheets and chartsheets, in workbook order)
        for i, worksheet in enumerate(self._workbook.worksheets):
            if getattr(worksheet, '_is_chartsheet', False):
                target = getattr(worksheet, '_chartsheet_part_path', '')
                # Workbook rels are relative to xl/, so drop the leading "xl/".
                if target.startswith('xl/'):
                    target = target[len('xl/'):]
                content += f'    <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="{self._escape_xml(target)}"/>\n'
            else:
                content += f'    <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>\n'
        
        # Add styles and shared strings relationships
        content += '    <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        content += '    <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
        if getattr(self._workbook, '_theme_xml', None) is not None or self._needs_default_theme():
            content += '    <Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\n'

        # Emit extra workbook rels (external links, etc.) preserved from the source file
        for r in getattr(self._workbook, '_source_extra_workbook_rels', []):
            content += f'    <Relationship Id="{r["rel_id"]}" Type="{r["rel_type"]}" Target="{r["target"]}"/>\n'
        calc_chain_rel = getattr(self._workbook, '_source_calc_chain_rel', None)
        if self._can_preserve_calc_chain() and calc_chain_rel:
            content += (
                f'    <Relationship Id="{calc_chain_rel["rel_id"]}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" '
                f'Target="{calc_chain_rel["target"]}"/>\n'
            )

        content += '</Relationships>\n'
        zipf.writestr('xl/_rels/workbook.xml.rels', content)
    
    def _write_workbook_xml(self, zipf):
        """Writes xl/workbook.xml file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        workbook_attrs = [
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        ]
        source_workbook_extra_attrs = getattr(self._workbook, '_source_workbook_root_extra_attrs', '')
        workbook_attrs = self._merge_root_attrs(workbook_attrs, source_workbook_extra_attrs)
        content += f'<workbook {" ".join(workbook_attrs)}>\n'

        # Write workbook properties
        props = self._workbook.properties

        # File version (ECMA-376 Section 18.2.10)
        content += self._wb_props_writer.format_file_version_xml(props.file_version)

        # Workbook properties (ECMA-376 Section 18.2.13)
        content += self._wb_props_writer.format_workbook_pr_xml(props.workbook_pr)

        source_alt_content = getattr(self._workbook, '_source_workbook_alt_content_xml', None)
        if source_alt_content:
            content += source_alt_content + '\n'

        source_revision_ptr = getattr(self._workbook, '_source_workbook_revision_ptr_xml', None)
        if source_revision_ptr:
            content += source_revision_ptr + '\n'

        # Workbook protection (ECMA-376 Section 18.2.29)
        content += self._wb_props_writer.format_workbook_protection_xml(props.protection)

        # Book views (ECMA-376 Section 18.2.1)
        content += self._wb_props_writer.format_book_views_xml(props.view)

        # Sheets
        content += '    <sheets>\n'

        # Add sheet elements
        for i, worksheet in enumerate(self._workbook.worksheets):
            state_attr = ''
            if worksheet.visible == False:
                state_attr = ' state="hidden"'
            elif worksheet.visible == 'veryHidden':
                state_attr = ' state="veryHidden"'
            content += f'        <sheet name="{self._escape_xml(worksheet.name)}" sheetId="{i+1}"{state_attr} r:id="rId{i+1}"/>\n'

        content += '    </sheets>\n'

        external_link_rels = [
            r for r in getattr(self._workbook, '_source_extra_workbook_rels', [])
            if r.get('rel_type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
        ]
        if external_link_rels:
            content += '    <externalReferences>\n'
            for r in external_link_rels:
                content += f'        <externalReference r:id="{r["rel_id"]}"/>\n'
            content += '    </externalReferences>\n'

        # Defined names (ECMA-376 Section 18.2.6)
        content += self._wb_props_writer.format_defined_names_xml(props.defined_names)

        # Calculation properties (ECMA-376 Section 18.2.2)
        content += self._wb_props_writer.format_calc_pr_xml(props.calculation)

        source_workbook_extlst = getattr(self._workbook, '_source_workbook_extlst_xml', None)
        if source_workbook_extlst:
            content += source_workbook_extlst + '\n'

        content += '</workbook>\n'
        zipf.writestr('xl/workbook.xml', content)
    
    def _write_worksheet_xml(self, zipf, worksheet, sheet_num):
        """
        Writes worksheet XML file with ECMA-376 compliant cell values.
        
        ECMA-376 Part 1, Section 18.3.1.73 specifies that cells must be grouped
        by row elements within the sheetData element.
        
        Args:
            zipf: The ZIP file object to write to.
            worksheet: The worksheet object to save.
            sheet_num: The worksheet number (1-based).
        """
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        worksheet_xml_attrs = [
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        ]
        ws_props = worksheet.properties
        if ws_props.format.dy_descent is not None:
            worksheet_xml_attrs.append('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
            worksheet_xml_attrs.append('xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"')
            worksheet_xml_attrs.append('mc:Ignorable="x14ac"')
        source_root_extra_attrs = getattr(worksheet, '_source_root_extra_attrs', '')
        worksheet_xml_attrs = self._merge_root_attrs(worksheet_xml_attrs, source_root_extra_attrs)
        content += f'<worksheet {" ".join(worksheet_xml_attrs)}>\n'
        drawing_rel_id = None
        next_reserved_rel_id = 1
        if self._comment_writer.worksheet_has_comments(worksheet):
            next_reserved_rel_id = 3
        has_charts = getattr(worksheet, 'charts', None) is not None and worksheet.charts.count > 0
        has_pictures = getattr(worksheet, 'pictures', None) is not None and worksheet.pictures.count > 0
        has_shapes = getattr(worksheet, 'shapes', None) is not None and worksheet.shapes.count > 0
        has_source_drawing = getattr(worksheet, '_source_drawing_xml', None) is not None
        if has_charts or has_pictures or has_shapes or has_source_drawing:
            drawing_rel_id = f"rId{next_reserved_rel_id}"
            self._sheet_drawing_rel_ids[sheet_num] = drawing_rel_id
            next_reserved_rel_id += 1
            # Record actual drawing path: prefer original stored path to avoid
            # conflicts with chart extra parts (e.g. chartUserShapes) that may
            # occupy the sheet-number-based path.
            stored_path = getattr(worksheet, '_source_drawing_part_path', None)
            self._sheet_drawing_paths[sheet_num] = stored_path or f'xl/drawings/drawing{sheet_num}.xml'

        # Assign rel IDs for tables
        has_tables = getattr(worksheet, 'tables', None) is not None and worksheet.tables.count > 0
        table_rel_ids = {}
        if has_tables:
            for _j in range(worksheet.tables.count):
                table_rel_ids[_j] = f"rId{next_reserved_rel_id}"
                next_reserved_rel_id += 1
        self._sheet_table_rel_ids[sheet_num] = table_rel_ids

        source_sheet_pr_xml = getattr(worksheet, '_source_sheet_pr_xml', None)
        if source_sheet_pr_xml:
            content += f'    {source_sheet_pr_xml}\n'

        dimension_ref = self._compute_dimension_ref(worksheet)
        if dimension_ref:
            content += f'    <dimension ref="{self._escape_xml(dimension_ref)}"/>\n'

        # Sheet views (ECMA-376 Section 18.3.1.88)
        active_tab = getattr(getattr(self._workbook, "properties", None).view, "active_tab", 0)
        is_active_sheet = (sheet_num - 1) == int(active_tab)
        content += self._ws_props_writer.format_sheet_views_xml(ws_props, is_selected=is_active_sheet)

        # Sheet format properties (ECMA-376 Section 18.3.1.82)
        content += self._ws_props_writer.format_sheet_format_pr_xml(ws_props.format)

        # Write column widths/hidden columns if configured
        if (
            getattr(worksheet, '_column_widths', None)
            or getattr(worksheet, '_column_styles', None)
            or getattr(worksheet, '_best_fit_columns', None)
            or getattr(worksheet, '_hidden_columns', None)
        ):
            content += self._format_cols_xml(worksheet)

        content += '    <sheetData>\n'
        
        # Get all cells with their references
        cells = worksheet.cells.get_all_cells()
        
        # Sort cells by reference to ensure proper order
        sorted_refs = sorted(cells.keys(), key=self._cell_reference_sort_key)
        
        # Group cells by row (ECMA-376 requirement)
        rows = {}
        for ref in sorted_refs:
            row, col = self._cell_reference_sort_key(ref)
            if row not in rows:
                rows[row] = []
            rows[row].append((ref, cells[ref]))

        # Ensure rows with custom heights are included even if they have no cells
        if getattr(worksheet, '_row_heights', None):
            for row_num in worksheet._row_heights.keys():
                if row_num not in rows:
                    rows[row_num] = []
        if getattr(worksheet, '_hidden_rows', None):
            for row_num in worksheet._hidden_rows:
                if row_num not in rows:
                    rows[row_num] = []
        
        # Write row elements with cells (ECMA-376 compliant structure)
        for row_num in sorted(rows.keys()):
            row_attrs = [f'r="{row_num}"']
            row_height = None
            if getattr(worksheet, '_row_heights', None):
                row_height = worksheet._row_heights.get(row_num)
            if row_height is not None:
                row_attrs.append(f'ht="{row_height}"')
                if row_num in getattr(worksheet, '_custom_height_rows', set()):
                    row_attrs.append('customHeight="1"')
            if getattr(worksheet, '_hidden_rows', None) and row_num in worksheet._hidden_rows:
                row_attrs.append('hidden="1"')
            content += f'        <row {" ".join(row_attrs)}>\n'
            for ref, cell in rows[row_num]:
                content += self._format_cell_xml(ref, cell)
            content += '        </row>\n'
        
        content += '    </sheetData>\n'

        # Sheet protection (ECMA-376 Section 18.3.1.85)
        content += self._ws_props_writer.format_sheet_protection_xml(ws_props.protection)

        # Write auto filter settings (ECMA-376 Section 18.3.1.2)
        # autoFilter must come AFTER sheetData per ECMA-376 schema sequence
        if worksheet.auto_filter.range is not None:
            content += self._autofilter_writer.format_auto_filter_xml(worksheet.auto_filter)

        # Write merged cells (ECMA-376 Section 18.3.1.55)
        # mergeCells is written after sheetData/autoFilter and before conditional formatting.
        content += self._format_merge_cells_xml(worksheet)

        source_phonetic_pr_xml = getattr(worksheet, '_source_phonetic_pr_xml', None)
        if source_phonetic_pr_xml:
            content += f'    {source_phonetic_pr_xml}\n'

        # Write conditional formatting (ECMA-376 Section 18.3.1.18)
        # conditionalFormatting must come AFTER autoFilter per ECMA-376 schema sequence
        if len(worksheet.conditional_formats) > 0:
            content += self._cf_writer.format_conditional_formatting_xml(worksheet.conditional_formats)

        # Write hyperlinks (ECMA-376 Section 18.3.1.48)
        # hyperlinks must come AFTER conditionalFormatting per ECMA-376 schema sequence
        if worksheet.hyperlinks.count > 0:
            # Reset relationship counter for this worksheet
            self._hyperlink_writer.reset_relationship_counter(start_rel_id=next_reserved_rel_id)
            content += self._hyperlink_writer.format_hyperlinks_xml(worksheet)

        # Write data validations (ECMA-376 Section 18.3.1.30, 18.3.1.31)
        # dataValidations must come AFTER hyperlinks per ECMA-376 schema sequence
        if worksheet.data_validations.count > 0:
            content += self._format_data_validations_xml(worksheet.data_validations)

        # Print options (ECMA-376 Section 18.3.1.70)
        content += self._ws_props_writer.format_print_options_xml(ws_props.print_options)

        # Page margins (ECMA-376 Section 18.3.1.62)
        content += self._ws_props_writer.format_page_margins_xml(ws_props.page_margins)

        # Page setup (ECMA-376 Section 18.3.1.63)
        content += self._ws_props_writer.format_page_setup_xml(ws_props.page_setup)

        # Header/footer (ECMA-376 Section 18.3.1.46)
        content += self._ws_props_writer.format_header_footer_xml(ws_props.header_footer)

        # Manual page breaks (ECMA-376 Section 18.3.1.73/18.3.1.17)
        content += self._format_row_breaks_xml(worksheet)
        content += self._format_col_breaks_xml(worksheet)

        # Add drawing reference if worksheet has drawing content (charts/pictures/source drawing).
        if drawing_rel_id:
            content += f'    <drawing r:id="{drawing_rel_id}"/>\n'

        # Add legacy drawing reference if worksheet has comments
        if self._comment_writer.worksheet_has_comments(worksheet):
            content += '    <legacyDrawing r:id="rId1"/>\n'
        else:
            # Preserve legacyDrawing from source extra rels (e.g. vmlDrawing with non-standard numbering)
            extra_rels = getattr(worksheet, '_source_extra_sheet_rels', [])
            vml_rel = next(
                (r for r in extra_rels if 'vmlDrawing' in r.get('rel_type', '')),
                None
            )
            if vml_rel:
                content += f'    <legacyDrawing r:id="{vml_rel["rel_id"]}"/>\n'

        # Add <tableParts> if worksheet has tables (ECMA-376 §18.3.1 — after legacyDrawing)
        _table_rel_ids = self._sheet_table_rel_ids.get(sheet_num, {})
        if _table_rel_ids:
            _tbl_count = len(_table_rel_ids)
            content += f'    <tableParts count="{_tbl_count}">\n'
            for _j in range(_tbl_count):
                content += f'        <tablePart r:id="{_table_rel_ids[_j]}"/>\n'
            content += '    </tableParts>\n'

        # Add sparkline <extLst> (ECMA-376: last element before </worksheet>)
        _sparkline_xml = self._sparkline_writer.format_sparkline_extlst_xml(worksheet)
        if _sparkline_xml:
            content += _sparkline_xml

        content += '</worksheet>\n'
        zipf.writestr(f'xl/worksheets/sheet{sheet_num}.xml', content)

    def _format_cols_xml(self, worksheet):
        """
        Formats column width settings as <cols> XML.
        """
        col_widths = getattr(worksheet, '_column_widths', None) or {}
        col_styles = getattr(worksheet, '_column_styles', None) or {}
        best_fit_cols = getattr(worksheet, '_best_fit_columns', None) or set()
        hidden_cols = getattr(worksheet, '_hidden_columns', None) or set()
        if not col_widths and not hidden_cols and not col_styles and not best_fit_cols:
            return ''

        lines = ['    <cols>']
        all_cols = sorted(
            set(col_widths.keys())
            | set(hidden_cols)
            | set(col_styles.keys())
            | set(best_fit_cols)
        )
        default_width = getattr(getattr(worksheet, 'properties', None).format, 'default_col_width', None)

        def col_signature(col_idx):
            return (
                col_widths.get(col_idx),
                col_styles.get(col_idx),
                col_idx in hidden_cols,
                col_idx in best_fit_cols,
            )

        range_start = all_cols[0]
        range_end = all_cols[0]
        current_sig = col_signature(all_cols[0])

        def emit_range(start_idx, end_idx, sig):
            width, style_idx, hidden, best_fit = sig
            attrs = [f'min="{start_idx}"', f'max="{end_idx}"']
            if width is not None:
                attrs.append(f'width="{width}"')
                if default_width is None or abs(float(width) - float(default_width)) > 1e-9:
                    attrs.append('customWidth="1"')
            if style_idx is not None:
                attrs.append(f'style="{style_idx}"')
            if best_fit:
                attrs.append('bestFit="1"')
            if hidden:
                attrs.append('hidden="1"')
            lines.append(f'        <col {" ".join(attrs)}/>')

        for col_idx in all_cols[1:]:
            sig = col_signature(col_idx)
            if col_idx == range_end + 1 and sig == current_sig:
                range_end = col_idx
                continue
            emit_range(range_start, range_end, current_sig)
            range_start = range_end = col_idx
            current_sig = sig

        emit_range(range_start, range_end, current_sig)
        lines.append('    </cols>\n')
        return '\n'.join(lines)
    
    def _cell_reference_sort_key(self, ref):
        """
        Creates a sort key for cell references.
        
        Args:
            ref (str): Cell reference (e.g., "A1", "B3")
            
        Returns:
            tuple: (row, column) for sorting
        """
        from .cells import Cells
        row, col = Cells.coordinate_from_string(ref)
        return (row, col)

    def _format_row_breaks_xml(self, worksheet):
        """
        Formats manual horizontal page breaks as <rowBreaks> XML.

        Args:
            worksheet: The worksheet containing manual row breaks.

        Returns:
            str: XML string for rowBreaks.
        """
        breaks = sorted(getattr(worksheet, '_horizontal_page_breaks', set()))
        if not breaks:
            return ''

        lines = [f'    <rowBreaks count="{len(breaks)}" manualBreakCount="{len(breaks)}">']
        for row_idx in breaks:
            lines.append(f'        <brk id="{int(row_idx)}" max="16383" man="1"/>')
        lines.append('    </rowBreaks>\n')
        return '\n'.join(lines)

    def _format_col_breaks_xml(self, worksheet):
        """
        Formats manual vertical page breaks as <colBreaks> XML.

        Args:
            worksheet: The worksheet containing manual column breaks.

        Returns:
            str: XML string for colBreaks.
        """
        breaks = sorted(getattr(worksheet, '_vertical_page_breaks', set()))
        if not breaks:
            return ''

        lines = [f'    <colBreaks count="{len(breaks)}" manualBreakCount="{len(breaks)}">']
        for col_idx in breaks:
            lines.append(f'        <brk id="{int(col_idx)}" max="1048575" man="1"/>')
        lines.append('    </colBreaks>\n')
        return '\n'.join(lines)

    def _format_merge_cells_xml(self, worksheet):
        """
        Formats merged cell ranges as <mergeCells> XML.

        Args:
            worksheet: The worksheet containing merged ranges.

        Returns:
            str: XML string for mergeCells.
        """
        merged = list(getattr(worksheet, '_merged_cells', []) or [])
        if not merged:
            return ''

        lines = [f'    <mergeCells count="{len(merged)}">']
        for merge_ref in merged:
            lines.append(f'        <mergeCell ref="{self._escape_xml(str(merge_ref).upper())}"/>')
        lines.append('    </mergeCells>\n')
        return '\n'.join(lines)

    def _format_data_validations_xml(self, validations):
        """
        Formats data validations as XML according to ECMA-376 specification.

        Args:
            validations (DataValidationCollection): The data validations collection.

        Returns:
            str: XML string for data validations.
        """
        from .data_validation import (
            DataValidationType, DataValidationOperator,
            DataValidationAlertStyle, DataValidationImeMode
        )

        # Mapping from enum values to XML attribute values
        type_map = {
            DataValidationType.NONE: 'none',
            DataValidationType.WHOLE_NUMBER: 'whole',
            DataValidationType.DECIMAL: 'decimal',
            DataValidationType.LIST: 'list',
            DataValidationType.DATE: 'date',
            DataValidationType.TIME: 'time',
            DataValidationType.TEXT_LENGTH: 'textLength',
            DataValidationType.CUSTOM: 'custom',
        }

        operator_map = {
            DataValidationOperator.BETWEEN: 'between',
            DataValidationOperator.NOT_BETWEEN: 'notBetween',
            DataValidationOperator.EQUAL: 'equal',
            DataValidationOperator.NOT_EQUAL: 'notEqual',
            DataValidationOperator.GREATER_THAN: 'greaterThan',
            DataValidationOperator.LESS_THAN: 'lessThan',
            DataValidationOperator.GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',
            DataValidationOperator.LESS_THAN_OR_EQUAL: 'lessThanOrEqual',
        }

        alert_map = {
            DataValidationAlertStyle.STOP: 'stop',
            DataValidationAlertStyle.WARNING: 'warning',
            DataValidationAlertStyle.INFORMATION: 'information',
        }

        ime_map = {
            DataValidationImeMode.NO_CONTROL: 'noControl',
            DataValidationImeMode.OFF: 'off',
            DataValidationImeMode.ON: 'on',
            DataValidationImeMode.DISABLED: 'disabled',
            DataValidationImeMode.HIRAGANA: 'hiragana',
            DataValidationImeMode.FULL_KATAKANA: 'fullKatakana',
            DataValidationImeMode.HALF_KATAKANA: 'halfKatakana',
            DataValidationImeMode.FULL_ALPHA: 'fullAlpha',
            DataValidationImeMode.HALF_ALPHA: 'halfAlpha',
            DataValidationImeMode.FULL_HANGUL: 'fullHangul',
            DataValidationImeMode.HALF_HANGUL: 'halfHangul',
        }

        xml = f'<dataValidations count="{validations.count}"'

        if validations.disable_prompts:
            xml += ' disablePrompts="1"'
        if validations.x_window is not None:
            xml += f' xWindow="{validations.x_window}"'
        if validations.y_window is not None:
            xml += f' yWindow="{validations.y_window}"'

        xml += '>'

        for dv in validations:
            xml += '<dataValidation'

            # Required attribute: sqref
            if dv.sqref:
                xml += f' sqref="{self._escape_xml(dv.sqref)}"'

            # Type attribute (only if not default 'none')
            if dv.type != DataValidationType.NONE:
                xml += f' type="{type_map.get(dv.type, "none")}"'

            # Operator attribute (for types that use operators, only if not default)
            if dv.type in (DataValidationType.WHOLE_NUMBER, DataValidationType.DECIMAL,
                           DataValidationType.DATE, DataValidationType.TIME,
                           DataValidationType.TEXT_LENGTH):
                if dv.operator != DataValidationOperator.BETWEEN:
                    xml += f' operator="{operator_map.get(dv.operator, "between")}"'

            # Error style (only if not default 'stop')
            if dv.alert_style != DataValidationAlertStyle.STOP:
                xml += f' errorStyle="{alert_map.get(dv.alert_style, "stop")}"'

            # IME mode (only if not default 'noControl')
            if dv.ime_mode != DataValidationImeMode.NO_CONTROL:
                xml += f' imeMode="{ime_map.get(dv.ime_mode, "noControl")}"'

            # Boolean attributes
            if dv.allow_blank:
                xml += ' allowBlank="1"'

            # Note: showDropDown="1" means HIDE dropdown (counterintuitive ECMA-376 naming)
            if not dv.show_dropdown:
                xml += ' showDropDown="1"'

            if dv.show_input_message:
                xml += ' showInputMessage="1"'

            if dv.show_error_message:
                xml += ' showErrorMessage="1"'

            # String attributes
            if dv.error_title:
                xml += f' errorTitle="{self._escape_xml(dv.error_title)}"'

            if dv.error_message:
                xml += f' error="{self._escape_xml(dv.error_message)}"'

            if dv.input_title:
                xml += f' promptTitle="{self._escape_xml(dv.input_title)}"'

            if dv.input_message:
                xml += f' prompt="{self._escape_xml(dv.input_message)}"'

            xml += '>'

            # Formula elements
            if dv.formula1 is not None:
                xml += f'<formula1>{self._escape_xml(dv.formula1)}</formula1>'

            if dv.formula2 is not None:
                xml += f'<formula2>{self._escape_xml(dv.formula2)}</formula2>'

            xml += '</dataValidation>'

        xml += '</dataValidations>'
        return xml

    def _format_cell_xml(self, ref, cell):
        """
        Formats a cell as XML according to ECMA-376 specification.
        
        Args:
            ref (str): Cell reference (e.g., "A1")
            cell (Cell): The cell object
            
        Returns:
            str: XML representation of the cell
        """
        # Preserve original style index for loaded cells to avoid xf index drift.
        source_style_idx = getattr(cell, '_source_style_idx', None)
        if source_style_idx is not None:
            style_idx = int(source_style_idx)
        else:
            style_idx = self.get_or_create_cell_style(cell)
        
        # Format value using CellValueHandler for ECMA-376 compliance
        value_str, cell_type = CellValueHandler.format_value_for_xml(cell.value)
        
        # Handle shared strings
        if cell_type == CellValueHandler.TYPE_SHARED_STRING and value_str is not None:
            # Add to shared string table and get index
            shared_string_index = self._shared_string_table.add_string(value_str)
            value_str = str(shared_string_index)
        
        # Build cell XML
        # ECMA-376: cell element with r (reference), s (style), and t (type) attributes
        # ECMA-376: formula (<f>) must come before value (<v>)
        
        if style_idx > 0 and cell_type is not None:
            xml = f'        <c r="{ref}" s="{style_idx}" t="{cell_type}">\n'
        elif style_idx > 0:
            xml = f'        <c r="{ref}" s="{style_idx}">\n'
        elif cell_type is not None:
            xml = f'        <c r="{ref}" t="{cell_type}">\n'
        else:
            xml = f'        <c r="{ref}">\n'
        
        # Add formula if present (ECMA-376: formula must come before value)
        if cell.formula:
            # Remove leading '=' from formula for ECMA-376 compliance
            formula_text = cell.formula.lstrip('=')
            escaped_formula = self._escape_xml(formula_text)
            xml += f'            <f>{escaped_formula}</f>\n'
        
        # Add value if present
        if value_str is not None:
            escaped_value = self._escape_xml(value_str)
            xml += f'            <v>{escaped_value}</v>\n'
        
        xml += '        </c>\n'
        
        return xml
    
    def _escape_xml(self, text):
        """
        Escapes special characters for XML according to ECMA-376.
        
        ECMA-376 Part 1, Section 3.2.20 specifies that the following characters
        must be escaped in XML content:
        - & (ampersand) -> &amp;
        - < (less than) -> &lt;
        - > (greater than) -> &gt;
        - " (double quote) -> &quot;
        - ' (apostrophe/single quote) -> &apos;
        
        Note: The > character only needs to be escaped when it appears in the
        sequence ]]> to avoid confusion with CDATA section end markers.
        However, it's good practice to always escape it for consistency.
        
        Args:
            text (str): The text to escape
            
        Returns:
            str: The escaped text, or None if input is None
        """
        if text is None:
            return None
        
        # Ensure we're working with a string
        if not isinstance(text, str):
            text = str(text)
        
        # Escape characters in the correct order to avoid double-escaping
        # Order matters: & must be escaped first to avoid escaping the & in other entities
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        
        return text

    def _write_worksheet_relationships(self, zipf, sheet_num):
        """Writes xl/worksheets/_rels/sheet{sheet_num}.xml.rels file."""
        worksheet = self._workbook.worksheets[sheet_num - 1]

        # Collect existing relationships (comments, VML, drawing)
        existing_rels = []
        if self._comment_writer.worksheet_has_comments(worksheet):
            existing_rels.append(('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                                 f'../drawings/vmlDrawing{sheet_num}.vml', None))
            existing_rels.append(('rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                                 f'../comments{sheet_num}.xml', None))
        else:
            # Preserve source extra sheet rels (vmlDrawing, comments, etc.) for round-trip fidelity
            extra_rels = getattr(worksheet, '_source_extra_sheet_rels', [])
            for r in extra_rels:
                existing_rels.append((
                    r['rel_id'],
                    r['rel_type'],
                    r['target'],
                    r['target_mode'] if r['target_mode'] else None,
                ))
        drawing_rel_id = self._sheet_drawing_rel_ids.get(sheet_num)
        if drawing_rel_id:
            drawing_path = self._sheet_drawing_paths.get(sheet_num, f'xl/drawings/drawing{sheet_num}.xml')
            drawing_target = '../' + drawing_path[len('xl/'):]  # "xl/drawings/drawingN.xml" -> "../drawings/drawingN.xml"
            existing_rels.append((drawing_rel_id, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                                 drawing_target, None))

        # Add table relationship entries
        _tbl_rel_ids = self._sheet_table_rel_ids.get(sheet_num, {})
        _ws_tables = getattr(worksheet, 'tables', None)
        if _tbl_rel_ids and _ws_tables:
            for _j in range(_ws_tables.count):
                _gidx = self._table_global_indices.get((sheet_num, _j), _j + 1)
                _rel_id = _tbl_rel_ids.get(_j)
                if _rel_id:
                    _src_path = getattr(_ws_tables[_j], '_source_part_path', None)
                    _target_path = _src_path or f'xl/tables/table{_gidx}.xml'
                    _target = '../' + _target_path[len('xl/'):]
                    existing_rels.append((_rel_id,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
                        _target, None))

        # Get hyperlink relationships
        hyperlink_rels = self._hyperlink_writer.get_hyperlink_relationships(worksheet)

        # Only write relationships file if there are relationships to write
        if existing_rels or hyperlink_rels:
            content = HyperlinkRelationshipWriter.format_relationships_xml(hyperlink_rels, existing_rels)
            zipf.writestr(f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels', content)

    def _write_styles_xml(self, zipf):
        """Writes xl/styles.xml file."""
        if self._can_preserve_source_styles():
            zipf.writestr('xl/styles.xml', self._workbook._source_styles_xml)
            return

        # Register default styles
        self.register_default_styles()
        
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        
        # Write number formats
        custom_num_fmts = {k: v for k, v in self._workbook._num_formats.items() if k >= 164}
        content += f'    <numFmts count="{len(custom_num_fmts)}">\n'
        for num_fmt_id, format_code in sorted(custom_num_fmts.items()):
            escaped_code = self._escape_xml(format_code)
            content += f'        <numFmt numFmtId="{num_fmt_id}" formatCode="{escaped_code}"/>\n'
        content += '    </numFmts>\n'
        
        # Write fonts
        content += f'    <fonts count="{len(self._workbook._font_styles)}">\n'
        for font_idx in sorted(self._workbook._font_styles.keys()):
            font_data = self._workbook._font_styles[font_idx]
            content += self._format_font_xml(font_data)
        content += '    </fonts>\n'
        
        # Write fills
        content += f'    <fills count="{len(self._workbook._fill_styles)}">\n'
        for fill_idx in sorted(self._workbook._fill_styles.keys()):
            fill_data = self._workbook._fill_styles[fill_idx]
            content += self._format_fill_xml(fill_data)
        content += '    </fills>\n'
        
        # Write borders
        content += f'    <borders count="{len(self._workbook._border_styles)}">\n'
        for border_idx in sorted(self._workbook._border_styles.keys()):
            border_data = self._workbook._border_styles[border_idx]
            content += self._format_border_xml(border_data)
        content += '    </borders>\n'

        # Base style XF table required by cellXfs xfId references.
        content += '    <cellStyleXfs count="1">\n'
        content += '        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>\n'
        content += '    </cellStyleXfs>\n'
        
        # Write cellXfs.
        # Preserve original xf ordering/duplicates when loaded from existing files.
        if hasattr(self._workbook, '_cell_xf_by_index') and self._workbook._cell_xf_by_index:
            max_xf_idx = max(self._workbook._cell_xf_by_index.keys())
            if max_xf_idx < 0:
                max_xf_idx = 0
            content += f'    <cellXfs count="{max_xf_idx + 1}">\n'
            for xf_idx in range(0, max_xf_idx + 1):
                if xf_idx == 0 and xf_idx not in self._workbook._cell_xf_by_index:
                    content += '        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
                    continue
                cell_style_key = self._workbook._cell_xf_by_index.get(xf_idx, (0, 0, 0, 0, 0, 0))
                font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx = cell_style_key
                apply_number_format = f' applyNumberFormat="1"' if num_fmt_idx != 0 else ''
                apply_protection = f' applyProtection="1"' if protection_idx != 0 else ''
                content += (
                    f'        <xf numFmtId="{num_fmt_idx}" fontId="{font_idx}" fillId="{fill_idx}" '
                    f'borderId="{border_idx}" xfId="0"{apply_number_format}{apply_protection}>\n'
                )
                if alignment_idx > 0 and alignment_idx in self._workbook._alignment_styles:
                    align_data = self._workbook._alignment_styles[alignment_idx]
                    content += self._format_alignment_xml(align_data)
                if protection_idx > 0 and protection_idx in self._workbook._protection_styles:
                    prot_data = self._workbook._protection_styles[protection_idx]
                    content += self._format_protection_xml(prot_data)
                content += '        </xf>\n'
            content += '    </cellXfs>\n'
        else:
            content += f'    <cellXfs count="{len(self._workbook._cell_styles) + 1}">\n'
            content += '        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
            for cell_style_key, xf_idx in sorted(self._workbook._cell_styles.items(), key=lambda x: x[1]):
                font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx = cell_style_key
                apply_number_format = f' applyNumberFormat="1"' if num_fmt_idx != 0 else ''
                apply_protection = f' applyProtection="1"' if protection_idx != 0 else ''
                content += (
                    f'        <xf numFmtId="{num_fmt_idx}" fontId="{font_idx}" fillId="{fill_idx}" '
                    f'borderId="{border_idx}" xfId="0"{apply_number_format}{apply_protection}>\n'
                )
                if alignment_idx > 0:
                    align_data = self._workbook._alignment_styles[alignment_idx]
                    content += self._format_alignment_xml(align_data)
                if protection_idx > 0:
                    prot_data = self._workbook._protection_styles[protection_idx]
                    content += self._format_protection_xml(prot_data)
                content += '        </xf>\n'
            content += '    </cellXfs>\n'

        # At minimum define the built-in Normal style.
        content += '    <cellStyles count="1">\n'
        content += '        <cellStyle name="Normal" xfId="0" builtinId="0"/>\n'
        content += '    </cellStyles>\n'

        # Write differential formatting (dxf) for conditional formatting
        if len(self._dxf_styles) > 0:
            content += f'    <dxfs count="{len(self._dxf_styles)}">\n'
            for dxf_data in self._dxf_styles:
                content += self._format_dxf_xml(dxf_data)
            content += '    </dxfs>\n'
        else:
            content += '    <dxfs count="0"/>\n'

        content += '</styleSheet>\n'
        zipf.writestr('xl/styles.xml', content)

    def _format_dxf_xml(self, dxf_data):
        """
        Formats differential formatting (dxf) as XML for conditional formatting.

        Args:
            dxf_data (dict): Dictionary containing font, fill, and border data.

        Returns:
            str: XML representation of the dxf element.
        """
        xml = '        <dxf>\n'

        # Add font if present
        if 'font' in dxf_data:
            font = dxf_data['font']
            xml += '            <font>\n'
            if font.get('bold'):
                xml += '                <b val="1"/>\n'
            if font.get('italic'):
                xml += '                <i val="1"/>\n'
            if font.get('underline'):
                xml += '                <u/>\n'
            if font.get('strikethrough'):
                xml += '                <strike/>\n'
            if font.get('color'):
                xml += f'                <color rgb="{font["color"]}"/>\n'
            xml += '            </font>\n'

        # Add fill if present
        if 'fill' in dxf_data:
            fill = dxf_data['fill']
            xml += '            <fill>\n'
            xml += f'                <patternFill patternType="{fill.get("pattern_type", "solid")}">\n'
            if fill.get('fg_color'):
                xml += f'                    <fgColor rgb="{fill["fg_color"]}"/>\n'
            if fill.get('bg_color'):
                xml += f'                    <bgColor rgb="{fill["bg_color"]}"/>\n'
            xml += '                </patternFill>\n'
            xml += '            </fill>\n'

        # Add border if present
        if 'border' in dxf_data:
            border = dxf_data['border']
            style = border.get('style', 'thin')
            color = border.get('color', 'FF000000')
            xml += '            <border>\n'
            xml += f'                <left style="{style}"><color rgb="{color}"/></left>\n'
            xml += f'                <right style="{style}"><color rgb="{color}"/></right>\n'
            xml += f'                <top style="{style}"><color rgb="{color}"/></top>\n'
            xml += f'                <bottom style="{style}"><color rgb="{color}"/></bottom>\n'
            xml += '            </border>\n'

        xml += '        </dxf>\n'
        return xml

    def _format_font_xml(self, font_data):
        """Formats font data as XML."""
        xml = '        <font>\n'
        if font_data.get('bold'):
            xml += '            <b/>\n'
        if font_data.get('italic'):
            xml += '            <i/>\n'
        if font_data.get('underline'):
            underline_type = font_data.get('underline_type', 'single')
            if underline_type in (None, 'single'):
                xml += '            <u/>\n'
            else:
                xml += f'            <u val="{underline_type}"/>\n'
        if font_data.get('strikethrough'):
            xml += '            <strike/>\n'
        vertical_alignment = font_data.get(
            'vertical_alignment', 'baseline'
        )
        if vertical_alignment in ('superscript', 'subscript'):
            xml += (
                f'            <vertAlign val="{vertical_alignment}"/>\n'
            )
        xml += f'            <sz val="{font_data["size"]}"/>\n'
        color_type = font_data.get('color_type', 'rgb')
        color_value = font_data.get('color_value', font_data.get('color', 'FF000000'))
        color_tint = font_data.get('color_tint')
        if color_type == 'theme':
            tint_attr = f' tint="{color_tint}"' if color_tint is not None else ''
            xml += f'            <color theme="{color_value}"{tint_attr}/>\n'
        elif color_type == 'indexed':
            tint_attr = f' tint="{color_tint}"' if color_tint is not None else ''
            xml += f'            <color indexed="{color_value}"{tint_attr}/>\n'
        elif color_type == 'auto':
            xml += f'            <color auto="{color_value}"/>\n'
        elif color_type is None:
            pass
        else:
            xml += f'            <color rgb="{font_data.get("color", "FF000000")}"/>\n'
        xml += f'            <name val="{font_data["name"]}"/>\n'
        if font_data.get('family') is not None:
            xml += f'            <family val="{font_data["family"]}"/>\n'
        if font_data.get('charset') is not None:
            xml += f'            <charset val="{font_data["charset"]}"/>\n'
        if font_data.get('scheme') is not None:
            xml += f'            <scheme val="{font_data["scheme"]}"/>\n'
        xml += '        </font>\n'
        return xml
    
    def _format_fill_xml(self, fill_data):
        """Formats fill data as XML."""
        pattern_type = fill_data['pattern_type']
        xml = f'        <fill>\n            <patternFill patternType="{pattern_type}"'
        if pattern_type != 'none' and pattern_type != 'gray125':
            xml += '>\n'
            xml += f'                {self._format_fill_color_tag("fgColor", fill_data, "fg")}\n'
            xml += f'                {self._format_fill_color_tag("bgColor", fill_data, "bg")}\n'
            xml += '            </patternFill>\n'
        else:
            xml += '/>\n'
        xml += '        </fill>\n'
        return xml

    def _format_fill_color_tag(self, tag_name, fill_data, prefix):
        color_type = fill_data.get(f'{prefix}_color_type', 'rgb')
        color_value = fill_data.get(f'{prefix}_color_value', fill_data.get(f'{prefix}_color', 'FFFFFFFF'))
        tint = fill_data.get(f'{prefix}_color_tint')
        tint_attr = f' tint="{tint}"' if tint is not None else ''
        if color_type == 'theme':
            return f'<{tag_name} theme="{color_value}"{tint_attr}/>'
        if color_type == 'indexed':
            return f'<{tag_name} indexed="{color_value}"{tint_attr}/>'
        if color_type == 'auto':
            return f'<{tag_name} auto="{color_value}"{tint_attr}/>'
        return f'<{tag_name} rgb="{fill_data.get(f"{prefix}_color", "FFFFFFFF")}"/>'
    
    def _format_border_xml(self, border_data):
        """Formats border data as XML."""
        xml = '        <border>\n'
        for side in ['left', 'right', 'top', 'bottom']:
            side_data = border_data[side]
            if side_data['style'] != 'none':
                xml += f'            <{side} style="{side_data["style"]}">\n'
                if side_data.get('automatic_color'):
                    xml += '                <color auto="1"/>\n'
                else:
                    xml += (
                        f'                <color rgb="{side_data["color"]}"/>\n'
                    )
                xml += f'            </{side}>\n'
        xml += '        </border>\n'
        return xml
    
    def _format_alignment_xml(self, align_data):
        """Formats alignment data as XML."""
        attrs = []
        if align_data['horizontal'] != 'general':
            attrs.append(f'horizontal="{align_data["horizontal"]}"')
        if align_data['vertical'] != 'bottom':
            attrs.append(f'vertical="{align_data["vertical"]}"')
        if align_data['wrap_text']:
            attrs.append('wrapText="1"')
        if align_data['indent'] != 0:
            attrs.append(f'indent="{align_data["indent"]}"')
        if align_data['text_rotation'] != 0:
            attrs.append(f'textRotation="{align_data["text_rotation"]}"')
        if align_data['shrink_to_fit']:
            attrs.append('shrinkToFit="1"')
        if align_data['reading_order'] != 0:
            attrs.append(f'readingOrder="{align_data["reading_order"]}"')
        if align_data['relative_indent'] != 0:
            attrs.append(f'relativeIndent="{align_data["relative_indent"]}"')
        
        if attrs:
            xml = '            <alignment ' + ' '.join(attrs) + '/>\n'
        else:
            xml = '            <alignment/>\n'
        return xml

    def _format_protection_xml(self, prot_data):
        """
        Formats protection data as XML.

        ECMA-376 Section: 18.8.33 (protection element)
        Default values: locked="1", hidden="0"
        """
        attrs = []
        # Only write non-default values
        if not prot_data['locked']:  # Default is True (1)
            attrs.append('locked="0"')
        if prot_data['hidden']:  # Default is False (0)
            attrs.append('hidden="1"')

        if attrs:
            xml = '            <protection ' + ' '.join(attrs) + '/>\n'
        else:
            # If both are default, we can omit the element entirely
            xml = ''
        return xml

    def _write_shared_strings_xml(self, zipf):
        """Writes xl/sharedStrings.xml file."""
        content = self._shared_string_table.to_xml()
        zipf.writestr('xl/sharedStrings.xml', content)

    def _needs_default_theme(self):
        """Returns True when the workbook needs a default theme injected.

        A default theme is needed when there is no loaded theme (_theme_xml is None)
        but the workbook contains at least one new (programmatically-created) chartEx
        chart (TREEMAP, WATERFALL, BOX_WHISKER, SUNBURST, HISTOGRAM, FUNNEL, or MAP). Those charts write a companion style{n}.xml
        that references scheme colours which Excel can only resolve with a theme.
        Without the theme Excel repairs the file and removes the drawing.
        """
        if getattr(self._workbook, '_theme_xml', None) is not None:
            return False
        for ws in self._workbook.worksheets:
            for chart in getattr(ws, 'charts', []):
                if chart.type in (ChartType.TREEMAP, ChartType.WATERFALL, ChartType.BOX_WHISKER, ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP):
                    if getattr(chart, '_source_chart_xml', None) is None:
                        return True
        return False

    def _write_table_parts_for_worksheet(self, zipf, worksheet, sheet_num):
        """Write xl/tables/tableN.xml files for all tables in one worksheet."""
        ws_tables = getattr(worksheet, 'tables', None)
        if not ws_tables or ws_tables.count == 0:
            return
        for j, table in enumerate(ws_tables):
            gidx = self._table_global_indices.get((sheet_num, j), j + 1)
            if getattr(table, '_source_table_xml', None) is not None:
                path = getattr(table, '_source_part_path', None) or f'xl/tables/table{gidx}.xml'
                table_xml = self._sync_table_column_names(table._source_table_xml, worksheet)
                zipf.writestr(path, table_xml)
            else:
                path = f'xl/tables/table{gidx}.xml'
                xml_str = self._table_writer.format_table_xml(table, gidx)
                zipf.writestr(path, xml_str.encode('utf-8'))

    @staticmethod
    def _col_index_to_letter(idx):
        """0-based column index -> column letter(s), e.g. 0 -> 'A', 27 -> 'AB'."""
        letters = ''
        idx += 1
        while idx > 0:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(ord('A') + rem) + letters
        return letters

    def _sync_table_column_names(self, table_xml, worksheet):
        """
        Rewrites a preserved table's <tableColumn> names to match its current
        header-row cells.

        The table part is round-tripped as source bytes, so editing a header cell
        (as the round-trip test does to A1) would otherwise leave the column name
        out of sync with the header text. Excel requires them to match and rejects
        the file otherwise. When the header cell is unchanged this is a no-op, so
        untouched tables serialize byte-for-byte as before.
        """
        from .table import _parse_cell_range

        try:
            text = table_xml.decode('utf-8')
        except (UnicodeDecodeError, AttributeError):
            return table_xml

        table_tag = re.search(r'<table\b[^>]*>', text)
        if not table_tag:
            return table_xml
        tag = table_tag.group(0)

        ref_m = re.search(r'\bref="([^"]+)"', tag)
        if not ref_m:
            return table_xml
        rng = _parse_cell_range(ref_m.group(1))
        if not rng:
            return table_xml
        start_row, start_col, _end_row, end_col = rng

        hrc_m = re.search(r'\bheaderRowCount="([^"]+)"', tag)
        try:
            header_row_count = int(hrc_m.group(1)) if hrc_m else 1
        except ValueError:
            header_row_count = 1
        if header_row_count == 0:
            # Header-less table: column names are independent of any cell.
            return table_xml

        cells_map = getattr(worksheet.cells, '_cells', {})
        col_iter = iter(range(start_col, end_col + 1))

        def replace_column(mo):
            col_tag = mo.group(0)
            try:
                col = next(col_iter)
            except StopIteration:
                return col_tag
            ref = f'{self._col_index_to_letter(col)}{start_row + 1}'
            cell = cells_map.get(ref)
            value = getattr(cell, 'value', None) if cell is not None else None
            # Only sync to a real, non-empty string; Excel forbids blank names.
            if not isinstance(value, str) or value == '':
                return col_tag
            new_name = self._escape_xml(value)
            if re.search(r'\bname="[^"]*"', col_tag):
                return re.sub(r'\bname="[^"]*"', f'name="{new_name}"', col_tag, count=1)
            return col_tag

        new_text = re.sub(r'<tableColumn\b[^>]*>', replace_column, text)
        return new_text.encode('utf-8')

    def _write_theme_xml(self, zipf):
        """Writes xl/theme/theme1.xml.

        Uses the loaded theme when available; otherwise injects a minimal default
        theme for workbooks that contain new chartEx charts.
        """
        theme_xml = getattr(self._workbook, '_theme_xml', None)
        if theme_xml is not None:
            zipf.writestr('xl/theme/theme1.xml', theme_xml)
        elif self._needs_default_theme():
            zipf.writestr('xl/theme/theme1.xml', _DEFAULT_THEME_XML.encode('utf-8'))
    
    # Style management methods for XML creation
    
    def register_default_styles(self):
        """Registers default styles for fonts, fills, borders, and alignments."""
        # Default font (Calibri, 11pt, black) - index 0
        if 0 not in self._workbook._font_styles:
            self._workbook._font_styles[0] = {
                'name': 'Calibri',
                'size': 11,
                'color': 'FF000000',
                'color_type': 'rgb',
                'color_value': 'FF000000',
                'color_tint': None,
                'family': None,
                'charset': None,
                'scheme': None,
                'bold': False,
                'italic': False,
                'underline': False,
                'underline_type': 'none',
                'vertical_alignment': 'baseline',
                'strikethrough': False
            }
        
        # Default fills
        if 0 not in self._workbook._fill_styles:
            self._workbook._fill_styles[0] = {  # No fill
                'pattern_type': 'none',
                'fg_color': 'FFFFFFFF',
                'bg_color': 'FFFFFFFF',
                'fg_color_type': 'rgb',
                'fg_color_value': 'FFFFFFFF',
                'fg_color_tint': None,
                'bg_color_type': 'rgb',
                'bg_color_value': 'FFFFFFFF',
                'bg_color_tint': None,
            }
        if 1 not in self._workbook._fill_styles:
            self._workbook._fill_styles[1] = {  # Gray pattern
                'pattern_type': 'gray125',
                'fg_color': 'FFFFFFFF',
                'bg_color': 'FFFFFFFF',
                'fg_color_type': 'rgb',
                'fg_color_value': 'FFFFFFFF',
                'fg_color_tint': None,
                'bg_color_type': 'rgb',
                'bg_color_value': 'FFFFFFFF',
                'bg_color_tint': None,
            }
        
        # Default borders
        if 0 not in self._workbook._border_styles:
            self._workbook._border_styles[0] = {
                'top': {
                    'style': 'none', 'color': 'FF000000',
                    'automatic_color': False,
                },
                'bottom': {
                    'style': 'none', 'color': 'FF000000',
                    'automatic_color': False,
                },
                'left': {
                    'style': 'none', 'color': 'FF000000',
                    'automatic_color': False,
                },
                'right': {
                    'style': 'none', 'color': 'FF000000',
                    'automatic_color': False,
                },
            }

        # Default protection (locked=True, hidden=False)
        if 0 not in self._workbook._protection_styles:
            self._workbook._protection_styles[0] = {
                'locked': True,
                'hidden': False
            }
        
        # Default alignment (general/bottom) - index 0
        if 0 not in self._workbook._alignment_styles:
            self._workbook._alignment_styles[0] = {
                'horizontal': 'general',
                'vertical': 'bottom',
                'wrap_text': False,
                'indent': 0,
                'text_rotation': 0,
                'shrink_to_fit': False,
                'reading_order': 0,
                'relative_indent': 0
            }
    
    def get_or_create_font_style(self, font):
        """Gets or creates a font style index."""
        # Check if this font already exists by comparing with existing fonts
        for idx, font_data in self._workbook._font_styles.items():
            if (font_data['name'] == font.name and
                font_data['size'] == font.size and
                font_data['color'] == font.color and
                font_data['bold'] == font.bold and
                font_data['italic'] == font.italic and
                font_data['underline'] == font.underline and
                font_data.get(
                    'underline_type',
                    'single' if font_data['underline'] else 'none',
                ) == getattr(
                    font, 'underline_type',
                    'single' if font.underline else 'none',
                ) and
                font_data.get('vertical_alignment', 'baseline') == getattr(
                    font, 'vertical_alignment', 'baseline'
                ) and
                font_data['strikethrough'] == font.strikethrough):
                return idx
        
        # Create new font style
        new_idx = len(self._workbook._font_styles)
        self._workbook._font_styles[new_idx] = {
            'name': font.name,
            'size': font.size,
            'color': font.color,
            'bold': font.bold,
            'italic': font.italic,
            'underline': font.underline,
            'underline_type': getattr(
                font, 'underline_type',
                'single' if font.underline else 'none',
            ),
            'vertical_alignment': getattr(
                font, 'vertical_alignment', 'baseline'
            ),
            'strikethrough': font.strikethrough
        }
        return new_idx
    
    def get_or_create_fill_style(self, fill):
        """Gets or creates a fill style index."""
        fg_color_type = getattr(fill, '_fg_color_type', 'rgb')
        fg_color_value = getattr(fill, '_fg_color_value', fill.foreground_color)
        fg_color_tint = getattr(fill, '_fg_color_tint', None)
        bg_color_type = getattr(fill, '_bg_color_type', 'rgb')
        bg_color_value = getattr(fill, '_bg_color_value', fill.background_color)
        bg_color_tint = getattr(fill, '_bg_color_tint', None)
        # Check if this fill already exists by comparing with existing fills
        for idx, fill_data in self._workbook._fill_styles.items():
            if (fill_data['pattern_type'] == fill.pattern_type and
                fill_data['fg_color'] == fill.foreground_color and
                fill_data['bg_color'] == fill.background_color and
                fill_data.get('fg_color_type', 'rgb') == fg_color_type and
                fill_data.get('fg_color_value', fill_data.get('fg_color')) == fg_color_value and
                fill_data.get('fg_color_tint') == fg_color_tint and
                fill_data.get('bg_color_type', 'rgb') == bg_color_type and
                fill_data.get('bg_color_value', fill_data.get('bg_color')) == bg_color_value and
                fill_data.get('bg_color_tint') == bg_color_tint):
                return idx
        
        # Create new fill style
        new_idx = len(self._workbook._fill_styles)
        self._workbook._fill_styles[new_idx] = {
            'pattern_type': fill.pattern_type,
            'fg_color': fill.foreground_color,
            'bg_color': fill.background_color,
            'fg_color_type': fg_color_type,
            'fg_color_value': fg_color_value,
            'fg_color_tint': fg_color_tint,
            'bg_color_type': bg_color_type,
            'bg_color_value': bg_color_value,
            'bg_color_tint': bg_color_tint,
        }
        return new_idx
    
    def get_or_create_border_style(self, borders):
        """Gets or creates a border style index."""
        # Check if this border already exists by comparing with existing borders
        for idx, border_data in self._workbook._border_styles.items():
            if (border_data['top']['style'] == borders.top.line_style and
                border_data['top']['color'] == borders.top.color and
                border_data['top'].get('automatic_color', False) ==
                borders.top.automatic_color and
                border_data['bottom']['style'] == borders.bottom.line_style and
                border_data['bottom']['color'] == borders.bottom.color and
                border_data['bottom'].get('automatic_color', False) ==
                borders.bottom.automatic_color and
                border_data['left']['style'] == borders.left.line_style and
                border_data['left']['color'] == borders.left.color and
                border_data['left'].get('automatic_color', False) ==
                borders.left.automatic_color and
                border_data['right']['style'] == borders.right.line_style and
                border_data['right']['color'] == borders.right.color and
                border_data['right'].get('automatic_color', False) ==
                borders.right.automatic_color):
                return idx
        
        # Create new border style with all four sides
        new_idx = len(self._workbook._border_styles)
        self._workbook._border_styles[new_idx] = {
            'top': {
                'style': borders.top.line_style,
                'color': borders.top.color,
                'automatic_color': borders.top.automatic_color,
            },
            'bottom': {
                'style': borders.bottom.line_style,
                'color': borders.bottom.color,
                'automatic_color': borders.bottom.automatic_color,
            },
            'left': {
                'style': borders.left.line_style,
                'color': borders.left.color,
                'automatic_color': borders.left.automatic_color,
            },
            'right': {
                'style': borders.right.line_style,
                'color': borders.right.color,
                'automatic_color': borders.right.automatic_color,
            }
        }
        return new_idx
    
    def get_or_create_alignment_style(self, alignment):
        """Gets or creates an alignment style index."""
        # Check if this alignment already exists by comparing with existing alignments
        for idx, align_data in self._workbook._alignment_styles.items():
            if (align_data['horizontal'] == alignment.horizontal and
                align_data['vertical'] == alignment.vertical and
                align_data['wrap_text'] == alignment.wrap_text and
                align_data['indent'] == alignment.indent and
                align_data['text_rotation'] == alignment.text_rotation and
                align_data['shrink_to_fit'] == alignment.shrink_to_fit and
                align_data['reading_order'] == alignment.reading_order and
                align_data['relative_indent'] == alignment.relative_indent):
                return idx
        
        # Create new alignment style
        new_idx = len(self._workbook._alignment_styles)
        self._workbook._alignment_styles[new_idx] = {
            'horizontal': alignment.horizontal,
            'vertical': alignment.vertical,
            'wrap_text': alignment.wrap_text,
            'indent': alignment.indent,
            'text_rotation': alignment.text_rotation,
            'shrink_to_fit': alignment.shrink_to_fit,
            'reading_order': alignment.reading_order,
            'relative_indent': alignment.relative_indent
        }
        return new_idx

    def get_or_create_protection_style(self, protection):
        """Gets or creates a protection style index."""
        # Check if this protection already exists by comparing with existing protection styles
        for idx, prot_data in self._workbook._protection_styles.items():
            if (prot_data['locked'] == protection.locked and
                prot_data['hidden'] == protection.hidden):
                return idx

        # Create new protection style
        new_idx = len(self._workbook._protection_styles)
        self._workbook._protection_styles[new_idx] = {
            'locked': protection.locked,
            'hidden': protection.hidden
        }
        return new_idx

    def get_or_create_number_format_style(self, number_format):
        """Gets or creates a number format style index."""
        # Built-in number formats (0-163)
        builtin_formats = {
            'General': 0,
            '0': 1,
            '0.00': 2,
            '#,##0': 3,
            '#,##0.00': 4,
            '$#,##0_);($#,##0)': 5,
            '$#,##0_);[Red]($#,##0)': 6,
            '$#,##0.00_);($#,##0.00)': 7,
            '$#,##0.00_);[Red]($#,##0.00)': 8,
            '0%': 9,
            '0.00%': 10,
            '0.00E+00': 11,
            '# ?/?': 12,
            '# ??/??': 13,
            'mm-dd-yy': 14,
            'd-mmm-yy': 15,
            'd-mmm': 16,
            'mmm-yy': 17,
            'h:mm AM/PM': 18,
            'h:mm:ss AM/PM': 19,
            'h:mm': 20,
            'h:mm:ss': 21,
            'm/d/yy h:mm': 22,
            '#,##0_);(#,##0)': 37,
            '#,##0_);[Red](#,##0)': 38,
            '#,##0.00_);(#,##0.00)': 39,
            '#,##0.00_);[Red](#,##0.00)': 40,
            'mm:ss': 45,
            '[h]:mm:ss': 46,
            'mm:ss.0': 47,
            '##0.0E+0': 48,
            '@': 49
        }
        
        # Check if this is a built-in format
        if number_format in builtin_formats:
            return builtin_formats[number_format]
        
        # Check if this custom number format already exists
        for idx, fmt in self._workbook._num_formats.items():
            if fmt == number_format:
                return idx
        
        # Create new custom number format style (start from ID 164)
        new_idx = 164 + len([k for k in self._workbook._num_formats.keys() if k >= 164])
        self._workbook._num_formats[new_idx] = number_format
        return new_idx
    
    def get_or_create_cell_style(self, cell):
        """Gets or creates a cell xf style index."""
        font_idx = self.get_or_create_font_style(cell.style.font)
        fill_idx = self.get_or_create_fill_style(cell.style.fill)
        border_idx = self.get_or_create_border_style(cell.style.borders)
        num_fmt_idx = self.get_or_create_number_format_style(cell.style.number_format)
        alignment_idx = self.get_or_create_alignment_style(cell.style.alignment)
        protection_idx = self.get_or_create_protection_style(cell.style.protection)

        # Check if this is the default style (all indices are 0)
        # If so, return 0 to use the default xf in cellXfs
        if (font_idx == 0 and fill_idx == 0 and border_idx == 0 and
            num_fmt_idx == 0 and alignment_idx == 0 and protection_idx == 0):
            return 0

        key = (font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx)
        if key not in self._workbook._cell_styles:
            # Allocate the next free xf index across both style maps.
            used_indices = set(self._workbook._cell_styles.values())
            if hasattr(self._workbook, '_cell_xf_by_index') and self._workbook._cell_xf_by_index:
                used_indices.update(self._workbook._cell_xf_by_index.keys())
            next_idx = (max(used_indices) + 1) if used_indices else 1
            self._workbook._cell_styles[key] = next_idx
            if not hasattr(self._workbook, '_cell_xf_by_index'):
                self._workbook._cell_xf_by_index = {}
            self._workbook._cell_xf_by_index[next_idx] = key
        return self._workbook._cell_styles[key]

    def _write_core_properties_xml(self, zipf):
        """
        Writes docProps/core.xml file with core document properties.
        
        ECMA-376 Part 2, Section 11 - Core Properties
        
        Uses Dublin Core (dc:) and OPC Core Properties (cp:) namespaces.
        """
        from datetime import datetime, timezone
        
        # Get document properties if available
        doc_props = getattr(self._workbook, 'document_properties', None)

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        content += 'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        content += 'xmlns:dcterms="http://purl.org/dc/terms/" '
        content += 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        content += 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'

        if doc_props and doc_props.core:
            core = doc_props.core

            if core.title:
                content += f'    <dc:title>{self._escape_xml(core.title)}</dc:title>\n'
            if core.subject:
                content += f'    <dc:subject>{self._escape_xml(core.subject)}</dc:subject>\n'
            if core.creator:
                content += f'    <dc:creator>{self._escape_xml(core.creator)}</dc:creator>\n'
            if core.keywords:
                content += f'    <cp:keywords>{self._escape_xml(core.keywords)}</cp:keywords>\n'
            if core.description:
                content += f'    <dc:description>{self._escape_xml(core.description)}</dc:description>\n'
            if core.last_modified_by:
                content += f'    <cp:lastModifiedBy>{self._escape_xml(core.last_modified_by)}</cp:lastModifiedBy>\n'
            if core.revision:
                content += f'    <cp:revision>{self._escape_xml(str(core.revision))}</cp:revision>\n'
            if core.category:
                content += f'    <cp:category>{self._escape_xml(core.category)}</cp:category>\n'
            if core.content_status:
                content += f'    <cp:contentStatus>{self._escape_xml(core.content_status)}</cp:contentStatus>\n'

            # Handle dates - format as W3CDTF
            if core.created:
                if isinstance(core.created, datetime):
                    created_str = core.created.strftime('%Y-%m-%dT%H:%M:%SZ')
                else:
                    created_str = str(core.created)
                content += f'    <dcterms:created xsi:type="dcterms:W3CDTF">{created_str}</dcterms:created>\n'

            if core.modified:
                if isinstance(core.modified, datetime):
                    modified_str = core.modified.strftime('%Y-%m-%dT%H:%M:%SZ')
                else:
                    modified_str = str(core.modified)
                content += f'    <dcterms:modified xsi:type="dcterms:W3CDTF">{modified_str}</dcterms:modified>\n'
        else:
            # Write default created/modified dates
            now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
            content += f'    <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>\n'
            content += f'    <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>\n'

        content += '</cp:coreProperties>\n'
        zipf.writestr('docProps/core.xml', content)

    def _write_app_properties_xml(self, zipf):
        """
        Writes docProps/app.xml file with extended/application properties.

        ECMA-376 Part 1, Section 22.2 - Extended Properties
        """
        # Get document properties if available
        doc_props = getattr(self._workbook, 'document_properties', None)

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        content += 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n'

        if doc_props and doc_props.extended:
            ext = doc_props.extended

            content += f'    <Application>{self._escape_xml(ext.application or "Microsoft Excel")}</Application>\n'
            content += f'    <DocSecurity>{ext.doc_security}</DocSecurity>\n'
            content += f'    <ScaleCrop>{"true" if ext.scale_crop else "false"}</ScaleCrop>\n'

            if ext.company:
                content += f'    <Company>{self._escape_xml(ext.company)}</Company>\n'
            if ext.manager:
                content += f'    <Manager>{self._escape_xml(ext.manager)}</Manager>\n'
            if ext.hyperlink_base:
                content += f'    <HyperlinkBase>{self._escape_xml(ext.hyperlink_base)}</HyperlinkBase>\n'
            if ext.app_version:
                content += f'    <AppVersion>{self._escape_xml(ext.app_version)}</AppVersion>\n'

            content += f'    <LinksUpToDate>{"true" if ext.links_up_to_date else "false"}</LinksUpToDate>\n'
            content += f'    <SharedDoc>{"true" if ext.shared_doc else "false"}</SharedDoc>\n'
        else:
            # Write default values
            content += '    <Application>Microsoft Excel</Application>\n'
            content += '    <DocSecurity>0</DocSecurity>\n'
            content += '    <ScaleCrop>false</ScaleCrop>\n'
            content += '    <LinksUpToDate>false</LinksUpToDate>\n'
            content += '    <SharedDoc>false</SharedDoc>\n'

        # Add heading pairs and titles of parts (worksheet names)
        worksheet_count = len(self._workbook.worksheets)
        content += '    <HeadingPairs>\n'
        content += '        <vt:vector size="2" baseType="variant">\n'
        content += '            <vt:variant>\n'
        content += '                <vt:lpstr>Worksheets</vt:lpstr>\n'
        content += '            </vt:variant>\n'
        content += '            <vt:variant>\n'
        content += f'                <vt:i4>{worksheet_count}</vt:i4>\n'
        content += '            </vt:variant>\n'
        content += '        </vt:vector>\n'
        content += '    </HeadingPairs>\n'

        content += '    <TitlesOfParts>\n'
        content += f'        <vt:vector size="{worksheet_count}" baseType="lpstr">\n'
        for worksheet in self._workbook.worksheets:
            content += f'            <vt:lpstr>{self._escape_xml(worksheet.name)}</vt:lpstr>\n'
        content += '        </vt:vector>\n'
        content += '    </TitlesOfParts>\n'

        content += '</Properties>\n'
        zipf.writestr('docProps/app.xml', content)
