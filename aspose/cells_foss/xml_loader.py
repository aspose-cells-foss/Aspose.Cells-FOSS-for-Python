"""
Aspose.Cells for Python - XML Loader Module

This module provides XML loading functionality for Excel workbooks.
It handles parsing of workbook XML files and loading data into Workbook objects.

Compatible with Aspose.Cells for .NET API structure.
ECMA-376 Compliant cell value import.
"""

import xml.etree.ElementTree as ET
import posixpath
import re
from .cell_value_handler import CellValueHandler
from .comment_xml import CommentXMLReader
from .xml_autofilter_loader import AutoFilterXMLLoader
from .xml_conditional_format_loader import ConditionalFormatXMLLoader
from .xml_properties_loader import WorkbookPropertiesXMLLoader, WorksheetPropertiesXMLLoader
from .xml_hyperlink_handler import HyperlinkXMLLoader
from .xml_datavalidation_loader import DataValidationXmlLoader
from .xml_chart_loader import ChartXmlLoader
from .xml_table_loader import TableXmlLoader
from .xml_sparkline_loader import SparklineXmlLoader
from .style_resolver import StyleResolver
from .rich_text import RichTextRun


class XMLLoader:
    """
    Handles loading of Excel workbook XML files.
    
    This class provides methods to parse various XML components of an Excel workbook
    including workbook structure, shared strings, styles, and worksheet data.
    """
    
    def __init__(self, workbook):
        """
        Initializes the XML loader with a workbook instance.
        
        Args:
            workbook (Workbook): The workbook instance to load data into.
        """
        self.workbook = workbook
        self.ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        # Initialize comment reader
        self._comment_reader = CommentXMLReader()

        # Initialize autofilter loader
        self._autofilter_loader = AutoFilterXMLLoader(self.ns)

        # Initialize conditional format loader
        self._cf_loader = ConditionalFormatXMLLoader(self.ns, workbook)

        # Initialize hyperlink loader
        self._hyperlink_loader = HyperlinkXMLLoader(self.ns)
        self._chart_loader = ChartXmlLoader(self.ns)
        self._table_loader = TableXmlLoader()
        self._sparkline_loader = SparklineXmlLoader()

        # Initialize data validation loader
        self._dv_loader = DataValidationXmlLoader(self.ns['main'])

        # Initialize properties loaders
        self._wb_props_loader = WorkbookPropertiesXMLLoader(self.ns)
        self._ws_props_loader = WorksheetPropertiesXMLLoader(self.ns)
        self._content_type_overrides = {}
        self._content_type_defaults = {}

    @staticmethod
    def _extract_root_extra_attrs(xml_text, tag_name):
        match = re.search(rf'<{tag_name}\b([^>]*)>', xml_text, re.S)
        if not match:
            return ''
        attrs = match.group(1)
        attrs = re.sub(r'\sxmlns="[^"]*"', '', attrs)
        attrs = re.sub(r'\sxmlns:r="[^"]*"', '', attrs)
        return attrs.strip()

    @staticmethod
    def _extract_raw_element(xml_text, tag_name):
        patterns = [
            rf'(<{tag_name}\b[^>]*/>)',
            rf'(<{tag_name}\b[^>]*>.*?</{tag_name}>)',
        ]
        for pattern in patterns:
            match = re.search(pattern, xml_text, re.S)
            if match:
                return match.group(1)
        return None

    def load_workbook(self, zipf):
        """
        Loads workbook data from a ZIP file.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        # Load workbook.xml to get worksheet information
        workbook_xml_content = zipf.read('xl/workbook.xml')
        workbook_xml_text = workbook_xml_content.decode('utf-8', errors='replace')
        workbook_root = ET.fromstring(workbook_xml_content)
        self.workbook._source_workbook_root_extra_attrs = self._extract_root_extra_attrs(workbook_xml_text, 'workbook')
        self.workbook._source_workbook_alt_content_xml = self._extract_raw_element(workbook_xml_text, 'mc:AlternateContent')
        self.workbook._source_workbook_revision_ptr_xml = self._extract_raw_element(workbook_xml_text, 'xr:revisionPtr')
        self.workbook._source_workbook_extlst_xml = self._extract_raw_element(workbook_xml_text, 'extLst')

        # Load workbook properties
        self._load_workbook_properties(workbook_root)

        # Load document properties (docProps/core.xml and docProps/app.xml)
        self._load_document_properties(zipf)

        # Load package content type overrides (used by chart/theme roundtrip)
        self._load_content_type_overrides(zipf)

        # Load workbook theme part if present
        self._load_theme(zipf)

        # Load worksheet information
        self._load_worksheet_info(workbook_root)

        # Detect chartsheets and preserve their parts for round-trip.
        self._load_chartsheets(zipf, workbook_root)

        # Apply worksheet print areas/titles from workbook defined names.
        self._apply_print_settings_from_defined_names()

        # Load extra workbook rels (external links, etc.) for round-trip preservation
        self._load_extra_workbook_rels(zipf)

        # Load shared strings
        self._load_shared_strings(zipf)

        # Load styles
        self._load_styles(zipf)

        # Load worksheet data
        self._load_worksheets_data(zipf)

    def _load_content_type_overrides(self, zipf):
        """
        Loads [Content_Types].xml entries:
            Defaults: extension -> content-type
            Overrides: /part/name.xml -> content-type
        """
        self._content_type_defaults = {}
        self.workbook._content_type_defaults = {}
        self._content_type_overrides = {}
        try:
            content_types_xml = zipf.read('[Content_Types].xml')
            root = ET.fromstring(content_types_xml)
            ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
            for default in root.findall('ct:Default', namespaces=ns):
                extension = default.get('Extension')
                content_type = default.get('ContentType')
                if extension and content_type:
                    self._content_type_defaults[extension.lower()] = content_type
                    self.workbook._content_type_defaults[extension.lower()] = content_type
            for ov in root.findall('ct:Override', namespaces=ns):
                part_name = ov.get('PartName')
                content_type = ov.get('ContentType')
                if part_name and content_type:
                    self._content_type_overrides[part_name] = content_type
        except KeyError:
            self._content_type_defaults = {}
            self.workbook._content_type_defaults = {}
            self._content_type_overrides = {}
        except ET.ParseError:
            self._content_type_defaults = {}
            self.workbook._content_type_defaults = {}
            self._content_type_overrides = {}

    def _load_theme(self, zipf):
        """
        Loads the workbook theme and its color scheme.
        """
        try:
            self.workbook._theme_xml = zipf.read('xl/theme/theme1.xml')
        except KeyError:
            self.workbook._theme_xml = None
            self.workbook._theme_colors = {}
            return

        self.workbook._theme_colors = {}
        try:
            root = ET.fromstring(self.workbook._theme_xml)
        except ET.ParseError:
            return

        drawing_ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        scheme = root.find('.//a:clrScheme', namespaces=drawing_ns)
        if scheme is None:
            return
        for name in (
            'lt1', 'dk1', 'lt2', 'dk2',
            'accent1', 'accent2', 'accent3',
            'accent4', 'accent5', 'accent6',
            'hlink', 'folHlink',
        ):
            color_group = scheme.find(f'a:{name}', namespaces=drawing_ns)
            if color_group is None or not list(color_group):
                continue
            color = list(color_group)[0]
            value = color.get('lastClr') or color.get('val')
            if value:
                self.workbook._theme_colors[name] = 'FF' + value[-6:].upper()
    
    def _load_workbook_properties(self, workbook_root):
        """
        Loads workbook properties from workbook XML.

        Args:
            workbook_root: The XML root element of workbook.xml.
        """
        props = self.workbook.properties

        # Load file version
        self._wb_props_loader.load_file_version(props.file_version, workbook_root)

        # Load workbook properties
        self._wb_props_loader.load_workbook_pr(props.workbook_pr, workbook_root)

        # Load workbook protection
        self._wb_props_loader.load_workbook_protection(props.protection, workbook_root)

        # Load book views
        self._wb_props_loader.load_book_views(props.view, workbook_root)

        # Load calculation properties
        self._wb_props_loader.load_calc_pr(props.calculation, workbook_root)

        # Load defined names
        self._wb_props_loader.load_defined_names(props.defined_names, workbook_root)

    def _load_worksheet_properties(self, worksheet, worksheet_root):
        """
        Loads worksheet properties from worksheet XML.

        Args:
            worksheet (Worksheet): The worksheet to load properties into.
            worksheet_root: The XML root element of the worksheet.
        """
        props = worksheet.properties

        # Load sheet views (includes selection and pane)
        self._ws_props_loader.load_sheet_views(props, worksheet_root)

        # Load sheet format properties
        self._ws_props_loader.load_sheet_format_pr(props, worksheet_root)

        # Load sheet protection
        self._ws_props_loader.load_sheet_protection(props, worksheet_root)

        # Load print options
        self._ws_props_loader.load_print_options(props, worksheet_root)

        # Load page margins
        self._ws_props_loader.load_page_margins(props, worksheet_root)

        # Load page setup
        self._ws_props_loader.load_page_setup(props, worksheet_root)

        # Load header/footer
        self._ws_props_loader.load_header_footer(props, worksheet_root)

    def _load_data_validations(self, worksheet, worksheet_root):
        """
        Loads data validations from worksheet XML.

        Args:
            worksheet (Worksheet): The worksheet to load validations into.
            worksheet_root: The XML root element of the worksheet.
        """
        validations = self._dv_loader.load_data_validations(worksheet_root)
        worksheet._data_validations = validations

    def _load_worksheet_info(self, workbook_root):
        """
        Loads worksheet information from workbook XML.

        Args:
            workbook_root: The XML root element of workbook.xml.
        """
        from .worksheet import Worksheet

        sheets = workbook_root.findall('.//main:sheet', namespaces=self.ns)
        for sheet in sheets:
            sheet_name = sheet.get('name')
            worksheet = Worksheet(sheet_name)
            worksheet._workbook = self.workbook

            # Load visibility state
            state = sheet.get('state')
            if state == 'hidden':
                worksheet._visible = False
            elif state == 'veryHidden':
                worksheet._visible = 'veryHidden'

            self.workbook._worksheets.append(worksheet)

    CHARTSHEET_REL_TYPE = (
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'
    )

    def _load_chartsheets(self, zipf, workbook_root):
        """
        Marks sheets that are chartsheets and preserves their whole part subtree.

        A chartsheet is a `<sheet>` in workbook.xml whose workbook relationship has
        the chartsheet type. This library has no chartsheet object model, so the
        chartsheet XML and everything it transitively references (drawing, charts,
        printerSettings, ...) is captured verbatim and re-emitted on save. Without
        this the sheet would be rewritten as an empty worksheet and its charts lost,
        which corrupts the package.
        """
        try:
            rels_bytes = zipf.read('xl/_rels/workbook.xml.rels')
        except KeyError:
            return

        try:
            rels_root = ET.fromstring(rels_bytes)
        except ET.ParseError:
            return

        rns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        rel_by_id = {}
        for rel in rels_root.findall('r:Relationship', rns):
            rel_by_id[rel.get('Id')] = (rel.get('Type', ''), rel.get('Target', ''))

        r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        sheets = workbook_root.findall('.//main:sheet', namespaces=self.ns)
        for idx, sheet in enumerate(sheets):
            if idx >= len(self.workbook._worksheets):
                break
            rid = sheet.get(f'{{{r_ns}}}id')
            rel = rel_by_id.get(rid)
            if not rel:
                continue
            rel_type, target = rel
            worksheet = self.workbook._worksheets[idx]

            # Record the real source part number. Sheet parts are not necessarily
            # numbered 1..N in workbook order once chartsheets are interleaved, so the
            # positional index cannot be trusted to locate a worksheet's own part.
            m = re.search(r'(\d+)\.xml$', target)
            if m:
                worksheet._source_sheet_num = int(m.group(1))

            if rel_type != self.CHARTSHEET_REL_TYPE:
                continue

            part_path = self._resolve_part_path('xl', target)
            try:
                zipf.read(part_path)
            except KeyError:
                continue

            parts, overrides = self._collect_part_subtree(zipf, part_path)
            worksheet._is_chartsheet = True
            worksheet._chartsheet_part_path = part_path
            worksheet._chartsheet_parts = parts
            worksheet._chartsheet_content_types = overrides

    def _collect_part_subtree(self, zipf, root_part_path):
        """
        Reads a part and every internal part reachable through its rels chain.

        Returns (parts, overrides): `parts` is a list of {part_path, part_bytes}
        covering the root part, each part's own .rels file, and all descendants;
        `overrides` is a list of (part_name, content_type) for parts that carry a
        content-type Override in the source package.
        """
        parts = []
        overrides = []
        seen = set()
        queue = [root_part_path]

        while queue:
            path = queue.pop(0)
            if path in seen:
                continue
            seen.add(path)
            try:
                data = zipf.read(path)
            except KeyError:
                continue

            parts.append({'part_path': path, 'part_bytes': data})
            ctype = self._content_type_overrides.get(f'/{path}')
            if ctype:
                overrides.append((f'/{path}', ctype))

            dir_part, _, file_part = path.rpartition('/')
            rels_path = f'{dir_part}/_rels/{file_part}.rels'
            if rels_path in seen:
                continue
            try:
                rels_bytes = zipf.read(rels_path)
            except KeyError:
                continue
            seen.add(rels_path)
            parts.append({'part_path': rels_path, 'part_bytes': rels_bytes})

            try:
                rels_root = ET.fromstring(rels_bytes)
            except ET.ParseError:
                continue
            rns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            for rel in rels_root.findall('r:Relationship', rns):
                if rel.get('TargetMode', '') == 'External':
                    continue
                queue.append(self._resolve_part_path(dir_part, rel.get('Target', '')))

        return parts, overrides

    def _apply_print_settings_from_defined_names(self):
        """
        Applies worksheet print areas and repeated titles from defined names.
        """
        defined_names = getattr(self.workbook.properties, 'defined_names', None)
        if defined_names is None:
            return

        for dn in defined_names:
            if dn.name not in ('_xlnm.Print_Area', '_xlnm.Print_Titles'):
                continue
            sheet_id = dn.local_sheet_id
            if sheet_id is None or sheet_id < 0 or sheet_id >= len(self.workbook._worksheets):
                continue
            worksheet = self.workbook._worksheets[sheet_id]
            if dn.name == '_xlnm.Print_Area':
                worksheet._print_area = self._extract_print_area(
                    dn.refers_to, worksheet.name
                )
            else:
                rows, columns = self._extract_print_titles(dn.refers_to)
                worksheet.page_setup.print_title_rows = rows
                worksheet.page_setup.print_title_columns = columns

    def _extract_print_area(self, refers_to, sheet_name):
        """
        Extracts plain A1 print area string from defined-name formula text.

        Example input:
            "'Sheet1'!$A$1:$D$20,'Sheet1'!$F$1:$F$20"
        Output:
            "A1:D20,F1:F20"
        """
        if not refers_to:
            return None

        parts = []
        for token in self._split_defined_name_references(refers_to):
            part = token.strip()
            if not part:
                continue
            if '!' in part:
                _, addr = part.split('!', 1)
            else:
                addr = part
            addr = addr.replace('$', '').strip().upper()
            parts.append(addr)

        if not parts:
            return None
        return ','.join(parts)

    def _extract_print_titles(self, refers_to):
        rows = None
        columns = None
        for token in self._split_defined_name_references(refers_to):
            address = token.rsplit('!', 1)[-1].replace('$', '').strip().upper()
            if re.fullmatch(r'\d+(?::\d+)?', address):
                rows = address
            elif re.fullmatch(r'[A-Z]+(?::[A-Z]+)?', address):
                columns = address
        return rows, columns

    @staticmethod
    def _split_defined_name_references(refers_to):
        parts = []
        current = []
        quoted = False
        text = str(refers_to or '')
        index = 0
        while index < len(text):
            char = text[index]
            if char == "'":
                if quoted and index + 1 < len(text) and text[index + 1] == "'":
                    current.extend((char, char))
                    index += 2
                    continue
                quoted = not quoted
            if char == ',' and not quoted:
                part = ''.join(current).strip()
                if part:
                    parts.append(part)
                current = []
            else:
                current.append(char)
            index += 1
        part = ''.join(current).strip()
        if part:
            parts.append(part)
        return parts
    
    def _load_shared_strings(self, zipf):
        """
        Loads shared strings from the workbook.
        
        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            shared_strings_content = zipf.read('xl/sharedStrings.xml')
            shared_strings_root = ET.fromstring(shared_strings_content)
            self.workbook._shared_strings = []
            self.workbook._shared_string_runs = []
            for si in shared_strings_root.findall('.//main:si', namespaces=self.ns):
                text_parts = [
                    t.text if t.text is not None else ''
                    for t in si.findall('.//main:t', namespaces=self.ns)
                ]
                self.workbook._shared_strings.append(''.join(text_parts))
                rich_runs = tuple(
                    self._parse_shared_string_run(run)
                    for run in si.findall('main:r', namespaces=self.ns)
                )
                self.workbook._shared_string_runs.append(rich_runs)
        except KeyError:
            self.workbook._shared_strings = []
            self.workbook._shared_string_runs = []

    def _parse_shared_string_run(self, run):
        text = ''.join(
            node.text if node.text is not None else ''
            for node in run.findall('main:t', namespaces=self.ns)
        )
        properties = run.find('main:rPr', namespaces=self.ns)
        if properties is None:
            return RichTextRun(text=text)

        font_name = properties.find('main:rFont', namespaces=self.ns)
        font_size = properties.find('main:sz', namespaces=self.ns)
        color = properties.find('main:color', namespaces=self.ns)
        underline = properties.find('main:u', namespaces=self.ns)
        vertical_alignment = properties.find(
            'main:vertAlign', namespaces=self.ns
        )
        return RichTextRun(
            text=text,
            font_name=(font_name.get('val') if font_name is not None else None),
            font_size=(
                float(font_size.get('val'))
                if font_size is not None and font_size.get('val') is not None
                else None
            ),
            font_color=self._resolve_color_element(color),
            bold=self._optional_run_property(properties, 'b'),
            italic=self._optional_run_property(properties, 'i'),
            underline=self._optional_run_property(properties, 'u'),
            underline_type=(
                underline.get('val', 'single')
                if underline is not None else None
            ),
            strikethrough=self._optional_run_property(properties, 'strike'),
            vertical_alignment=(
                vertical_alignment.get('val', 'baseline')
                if vertical_alignment is not None else 'baseline'
            ),
        )

    def _resolve_color_element(self, color):
        if color is None:
            return None
        rgb = color.get('rgb')
        if rgb:
            return rgb.upper() if len(rgb) == 8 else 'FF' + rgb.upper()
        theme = color.get('theme')
        if theme is not None:
            return self._resolve_theme_color(theme)
        return None

    @staticmethod
    def _optional_run_property(properties, name):
        element = properties.find(
            f'{{http://schemas.openxmlformats.org/spreadsheetml/2006/main}}{name}'
        )
        if element is None:
            return None
        return element.get('val', '1') not in ('0', 'false', 'False', 'none')

    def _resolve_theme_color(self, theme_index):
        names = (
            'lt1', 'dk1', 'lt2', 'dk2',
            'accent1', 'accent2', 'accent3',
            'accent4', 'accent5', 'accent6',
            'hlink', 'folHlink',
        )
        try:
            name = names[int(theme_index)]
        except (TypeError, ValueError, IndexError):
            return None
        return self.workbook._theme_colors.get(name)
    
    def _load_styles(self, zipf):
        """
        Loads styles from the workbook.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            styles_content = zipf.read('xl/styles.xml')
            self.workbook._source_styles_xml = styles_content
            styles_root = ET.fromstring(styles_content)
            cell_xfs_elem = styles_root.find('.//main:cellXfs', namespaces=self.ns)
            self.workbook._source_cell_xfs_count = int(cell_xfs_elem.get('count', '0')) if cell_xfs_elem is not None else 0
            self._load_styles_xml(styles_root)
            # Load differential formatting (dxf) for conditional formatting
            self._load_dxf_styles(styles_root)
            self._load_table_styles(styles_root)
        except KeyError:
            # Use default styles
            from .xml_saver import XMLSaver
            saver = XMLSaver(self.workbook)
            saver.register_default_styles()
            self.workbook._dxf_styles = []
    
    @staticmethod
    def _resolve_part_path(base_dir, target):
        """
        Resolves a relationship Target into a zip part name.

        Targets are relative to the directory of the part that owns the rels file and
        may walk out of it ("../customXml/item1.xml"), so the joined path has to be
        normalized before it can be looked up in the archive.
        """
        if target.startswith('/'):
            return target[1:]
        return posixpath.normpath(posixpath.join(base_dir, target))

    def _collect_child_parts(self, zipf, base_dir, rels_bytes):
        """
        Reads the internal parts referenced by a rels file.

        Returns a list of dicts with part_path/part_bytes/content_type for every
        target that exists in the archive.
        """
        try:
            rels_root = ET.fromstring(rels_bytes)
        except ET.ParseError:
            return []

        ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        children = []
        for rel in rels_root.findall('r:Relationship', ns):
            if rel.get('TargetMode', '') == 'External':
                continue
            child_path = self._resolve_part_path(base_dir, rel.get('Target', ''))
            try:
                child_bytes = zipf.read(child_path)
            except KeyError:
                continue
            children.append({
                'part_path': child_path,
                'part_bytes': child_bytes,
                'content_type': self._content_type_overrides.get(f'/{child_path}'),
            })
        return children

    def _load_extra_workbook_rels(self, zipf):
        """
        Loads non-native workbook rels (e.g. externalLinks) for round-trip preservation.
        Stores them on workbook._source_extra_workbook_rels.
        """
        rels_path = 'xl/_rels/workbook.xml.rels'
        try:
            rels_content = zipf.read(rels_path)
        except KeyError:
            return

        rels_root = ET.fromstring(rels_content)
        ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}

        # Rel types that the saver writes natively — skip these.
        # Chartsheets are handled by _load_chartsheets/the chartsheet save path;
        # capturing them here too would double-write their parts and mis-type them.
        NATIVE_TYPES = {
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
            self.CHARTSHEET_REL_TYPE,
        }

        extra_rels = []
        for rel in rels_root.findall('r:Relationship', ns):
            rel_type = rel.get('Type', '')
            if rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain':
                target = rel.get('Target', '')
                rel_id = rel.get('Id', '')
                part_path = self._resolve_part_path('xl', target)
                try:
                    self.workbook._source_calc_chain_bytes = zipf.read(part_path)
                    self.workbook._source_calc_chain_rel = {
                        'rel_id': rel_id,
                        'target': target,
                        'part_path': part_path,
                        'content_type': self._content_type_overrides.get(f'/{part_path}')
                            or 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
                    }
                except KeyError:
                    pass
                continue
            if rel_type in NATIVE_TYPES:
                continue
            target = rel.get('Target', '')
            rel_id = rel.get('Id', '')
            target_mode = rel.get('TargetMode', '')
            if target_mode == 'External':
                continue

            # Resolve path relative to xl/ (targets may escape it, e.g. ../customXml/item1.xml)
            part_path = self._resolve_part_path('xl', target)

            part_bytes = None
            try:
                part_bytes = zipf.read(part_path)
            except KeyError:
                pass

            # A relationship whose part is missing would be dangling in the output
            # package, which makes Excel refuse to open the file. Drop it.
            if part_bytes is None:
                continue

            # Read the part's own rels file (e.g. externalLink1.xml.rels)
            dir_part, file_part = part_path.rsplit('/', 1)
            part_rels_path = f'{dir_part}/_rels/{file_part}.rels'
            part_rels_bytes = None
            try:
                part_rels_bytes = zipf.read(part_rels_path)
            except KeyError:
                pass

            # Pull in the parts that the part's own rels point at (e.g. customXml
            # itemProps), otherwise those rels dangle in turn.
            child_parts = []
            if part_rels_bytes is not None:
                child_parts = self._collect_child_parts(zipf, dir_part, part_rels_bytes)

            # Look up content type from preloaded overrides
            content_type = self._content_type_overrides.get(f'/{part_path}')

            extra_rels.append({
                'rel_id': rel_id,
                'rel_type': rel_type,
                'target': target,
                'part_path': part_path,
                'part_bytes': part_bytes,
                'part_rels_path': part_rels_path,
                'part_rels_bytes': part_rels_bytes,
                'child_parts': child_parts,
                'content_type': content_type,
            })

        if extra_rels:
            self.workbook._source_extra_workbook_rels = extra_rels

    def _load_worksheets_data(self, zipf):
        """
        Loads data for all worksheets.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        # Pre-pass: load extra sheet rels for ALL worksheets first.
        # This lets us know which comments/vmlDrawing files are "claimed" by each
        # sheet via its rels (files may use non-standard numbering, e.g. comments1.xml
        # for sheet2), so we can avoid loading another sheet's file by accident.
        worksheets_and_roots = []
        for i, worksheet in enumerate(self.workbook._worksheets):
            # Chartsheets are preserved verbatim elsewhere; they have no worksheet part.
            if getattr(worksheet, '_is_chartsheet', False):
                worksheets_and_roots.append((i, worksheet, None))
                continue
            # Locate the worksheet's own part. Defaults to the positional number, which
            # only diverges when non-worksheet sheets (chartsheets) shift the numbering.
            src_num = getattr(worksheet, '_source_sheet_num', i + 1)
            try:
                worksheet_content = zipf.read(f'xl/worksheets/sheet{src_num}.xml')
                worksheet_text = worksheet_content.decode('utf-8', errors='replace')
                worksheet_root = ET.fromstring(worksheet_content)
                worksheet._source_root_extra_attrs = self._extract_root_extra_attrs(worksheet_text, 'worksheet')
                worksheet._source_sheet_pr_xml = self._extract_raw_element(worksheet_text, 'sheetPr')
                worksheet._source_phonetic_pr_xml = self._extract_raw_element(worksheet_text, 'phoneticPr')
                dimension_xml = self._extract_raw_element(worksheet_text, 'dimension')
                worksheet._source_dimension_ref = None
                if dimension_xml:
                    m = re.search(r'ref="([^"]+)"', dimension_xml)
                    if m:
                        worksheet._source_dimension_ref = m.group(1)
                self._load_extra_sheet_rels(zipf, worksheet, src_num)
                worksheets_and_roots.append((i, worksheet, worksheet_root))
            except KeyError:
                worksheets_and_roots.append((i, worksheet, None))

        # Build map: comments part_path -> the worksheet that owns it via its rels.
        # Any comments file in this map must only be loaded by its owner sheet.
        claimed_comments_paths = {}
        for _, worksheet, _ in worksheets_and_roots:
            for r in getattr(worksheet, '_source_extra_sheet_rels', []):
                if 'comments' in r.get('rel_type', '').lower():
                    claimed_comments_paths[r['part_path']] = worksheet

        # Main pass: load cell data, comments, hyperlinks, charts.
        for i, worksheet, worksheet_root in worksheets_and_roots:
            if worksheet_root is None:
                continue
            src_num = getattr(worksheet, '_source_sheet_num', i + 1)
            try:
                self._load_worksheet_data(worksheet, worksheet_root)

                # Load comments for this worksheet.
                # Skip if the default-numbered comments file (xl/comments{N}.xml) is
                # claimed by a DIFFERENT sheet via that sheet's rels -- loading it here
                # would assign another sheet's comments to the wrong worksheet.
                default_comments_path = f'xl/comments{src_num}.xml'
                comments_owner = claimed_comments_paths.get(default_comments_path)
                if comments_owner is None or comments_owner is worksheet:
                    self._comment_reader.load_comments(zipf, worksheet, src_num)

                # Load hyperlinks for this worksheet
                self._hyperlink_loader.load_hyperlinks(worksheet, worksheet_root, zipf, src_num)

                # Load charts for this worksheet
                self._chart_loader.load_charts(
                    worksheet, worksheet_root, zipf, src_num,
                    content_type_overrides=self._content_type_overrides,
                    content_type_defaults=self._content_type_defaults,
                )

                # Load tables for this worksheet
                self._table_loader.load_tables(worksheet, worksheet_root, zipf, src_num)

                # Load sparklines for this worksheet (inline in extLst — no zipf needed)
                self._sparkline_loader.load_sparklines(worksheet, worksheet_root)
            except KeyError:
                # Worksheet file not found, skip
                pass
    
    def _load_extra_sheet_rels(self, zipf, worksheet, sheet_num):
        """
        Reads non-drawing, non-hyperlink rels from the sheet's rels file and stores
        them on worksheet._source_extra_sheet_rels for round-trip preservation.
        This handles cases where vmlDrawing/comments files use non-standard numbering.
        """
        rels_path = f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels'
        try:
            rels_content = zipf.read(rels_path)
        except KeyError:
            return

        rels_root = ET.fromstring(rels_content)
        ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}

        # Only preserve vmlDrawing and comments rels for round-trip.
        # Other rel types (drawing, hyperlink, printerSettings, etc.) are either
        # handled elsewhere or intentionally not preserved here to avoid rel-ID conflicts.
        PRESERVED_TYPES = {
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
        }

        extra_rels = []
        for rel in rels_root.findall('r:Relationship', ns):
            rel_type = rel.get('Type', '')
            target = rel.get('Target', '')
            target_mode = rel.get('TargetMode', '')
            rel_id = rel.get('Id', '')

            # Only collect vmlDrawing and comments rels
            if rel_type not in PRESERVED_TYPES:
                continue
            # Skip external targets
            if target_mode == 'External':
                continue

            # Resolve part path relative to xl/worksheets/
            # Target is like "../vmlDrawing1.vml" → "xl/vmlDrawing1.vml"
            part_path = self._resolve_part_path('xl/worksheets', target)

            part_bytes = None
            try:
                part_bytes = zipf.read(part_path)
            except KeyError:
                pass

            # Emitting a rel whose part is missing produces a dangling relationship
            # in the output package, which makes Excel refuse to open the file.
            if part_bytes is None:
                continue

            # A vmlDrawing carries its own rels when it embeds an image, and those
            # r:ids dangle just as badly if the .rels part or the media it points at
            # is left out of the saved package.
            dir_part, _, file_part = part_path.rpartition('/')
            part_rels_path = f'{dir_part}/_rels/{file_part}.rels'
            try:
                part_rels_bytes = zipf.read(part_rels_path)
            except KeyError:
                part_rels_bytes = None

            child_parts = []
            if part_rels_bytes is not None:
                child_parts = self._collect_child_parts(zipf, dir_part, part_rels_bytes)

            extra_rels.append({
                'rel_id': rel_id,
                'rel_type': rel_type,
                'target': target,
                'target_mode': target_mode,
                'part_path': part_path,
                'part_bytes': part_bytes,
                'part_rels_path': part_rels_path,
                'part_rels_bytes': part_rels_bytes,
                'child_parts': child_parts,
            })

        if extra_rels:
            worksheet._source_extra_sheet_rels = extra_rels

    def _load_worksheet_data(self, worksheet, worksheet_root):
        """
        Loads cell data from worksheet XML according to ECMA-376 specification.

        Args:
            worksheet (Worksheet): The worksheet object to load data into.
            worksheet_root: The XML root element of the worksheet.
        """
        # Load used range (dimension)
        dim_elem = worksheet_root.find('main:dimension', namespaces=self.ns)
        if dim_elem is not None:
            ref = dim_elem.get('ref')
            if ref:
                from .cells import Cells
                if ':' in ref:
                    start_ref, end_ref = ref.split(':', 1)
                else:
                    start_ref = end_ref = ref
                try:
                    min_row, min_col = Cells.coordinate_from_string(start_ref)
                    max_row, max_col = Cells.coordinate_from_string(end_ref)
                    worksheet._dimension = (min_row, min_col, max_row, max_col)
                except ValueError:
                    pass

        # Load worksheet properties
        self._load_worksheet_properties(worksheet, worksheet_root)

        # Manual break ids identify the row/column immediately before the
        # next printed page. Preserve them for PDF pagination and round-trip.
        worksheet._horizontal_page_breaks = {
            int(item.get('id'))
            for item in worksheet_root.findall(
                'main:rowBreaks/main:brk', namespaces=self.ns
            )
            if item.get('id') is not None
            and item.get('man', '1') in ('1', 'true', 'True')
        }
        worksheet._vertical_page_breaks = {
            int(item.get('id'))
            for item in worksheet_root.findall(
                'main:colBreaks/main:brk', namespaces=self.ns
            )
            if item.get('id') is not None
            and item.get('man', '1') in ('1', 'true', 'True')
        }

        # Load column widths and row heights
        self._load_column_dimensions(worksheet, worksheet_root)
        self._load_row_heights(worksheet, worksheet_root)

        # Load auto filter settings (ECMA-376 Section 18.3.1.2)
        self._autofilter_loader.load_auto_filter(worksheet, worksheet_root)

        # Load conditional formatting (ECMA-376 Section 18.3.1.18)
        self._cf_loader.load_conditional_formatting(worksheet, worksheet_root)

        # Load data validations (ECMA-376 Section 18.3.1.30, 18.3.1.31)
        self._load_data_validations(worksheet, worksheet_root)

        # Load merged cells (ECMA-376 Section 18.3.1.55)
        self._load_merged_cells(worksheet, worksheet_root)

        # Find shared string table reference
        shared_strings = self.workbook._shared_strings
        
        # Load cell data
        for row_elem in worksheet_root.findall('.//main:row', namespaces=self.ns):
            for cell_elem in row_elem.findall('main:c', namespaces=self.ns):
                cell_ref = cell_elem.get('r')
                cell_type = cell_elem.get('t', 'n')  # Default to numeric per ECMA-376
                
                # Check for formula first (ECMA-376: formula must come before value)
                f_elem = cell_elem.find('main:f', namespaces=self.ns)
                formula = f_elem.text if f_elem is not None else None
                # Add '=' prefix to formula if not present (ECMA-376 stores formulas without '=')
                if formula is not None and not formula.startswith('='):
                    formula = '=' + formula
                
                # Get cell style index
                s_elem = cell_elem.get('s')
                style_idx = int(s_elem) if s_elem is not None else 0
                
                # Get cell value using CellValueHandler for ECMA-376 compliance
                v_elem = cell_elem.find('main:v', namespaces=self.ns)
                value = None
                
                if v_elem is not None and v_elem.text is not None:
                    value_str = v_elem.text
                    # Use CellValueHandler to parse value according to ECMA-376
                    value = CellValueHandler.parse_value_from_xml(
                        value_str,
                        cell_type,
                        shared_strings
                    )
                
                # Create cell with value and formula
                from .cell import Cell
                cell = Cell(value, formula, worksheet=worksheet)
                if cell_type == 's' and v_elem is not None and v_elem.text is not None:
                    try:
                        shared_string_index = int(value_str)
                        cell._rich_text_runs = self.workbook._shared_string_runs[
                            shared_string_index
                        ]
                    except (ValueError, IndexError):
                        cell._rich_text_runs = ()
                
                # Apply style (including style 0/default) if available in style table.
                self._apply_cell_style(cell, style_idx)
                
                # Set cell value
                worksheet.cells[cell_ref] = cell

    def _load_merged_cells(self, worksheet, worksheet_root):
        """
        Loads merged cell ranges from worksheet XML.

        Args:
            worksheet (Worksheet): The worksheet object to load merged ranges into.
            worksheet_root: The XML root element of the worksheet.
        """
        merge_cells_elem = worksheet_root.find('main:mergeCells', namespaces=self.ns)
        if merge_cells_elem is None:
            worksheet._merged_cells = []
            return

        merged = []
        for merge_cell_elem in merge_cells_elem.findall('main:mergeCell', namespaces=self.ns):
            merge_ref = merge_cell_elem.get('ref')
            if merge_ref:
                merged.append(str(merge_ref).upper())
        worksheet._merged_cells = merged

    def _load_column_dimensions(self, worksheet, worksheet_root):
        """
        Loads column width settings from worksheet XML.
        """
        cols_elem = worksheet_root.find('main:cols', namespaces=self.ns)
        if cols_elem is None:
            return

        if not hasattr(worksheet, '_column_widths'):
            worksheet._column_widths = {}
        if not hasattr(worksheet, '_column_styles'):
            worksheet._column_styles = {}
        if not hasattr(worksheet, '_best_fit_columns'):
            worksheet._best_fit_columns = set()
        if not hasattr(worksheet, '_hidden_columns'):
            worksheet._hidden_columns = set()

        for col_elem in cols_elem.findall('main:col', namespaces=self.ns):
            min_val = col_elem.get('min')
            max_val = col_elem.get('max')
            width_val = col_elem.get('width')
            style_val = col_elem.get('style')
            hidden_val = col_elem.get('hidden')
            best_fit_val = col_elem.get('bestFit')
            if min_val is None or max_val is None:
                raise ValueError("Invalid column definition: missing min or max")
            try:
                # Handle float values by converting to float first, then to int
                min_col = int(float(min_val))
                max_col = int(float(max_val))
            except ValueError as exc:
                raise ValueError("Invalid column definition values") from exc
            if min_col < 1 or max_col < min_col:
                raise ValueError("Invalid column definition range")
            width = None
            if width_val is not None:
                try:
                    width = float(width_val)
                except ValueError as exc:
                    raise ValueError("Invalid column width value") from exc
                if width <= 0:
                    raise ValueError("Column width must be > 0")

            for col_idx in range(min_col, max_col + 1):
                if width is not None:
                    worksheet._column_widths[col_idx] = width
                if style_val is not None:
                    try:
                        worksheet._column_styles[col_idx] = int(style_val)
                    except ValueError:
                        pass
                if best_fit_val in ('1', 'true', 'True'):
                    worksheet._best_fit_columns.add(col_idx)
                if hidden_val in ('1', 'true', 'True'):
                    worksheet._hidden_columns.add(col_idx)

    def _load_row_heights(self, worksheet, worksheet_root):
        """
        Loads row height settings from worksheet XML.
        """
        if not hasattr(worksheet, '_row_heights'):
            worksheet._row_heights = {}
        if not hasattr(worksheet, '_custom_height_rows'):
            worksheet._custom_height_rows = set()
        if not hasattr(worksheet, '_loaded_height_rows'):
            worksheet._loaded_height_rows = set()
        if not hasattr(worksheet, '_hidden_rows'):
            worksheet._hidden_rows = set()

        for row_elem in worksheet_root.findall('.//main:row', namespaces=self.ns):
            ht = row_elem.get('ht')
            hidden_val = row_elem.get('hidden')
            if ht is None:
                if hidden_val not in ('1', 'true', 'True'):
                    continue
            row_num = row_elem.get('r')
            if row_num is None:
                raise ValueError("Row definition missing row index")
            try:
                row_idx = int(row_num)
            except ValueError as exc:
                raise ValueError("Invalid row height definition values") from exc
            if row_idx < 1:
                raise ValueError("Row index must be >= 1")
            if ht is not None:
                try:
                    height = float(ht)
                except ValueError as exc:
                    raise ValueError("Invalid row height value") from exc
                if height <= 0:
                    raise ValueError("Row height must be > 0")
                worksheet._row_heights[row_idx] = height
                worksheet._loaded_height_rows.add(row_idx)
                if row_elem.get('customHeight') in ('1', 'true', 'True'):
                    worksheet._custom_height_rows.add(row_idx)
            if hidden_val in ('1', 'true', 'True'):
                worksheet._hidden_rows.add(row_idx)

    def _apply_cell_style(self, cell, style_idx):
        """
        Applies a style to a cell based on style index.
        
        Args:
            cell (Cell): The cell to apply style to.
            style_idx (int): The style index to apply.
        """
        # Preserve original style index for lossless roundtrip.
        cell._source_style_idx = style_idx
        cell._style_index = style_idx
        cell.style = StyleResolver.style_from_index(self.workbook, style_idx)

    def _load_styles_xml(self, styles_root):
        """
        Loads styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        # Register built-in number formats
        builtin_formats = {
            0: 'General',
            1: '0',
            2: '0.00',
            3: '#,##0',
            4: '#,##0.00',
            5: '$#,##0_);($#,##0)',
            6: '$#,##0_);[Red]($#,##0)',
            7: '$#,##0.00_);($#,##0.00)',
            8: '$#,##0.00_);[Red]($#,##0.00)',
            9: '0%',
            10: '0.00%',
            11: '0.00E+00',
            12: '# ?/?',
            13: '# ??/??',
            14: 'm/d/yyyy',
            15: 'd-mmm-yy',
            16: 'd-mmm',
            17: 'mmm-yy',
            18: 'h:mm AM/PM',
            19: 'h:mm:ss AM/PM',
            20: 'h:mm',
            21: 'h:mm:ss',
            22: 'm/d/yy h:mm',
            37: '#,##0_);(#,##0)',
            38: '#,##0_);[Red](#,##0)',
            39: '#,##0.00_);(#,##0.00)',
            40: '#,##0.00_);[Red](#,##0.00)',
            45: 'mm:ss',
            46: '[h]:mm:ss',
            47: 'mm:ss.0',
            48: '##0.0E+0',
            49: '@'
        }
        self.workbook._num_formats.update(builtin_formats)
        
        # Load custom number formats
        num_fmts = styles_root.findall('.//main:numFmt', namespaces=self.ns)
        for num_fmt in num_fmts:
            num_fmt_id = int(num_fmt.get('numFmtId'))
            format_code = num_fmt.get('formatCode')
            self.workbook._num_formats[num_fmt_id] = format_code
        
        # Load fonts
        self._load_fonts(styles_root)
        
        # Load fills
        self._load_fills(styles_root)
        
        # Load borders
        self._load_borders(styles_root)
        
        # Load cellXfs
        self._load_cell_xfs(styles_root)
    
    def _load_fonts(self, styles_root):
        """
        Loads font styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        fonts = styles_root.findall('.//main:font', namespaces=self.ns)
        for i, font_elem in enumerate(fonts):
            sz_elem = font_elem.find('main:sz', namespaces=self.ns)
            color_elem = font_elem.find('main:color', namespaces=self.ns)
            name_elem = font_elem.find('main:name', namespaces=self.ns)
            family_elem = font_elem.find('main:family', namespaces=self.ns)
            charset_elem = font_elem.find('main:charset', namespaces=self.ns)
            scheme_elem = font_elem.find('main:scheme', namespaces=self.ns)
            b_elem = font_elem.find('main:b', namespaces=self.ns)
            i_elem = font_elem.find('main:i', namespaces=self.ns)
            u_elem = font_elem.find('main:u', namespaces=self.ns)
            strike_elem = font_elem.find('main:strike', namespaces=self.ns)
            vertical_alignment_elem = font_elem.find(
                'main:vertAlign', namespaces=self.ns
            )

            color_type = None
            color_value = None
            color_tint = None
            resolved_color = None
            if color_elem is not None:
                if color_elem.get('rgb') is not None:
                    color_type = 'rgb'
                    color_value = color_elem.get('rgb')
                    resolved_color = color_value
                elif color_elem.get('theme') is not None:
                    color_type = 'theme'
                    color_value = color_elem.get('theme')
                    resolved_color = self._resolve_theme_color(color_value)
                elif color_elem.get('indexed') is not None:
                    color_type = 'indexed'
                    color_value = color_elem.get('indexed')
                elif color_elem.get('auto') is not None:
                    color_type = 'auto'
                    color_value = color_elem.get('auto')
                if color_elem.get('tint') is not None:
                    color_tint = color_elem.get('tint')
            
            font_data = {
                'name': name_elem.get('val') if name_elem is not None else 'Calibri',
                'size': float(sz_elem.get('val', 11)) if sz_elem is not None else 11,
                'color': resolved_color or 'FF000000',
                'color_type': color_type,
                'color_value': color_value,
                'color_tint': color_tint,
                'family': family_elem.get('val') if family_elem is not None else None,
                'charset': charset_elem.get('val') if charset_elem is not None else None,
                'scheme': scheme_elem.get('val') if scheme_elem is not None else None,
                'bold': b_elem is not None,
                'italic': i_elem is not None,
                'underline': u_elem is not None,
                'underline_type': (
                    u_elem.get('val', 'single')
                    if u_elem is not None else 'none'
                ),
                'vertical_alignment': (
                    vertical_alignment_elem.get('val', 'baseline')
                    if vertical_alignment_elem is not None else 'baseline'
                ),
                'strikethrough': strike_elem is not None
            }
            self.workbook._font_styles[i] = font_data

            # Keep workbook default style in sync with loaded default font (fontId=0).
            if i == 0 and self.workbook._styles:
                default_font = self.workbook._styles[0].font
                default_font.name = font_data['name']
                default_font.size = font_data['size']
                default_font.bold = font_data['bold']
                default_font.italic = font_data['italic']
                default_font.underline = font_data['underline']
                default_font.underline_type = font_data['underline_type']
                default_font.vertical_alignment = font_data['vertical_alignment']
                default_font.strikethrough = font_data['strikethrough']
                # Only apply explicit RGB color to style font.
                if font_data['color_type'] == 'rgb' and font_data['color_value']:
                    default_font.color = font_data['color_value']
    
    def _load_fills(self, styles_root):
        """
        Loads fill styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        fills = styles_root.findall('.//main:fill', namespaces=self.ns)
        for i, fill_elem in enumerate(fills):
            pattern_elem = fill_elem.find('main:patternFill', namespaces=self.ns)
            fg_color_elem = pattern_elem.find('main:fgColor', namespaces=self.ns) if pattern_elem is not None else None
            bg_color_elem = pattern_elem.find('main:bgColor', namespaces=self.ns) if pattern_elem is not None else None
            fg_type, fg_value, fg_tint = self._extract_color_attrs(fg_color_elem)
            bg_type, bg_value, bg_tint = self._extract_color_attrs(bg_color_elem)
            fg_rgb = fg_value if fg_type == 'rgb' else 'FFFFFFFF'
            bg_rgb = bg_value if bg_type == 'rgb' else 'FFFFFFFF'
            fill_data = {
                'pattern_type': pattern_elem.get('patternType', 'none') if pattern_elem is not None else 'none',
                'fg_color': fg_rgb,
                'bg_color': bg_rgb,
                'fg_color_type': fg_type,
                'fg_color_value': fg_value,
                'fg_color_tint': fg_tint,
                'bg_color_type': bg_type,
                'bg_color_value': bg_value,
                'bg_color_tint': bg_tint,
            }
            self.workbook._fill_styles[i] = fill_data

    def _extract_color_attrs(self, color_elem):
        if color_elem is None:
            return 'rgb', 'FFFFFFFF', None
        tint_val = color_elem.get('tint')
        tint = float(tint_val) if tint_val is not None else None
        if color_elem.get('rgb') is not None:
            return 'rgb', color_elem.get('rgb'), tint
        if color_elem.get('theme') is not None:
            return 'theme', color_elem.get('theme'), tint
        if color_elem.get('indexed') is not None:
            return 'indexed', color_elem.get('indexed'), tint
        if color_elem.get('auto') is not None:
            return 'auto', color_elem.get('auto'), tint
        return 'rgb', 'FFFFFFFF', tint
    
    def _load_borders(self, styles_root):
        """
        Loads border styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        borders = styles_root.findall('.//main:border', namespaces=self.ns)
        for i, border_elem in enumerate(borders):
            if i == 0:
                continue  # Skip default border
            left_elem = border_elem.find('main:left', namespaces=self.ns)
            right_elem = border_elem.find('main:right', namespaces=self.ns)
            top_elem = border_elem.find('main:top', namespaces=self.ns)
            bottom_elem = border_elem.find('main:bottom', namespaces=self.ns)
            diagonal_elem = border_elem.find(
                'main:diagonal', namespaces=self.ns
            )
            
            # Load left border
            left_style = 'none'
            left_color = 'FF000000'
            left_automatic = False
            if left_elem is not None:
                left_style = left_elem.get('style', 'none')
                left_color_elem = left_elem.find('main:color', namespaces=self.ns)
                if left_color_elem is not None:
                    left_color = left_color_elem.get('rgb', 'FF000000')
                    left_automatic = left_color_elem.get('auto') in (
                        '1', 'true', 'True'
                    )
            
            # Load right border
            right_style = 'none'
            right_color = 'FF000000'
            right_automatic = False
            if right_elem is not None:
                right_style = right_elem.get('style', 'none')
                right_color_elem = right_elem.find('main:color', namespaces=self.ns)
                if right_color_elem is not None:
                    right_color = right_color_elem.get('rgb', 'FF000000')
                    right_automatic = right_color_elem.get('auto') in (
                        '1', 'true', 'True'
                    )
            
            # Load top border
            top_style = 'none'
            top_color = 'FF000000'
            top_automatic = False
            if top_elem is not None:
                top_style = top_elem.get('style', 'none')
                top_color_elem = top_elem.find('main:color', namespaces=self.ns)
                if top_color_elem is not None:
                    top_color = top_color_elem.get('rgb', 'FF000000')
                    top_automatic = top_color_elem.get('auto') in (
                        '1', 'true', 'True'
                    )
            
            # Load bottom border
            bottom_style = 'none'
            bottom_color = 'FF000000'
            bottom_automatic = False
            if bottom_elem is not None:
                bottom_style = bottom_elem.get('style', 'none')
                bottom_color_elem = bottom_elem.find('main:color', namespaces=self.ns)
                if bottom_color_elem is not None:
                    bottom_color = bottom_color_elem.get('rgb', 'FF000000')
                    bottom_automatic = bottom_color_elem.get('auto') in (
                        '1', 'true', 'True'
                    )

            diagonal_style = 'none'
            diagonal_color = 'FF000000'
            if diagonal_elem is not None:
                diagonal_style = diagonal_elem.get('style', 'none')
                diagonal_color_elem = diagonal_elem.find(
                    'main:color', namespaces=self.ns
                )
                if diagonal_color_elem is not None:
                    diagonal_color = diagonal_color_elem.get(
                        'rgb', 'FF000000'
                    )
            
            border_data = {
                'top': {
                    'style': top_style, 'color': top_color,
                    'automatic_color': top_automatic,
                },
                'bottom': {
                    'style': bottom_style, 'color': bottom_color,
                    'automatic_color': bottom_automatic,
                },
                'left': {
                    'style': left_style, 'color': left_color,
                    'automatic_color': left_automatic,
                },
                'right': {
                    'style': right_style, 'color': right_color,
                    'automatic_color': right_automatic,
                },
                'diagonal': {
                    'style': diagonal_style,
                    'color': diagonal_color,
                },
                'diagonal_up': border_elem.get('diagonalUp') in ('1', 'true'),
                'diagonal_down': border_elem.get('diagonalDown') in ('1', 'true'),
            }
            self.workbook._border_styles[i] = border_data
    
    def _load_cell_xfs(self, styles_root):
        """
        Loads cell XF records from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        if not hasattr(self.workbook, '_cell_xf_by_index'):
            self.workbook._cell_xf_by_index = {}

        # Ensure style index 0 remains the ECMA-376 defaults.
        # Without this, the first non-default alignment/protection parsed from
        # cellXfs can accidentally take index 0 and be applied to default cells.
        if 0 not in self.workbook._alignment_styles:
            self.workbook._alignment_styles[0] = {
                'horizontal': 'general',
                'vertical': 'bottom',
                'wrap_text': False,
                'indent': 0,
                'text_rotation': 0,
                'shrink_to_fit': False,
                'reading_order': 0,
                'relative_indent': 0
            }
        if 0 not in self.workbook._protection_styles:
            self.workbook._protection_styles[0] = {
                'locked': True,
                'hidden': False
            }

        cell_xfs = styles_root.findall('.//main:cellXfs/main:xf', namespaces=self.ns)
        for i, xf_elem in enumerate(cell_xfs):
            font_idx = int(xf_elem.get('fontId', 0))
            fill_idx = int(xf_elem.get('fillId', 0))
            border_idx = int(xf_elem.get('borderId', 0))
            num_fmt_idx = int(xf_elem.get('numFmtId', 0))
            
            # Load alignment if present
            alignment_idx = 0
            alignment_elem = xf_elem.find('main:alignment', namespaces=self.ns)
            if alignment_elem is not None:
                # Check if this alignment already exists
                horizontal = alignment_elem.get('horizontal', 'general')
                vertical = alignment_elem.get('vertical', 'bottom')
                text_rotation = int(alignment_elem.get('textRotation', 0))
                wrap_text = alignment_elem.get('wrapText') == '1'
                shrink_to_fit = alignment_elem.get('shrinkToFit') == '1'
                indent = int(alignment_elem.get('indent', 0))
                reading_order = int(alignment_elem.get('readingOrder', 0))
                relative_indent = int(alignment_elem.get('relativeIndent', 0))
                
                # Check if this alignment already exists
                for idx, align_data in self.workbook._alignment_styles.items():
                    if (align_data['horizontal'] == horizontal and
                        align_data['vertical'] == vertical and
                        align_data['wrap_text'] == wrap_text and
                        align_data['indent'] == indent and
                        align_data['text_rotation'] == text_rotation and
                        align_data['shrink_to_fit'] == shrink_to_fit and
                        align_data['reading_order'] == reading_order and
                        align_data['relative_indent'] == relative_indent):
                        alignment_idx = idx
                        break
                
                # If not found, create new alignment style
                if alignment_idx == 0:
                    alignment_idx = len(self.workbook._alignment_styles)
                    self.workbook._alignment_styles[alignment_idx] = {
                        'horizontal': horizontal,
                        'vertical': vertical,
                        'wrap_text': wrap_text,
                        'indent': indent,
                        'text_rotation': text_rotation,
                        'shrink_to_fit': shrink_to_fit,
                        'reading_order': reading_order,
                        'relative_indent': relative_indent
                    }

            # Load protection if present
            protection_idx = 0
            protection_elem = xf_elem.find('main:protection', namespaces=self.ns)
            if protection_elem is not None:
                # Get protection attributes (default: locked=1, hidden=0)
                locked = protection_elem.get('locked', '1') == '1'
                hidden = protection_elem.get('hidden', '0') == '1'

                # Check if this protection already exists
                for idx, prot_data in self.workbook._protection_styles.items():
                    if (prot_data['locked'] == locked and
                        prot_data['hidden'] == hidden):
                        protection_idx = idx
                        break

                # If not found, create new protection style
                if protection_idx == 0 and not (locked and not hidden):  # Skip if it's the default
                    protection_idx = len(self.workbook._protection_styles)
                    self.workbook._protection_styles[protection_idx] = {
                        'locked': locked,
                        'hidden': hidden
                    }

            cell_style_key = (font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx)
            self.workbook._cell_styles[cell_style_key] = i  # Store the actual cellXf index
            self.workbook._cell_xf_by_index[i] = cell_style_key

    def _load_dxf_styles(self, styles_root):
        """
        Loads differential formatting (dxf) styles from styles XML.

        These are used for conditional formatting.

        Args:
            styles_root: The XML root element of styles.
        """
        self.workbook._dxf_styles = []

        dxfs_elem = styles_root.find('.//main:dxfs', namespaces=self.ns)
        if dxfs_elem is None:
            return

        for dxf_elem in dxfs_elem.findall('main:dxf', namespaces=self.ns):
            dxf_data = {}

            # Load font
            font_elem = dxf_elem.find('main:font', namespaces=self.ns)
            if font_elem is not None:
                font_data = {}
                b_elem = font_elem.find('main:b', namespaces=self.ns)
                if b_elem is not None:
                    font_data['bold'] = b_elem.get('val', '1') != '0'
                i_elem = font_elem.find('main:i', namespaces=self.ns)
                if i_elem is not None:
                    font_data['italic'] = i_elem.get('val', '1') != '0'
                u_elem = font_elem.find('main:u', namespaces=self.ns)
                if u_elem is not None:
                    font_data['underline'] = True
                strike_elem = font_elem.find('main:strike', namespaces=self.ns)
                if strike_elem is not None:
                    font_data['strikethrough'] = True
                color_elem = font_elem.find('main:color', namespaces=self.ns)
                if color_elem is not None:
                    font_data['color'] = self._resolve_dxf_color(
                        color_elem, 'FF000000'
                    )
                if font_data:
                    dxf_data['font'] = font_data

            # Load fill
            fill_elem = dxf_elem.find('main:fill', namespaces=self.ns)
            if fill_elem is not None:
                pattern_elem = fill_elem.find('main:patternFill', namespaces=self.ns)
                if pattern_elem is not None:
                    fill_data = {
                        'pattern_type': pattern_elem.get('patternType', 'solid')
                    }
                    fg_elem = pattern_elem.find('main:fgColor', namespaces=self.ns)
                    if fg_elem is not None:
                        fill_data['fg_color'] = self._resolve_dxf_color(
                            fg_elem, 'FFFFFFFF'
                        )
                    bg_elem = pattern_elem.find('main:bgColor', namespaces=self.ns)
                    if bg_elem is not None:
                        background_color = self._resolve_dxf_color(
                            bg_elem, 'FFFFFFFF'
                        )
                        fill_data['bg_color'] = background_color
                        if (
                            fg_elem is None
                            or (
                                fg_elem.get('theme') is not None
                                and bg_elem.get('rgb') is not None
                            )
                        ):
                            # Excel uses a lone DXF background color as the
                            # visible solid color for preset rules. Custom
                            # table styles also retain their concrete RGB in
                            # bgColor while fgColor points at a theme slot.
                            fill_data['fg_color'] = background_color
                    dxf_data['fill'] = fill_data

            # Load border (simplified - just check if any border is present)
            border_elem = dxf_elem.find('main:border', namespaces=self.ns)
            if border_elem is not None:
                borders = {}
                for side in ['left', 'right', 'top', 'bottom', 'horizontal']:
                    side_elem = border_elem.find(f'main:{side}', namespaces=self.ns)
                    if side_elem is not None and side_elem.get('style'):
                        color = 'FF000000'
                        color_elem = side_elem.find('main:color', namespaces=self.ns)
                        if color_elem is not None:
                            color = self._resolve_dxf_color(
                                color_elem, 'FF000000'
                            )
                        borders[side] = {
                            'style': side_elem.get('style'),
                            'color': color,
                        }
                if borders:
                    dxf_data['borders'] = borders
                    first = next(iter(borders.values()))
                    dxf_data['border'] = dict(first)

            self.workbook._dxf_styles.append(dxf_data)

    def _load_table_styles(self, styles_root):
        self.workbook._table_styles = {}
        table_styles = styles_root.find(
            './/main:tableStyles', namespaces=self.ns
        )
        if table_styles is None:
            return
        for style in table_styles.findall('main:tableStyle', namespaces=self.ns):
            name = style.get('name')
            if not name:
                continue
            elements = {}
            for element in style.findall(
                'main:tableStyleElement', namespaces=self.ns
            ):
                try:
                    dxf_id = int(element.get('dxfId'))
                except (TypeError, ValueError):
                    continue
                elements[element.get('type')] = {
                    'dxf_id': dxf_id,
                    'size': int(element.get('size', '1')),
                }
            self.workbook._table_styles[name] = elements

    def _resolve_dxf_color(self, element, default):
        color = self._resolve_color_element(element) or default
        tint_value = element.get('tint')
        if tint_value is None:
            return color
        try:
            tint = float(tint_value)
            rgb = color[-6:]
            channels = [int(rgb[index:index + 2], 16) for index in (0, 2, 4)]
        except (TypeError, ValueError):
            return color
        if tint >= 0:
            channels = [round(value + (255 - value) * tint) for value in channels]
        else:
            channels = [round(value * (1.0 + tint)) for value in channels]
        return 'FF' + ''.join(f'{max(0, min(255, value)):02X}' for value in channels)

    def _load_document_properties(self, zipf):
        """
        Loads document properties from docProps/core.xml and docProps/app.xml.

        ECMA-376 Part 2, Section 11 - Core Properties
        ECMA-376 Part 1, Section 22.2 - Extended Properties

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        # Load core properties
        self._load_core_properties(zipf)

        # Load extended/app properties
        self._load_app_properties(zipf)

    def _load_core_properties(self, zipf):
        """
        Loads core document properties from docProps/core.xml.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            core_xml_content = zipf.read('docProps/core.xml')
            core_root = ET.fromstring(core_xml_content)

            # Namespaces for core properties
            ns = {
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dc': 'http://purl.org/dc/elements/1.1/',
                'dcterms': 'http://purl.org/dc/terms/'
            }

            # Ensure document_properties exists
            if not hasattr(self.workbook, 'document_properties') or self.workbook.document_properties is None:
                from .document_properties import DocumentProperties
                self.workbook._document_properties = DocumentProperties()

            core = self.workbook.document_properties.core

            # Load Dublin Core properties
            title = core_root.find('dc:title', ns)
            if title is not None and title.text:
                core._title = title.text

            subject = core_root.find('dc:subject', ns)
            if subject is not None and subject.text:
                core._subject = subject.text

            creator = core_root.find('dc:creator', ns)
            if creator is not None and creator.text:
                core._creator = creator.text

            description = core_root.find('dc:description', ns)
            if description is not None and description.text:
                core._description = description.text

            # Load OPC Core Properties
            keywords = core_root.find('cp:keywords', ns)
            if keywords is not None and keywords.text:
                core._keywords = keywords.text

            last_modified_by = core_root.find('cp:lastModifiedBy', ns)
            if last_modified_by is not None and last_modified_by.text:
                core._last_modified_by = last_modified_by.text

            revision = core_root.find('cp:revision', ns)
            if revision is not None and revision.text:
                core._revision = revision.text

            category = core_root.find('cp:category', ns)
            if category is not None and category.text:
                core._category = category.text

            content_status = core_root.find('cp:contentStatus', ns)
            if content_status is not None and content_status.text:
                core._content_status = content_status.text

            # Load dates
            created = core_root.find('dcterms:created', ns)
            if created is not None and created.text:
                core._created = self._parse_datetime(created.text)

            modified = core_root.find('dcterms:modified', ns)
            if modified is not None and modified.text:
                core._modified = self._parse_datetime(modified.text)

        except KeyError:
            # docProps/core.xml not found, skip
            pass

    def _load_app_properties(self, zipf):
        """
        Loads extended/application properties from docProps/app.xml.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            app_xml_content = zipf.read('docProps/app.xml')
            app_root = ET.fromstring(app_xml_content)

            # Ensure document_properties exists
            if not hasattr(self.workbook, 'document_properties') or self.workbook.document_properties is None:
                from .document_properties import DocumentProperties
                self.workbook._document_properties = DocumentProperties()

            ext = self.workbook.document_properties.extended

            # Load properties (note: default namespace, so no prefix)
            application = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application')
            if application is not None and application.text:
                ext._application = application.text

            app_version = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}AppVersion')
            if app_version is not None and app_version.text:
                ext._app_version = app_version.text

            company = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company')
            if company is not None and company.text:
                ext._company = company.text

            manager = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Manager')
            if manager is not None and manager.text:
                ext._manager = manager.text

            hyperlink_base = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}HyperlinkBase')
            if hyperlink_base is not None and hyperlink_base.text:
                ext._hyperlink_base = hyperlink_base.text

            doc_security = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}DocSecurity')
            if doc_security is not None and doc_security.text:
                ext._doc_security = int(doc_security.text)

            scale_crop = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}ScaleCrop')
            if scale_crop is not None and scale_crop.text:
                ext._scale_crop = scale_crop.text.lower() == 'true'

            links_up_to_date = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}LinksUpToDate')
            if links_up_to_date is not None and links_up_to_date.text:
                ext._links_up_to_date = links_up_to_date.text.lower() == 'true'

            shared_doc = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}SharedDoc')
            if shared_doc is not None and shared_doc.text:
                ext._shared_doc = shared_doc.text.lower() == 'true'

        except KeyError:
            # docProps/app.xml not found, skip
            pass

    def _parse_datetime(self, date_str):
        """
        Parses a W3CDTF datetime string.

        Args:
            date_str: A datetime string in W3CDTF format (e.g., '2024-01-15T10:30:00Z')

        Returns:
            datetime object or the original string if parsing fails.
        """
        from datetime import datetime

        try:
            # Try ISO format with Z suffix
            if date_str.endswith('Z'):
                return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            # Try ISO format
            return datetime.fromisoformat(date_str)
        except (ValueError, AttributeError):
            # Return the string if parsing fails
            return date_str
