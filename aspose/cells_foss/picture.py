"""
Aspose.Cells for Python - Picture Module

Provides worksheet picture APIs and picture collection support.
"""

import os


class Picture:
    """Represents a worksheet picture anchored to cells."""

    def __init__(self, worksheet, image_bytes, image_extension, upper_left_row, upper_left_column, lower_right_row, lower_right_column, name=None):
        self._worksheet = worksheet
        self._image_bytes = image_bytes
        self._image_extension = (image_extension or "png").lower().lstrip(".")
        self._upper_left_row = int(upper_left_row)
        self._upper_left_column = int(upper_left_column)
        self._lower_right_row = int(lower_right_row)
        self._lower_right_column = int(lower_right_column)
        self._upper_left_row_offset = 0
        self._upper_left_column_offset = 0
        self._lower_right_row_offset = 0
        self._lower_right_column_offset = 0
        self._name = name
        self._source_part_path = None
        self._source_content_type = None
        # Extended picture attributes
        self._hyperlink_url = None          # URL for <a:hlinkClick> click hyperlink
        self._edit_as = None                # twoCellAnchor editAs: "twoCell", "oneCell", "absolute"
        self._no_change_aspect = True       # <a:picLocks noChangeAspect="1"/>
        self._source_cNvPr_extLst_xml = None   # raw XML string for extLst inside <xdr:cNvPr>
        self._source_blip_extLst_xml = None    # raw XML string for extLst inside <a:blip>
        self._source_spPr_xml = None           # raw XML string for inner content of <xdr:spPr>
        self._transform_x = None
        self._transform_y = None
        self._transform_width = None
        self._transform_height = None

    @property
    def name(self):
        return self._name

    @property
    def image_extension(self):
        return self._image_extension

    @property
    def image_bytes(self):
        return self._image_bytes

    @property
    def hyperlink_url(self):
        """URL for the picture's click hyperlink (<a:hlinkClick>)."""
        return self._hyperlink_url

    @hyperlink_url.setter
    def hyperlink_url(self, value):
        self._hyperlink_url = value

    @property
    def edit_as(self):
        """Anchor editAs mode: 'twoCell', 'oneCell', or 'absolute'. None means 'twoCell' (default)."""
        return self._edit_as

    @edit_as.setter
    def edit_as(self, value):
        self._edit_as = value

    @property
    def no_change_aspect(self):
        """Whether the picture's aspect ratio is locked (<a:picLocks noChangeAspect='1'/>)."""
        return self._no_change_aspect

    @no_change_aspect.setter
    def no_change_aspect(self, value):
        self._no_change_aspect = bool(value)

    def copy(self, worksheet):
        pic = Picture(
            worksheet,
            self._image_bytes,
            self._image_extension,
            self._upper_left_row,
            self._upper_left_column,
            self._lower_right_row,
            self._lower_right_column,
            name=self._name,
        )
        pic._upper_left_row_offset = self._upper_left_row_offset
        pic._upper_left_column_offset = self._upper_left_column_offset
        pic._lower_right_row_offset = self._lower_right_row_offset
        pic._lower_right_column_offset = self._lower_right_column_offset
        pic._source_part_path = self._source_part_path
        pic._source_content_type = self._source_content_type
        pic._hyperlink_url = self._hyperlink_url
        pic._edit_as = self._edit_as
        pic._no_change_aspect = self._no_change_aspect
        pic._source_cNvPr_extLst_xml = self._source_cNvPr_extLst_xml
        pic._source_blip_extLst_xml = self._source_blip_extLst_xml
        pic._source_spPr_xml = self._source_spPr_xml
        pic._transform_x = self._transform_x
        pic._transform_y = self._transform_y
        pic._transform_width = self._transform_width
        pic._transform_height = self._transform_height
        return pic


class PictureCollection:
    """Collection of pictures in a worksheet."""

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._pictures = []

    def add(self, image_path, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        with open(image_path, "rb") as f:
            image_bytes = f.read()
        extension = os.path.splitext(str(image_path))[1].lower().lstrip(".") or "png"
        name = os.path.basename(str(image_path))
        pic = Picture(
            self._worksheet,
            image_bytes,
            extension,
            upper_left_row,
            upper_left_column,
            lower_right_row,
            lower_right_column,
            name=name,
        )
        self._pictures.append(pic)
        self._worksheet._mark_drawing_dirty()
        return len(self._pictures) - 1

    def Add(self, image_path, upper_left_row, upper_left_column, lower_right_row, lower_right_column):
        """PascalCase alias of add()."""
        return self.add(image_path, upper_left_row, upper_left_column, lower_right_row, lower_right_column)

    def _add_loaded(self, image_bytes, image_extension, upper_left_row, upper_left_column, lower_right_row, lower_right_column, name=None):
        pic = Picture(
            self._worksheet,
            image_bytes,
            image_extension,
            upper_left_row,
            upper_left_column,
            lower_right_row,
            lower_right_column,
            name=name,
        )
        self._pictures.append(pic)
        return pic

    @property
    def count(self):
        return len(self._pictures)

    def __len__(self):
        return len(self._pictures)

    def __iter__(self):
        return iter(self._pictures)

    def __getitem__(self, index):
        return self._pictures[index]

    def copy(self, worksheet):
        cloned = PictureCollection(worksheet)
        cloned._pictures = [pic.copy(worksheet) for pic in self._pictures]
        return cloned
