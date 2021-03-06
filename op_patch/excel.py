from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Read an xlsx file into Python"""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
from sys import exc_info
from io import BytesIO
import os.path
import warnings

# compatibility imports
from openpyxl.compat import unicode, file

# package imports
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CORE,
    ARC_WORKBOOK,
    ARC_STYLE,
    ARC_THEME,
    SHARED_STRINGS,
    EXTERNAL_LINK,
    XLTM,
    XLTX,
)

from openpyxl.workbook import Workbook
from openpyxl.workbook.names.external import detect_external_links
from openpyxl.workbook.names.named_range import read_named_ranges
from openpyxl.packaging.relationship import get_dependents
from openpyxl.reader.strings import read_string_table
from openpyxl.reader.style import read_style_table
from openpyxl.reader.workbook import (
    read_content_types,
    read_excel_base_date,
    detect_worksheets,
    read_rels,
    read_workbook_code_name,
    read_workbook_settings,
)
from openpyxl.workbook.properties import read_properties, DocumentProperties
from openpyxl.worksheet.read_only import ReadOnlyWorksheet
from openpyxl.reader.worksheet import WorkSheetParser
from openpyxl.comments.reader import read_comments, get_comments_file
# Use exc_info for Python 2 compatibility with "except Exception[,/ as] e"

from .drawings import read_drawings, get_drawings_file

CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'
SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')


def repair_central_directory(zipFile, is_file_instance):
    ''' trims trailing data from the central directory
    code taken from http://stackoverflow.com/a/7457686/570216, courtesy of Uri Cohen
    '''

    f = zipFile if is_file_instance else open(zipFile, 'rb+')
    data = f.read()
    pos = data.find(CENTRAL_DIRECTORY_SIGNATURE)  # End of central directory signature
    if (pos > 0):
        sio = BytesIO(data)
        sio.seek(pos + 22)  # size of 'ZIP end of central directory record'
        sio.truncate()
        sio.seek(0)
        return sio

    f.seek(0)
    return f



def _validate_archive(filename):
    """
    Check the file is a valid zipfile
    """
    is_file_like = hasattr(filename, 'read')

    if not is_file_like and os.path.isfile(filename):
        file_format = os.path.splitext(filename)[-1]
        if file_format not in SUPPORTED_FORMATS:
            if file_format == '.xls':
                msg = ('openpyxl does not support the old .xls file format, '
                       'please use xlrd to read this file, or convert it to '
                       'the more recent .xlsx file format.')
            elif file_format == '.xlsb':
                msg = ('openpyxl does not support binary format .xlsb, '
                       'please convert this file to .xlsx format if you want '
                       'to open it with openpyxl')
            else:
                msg = ('openpyxl does not support %s file format, '
                       'please check you can open '
                       'it with Excel first. '
                       'Supported formats are: %s') % (file_format,
                                                       ','.join(SUPPORTED_FORMATS))
            raise InvalidFileException(msg)


    if is_file_like:
        # fileobject must have been opened with 'rb' flag
        # it is required by zipfile
        if getattr(filename, 'encoding', None) is not None:
            raise IOError("File-object must be opened in binary mode")

    try:
        archive = ZipFile(filename, 'r', ZIP_DEFLATED)
    except BadZipfile:
        f = repair_central_directory(filename, is_file_like)
        archive = ZipFile(f, 'r', ZIP_DEFLATED)
    return archive


def load_workbook(filename, read_only=False, use_iterators=False, keep_vba=False, guess_types=False, data_only=False):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`

    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool

    :param use_iterators: use lazy load for cells
    :type use_iterators: bool

    :param keep_vba: preseve vba content (this does NOT mean you can use it)
    :type keep_vba: bool

    :param guess_types: guess cell content type and do not read it from the file
    :type guess_types: bool

    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool

    :rtype: :class:`openpyxl.workbook.Workbook`

    .. note::

        When using lazy load, all worksheets will be :class:`openpyxl.worksheet.iter_worksheet.IterableWorksheet`
        and the returned workbook will be read-only.

    """
    archive = _validate_archive(filename)
    read_only = read_only or use_iterators

    wb = Workbook(guess_types=guess_types, data_only=data_only, read_only=read_only)

    if read_only and guess_types:
        warnings.warn('Data types are not guessed when using iterator reader')

    valid_files = archive.namelist()

    # If are going to preserve the vba then attach a copy of the archive to the
    # workbook so that is available for the save.
    if keep_vba:
        try:
            f = open(filename, 'rb')
            s = f.read()
            f.close()
        except:
            pos = filename.tell()
            filename.seek(0)
            s = filename.read()
            filename.seek(pos)
        wb.vba_archive = ZipFile(BytesIO(s), 'r')

    if read_only:
        wb._archive = ZipFile(filename)

    # get workbook-level information
    try:
        wb.properties = read_properties(archive.read(ARC_CORE))
    except KeyError:
        wb.properties = DocumentProperties()
    wb.active = read_workbook_settings(archive.read(ARC_WORKBOOK)) or 0

    # what content types do we have?
    cts = dict(read_content_types(archive))

    strings_path = cts.get(SHARED_STRINGS)
    if strings_path is not None:
        if strings_path.startswith("/"):
            strings_path = strings_path[1:]
        shared_strings = read_string_table(archive.read(strings_path))
    else:
        shared_strings = []

    wb.is_template = XLTX in cts or XLTM in cts

    try:
        wb.loaded_theme = archive.read(ARC_THEME)  # some writers don't output a theme, live with it (fixes #160)
    except KeyError:
        assert wb.loaded_theme == None, "even though the theme information is missing there is a theme object ?"

    parsed_styles = read_style_table(archive)
    if parsed_styles is not None:
        wb._differential_styles = parsed_styles.differential_styles
        wb._cell_styles = parsed_styles.cell_styles
        wb._named_styles = parsed_styles.named_styles
        wb._colors = parsed_styles.color_index
        wb._borders = parsed_styles.border_list
        wb._fonts = parsed_styles.font_list
        wb._fills = parsed_styles.fill_list
        wb._number_formats = parsed_styles.number_formats
        wb._protections = parsed_styles.protections
        wb._alignments = parsed_styles.alignments
        wb._colors = parsed_styles.color_index

    wb.excel_base_date = read_excel_base_date(archive)

    # get worksheets
    wb._sheets = []  # remove preset worksheet
    for sheet in detect_worksheets(archive):
        sheet_name = sheet['title']
        worksheet_path = sheet['path']
        if not worksheet_path in valid_files:
            continue

        if read_only:
            new_ws = ReadOnlyWorksheet(wb, sheet_name, worksheet_path, None,
                                       shared_strings)
            wb._add_sheet(new_ws)
        else:
            fh = archive.open(worksheet_path)
            parser = WorkSheetParser(wb, sheet_name, fh, shared_strings)
            parser.parse()
            new_ws = wb[sheet_name]
        new_ws.sheet_state = sheet['state']

        if wb.vba_archive is not None and new_ws.legacy_drawing is not None:
            # We need to get the file name of the legacy drawing
            dirname, basename = worksheet_path.rsplit('/', 1)
            rels_path = '/'.join((dirname, '_rels', basename + '.rels'))
            rels = get_dependents(archive, rels_path)
            new_ws.legacy_drawing = rels[new_ws.legacy_drawing].target

        if not read_only:
        # load comments into the worksheet cells
            comments_file = get_comments_file(worksheet_path, archive, valid_files)
            if comments_file is not None:
                read_comments(new_ws, archive.read(comments_file))
            drawings_file = get_drawings_file(worksheet_path, archive, valid_files)
            if drawings_file is not None:
                read_drawings(new_ws, drawings_file, archive, valid_files)
    wb._differential_styles = [] # reset
    wb._named_ranges = list(read_named_ranges(archive.read(ARC_WORKBOOK), wb))

    wb.code_name = read_workbook_code_name(archive.read(ARC_WORKBOOK))

    if EXTERNAL_LINK in cts:
        rels = read_rels(archive)
        wb._external_links = list(detect_external_links(rels, archive))


    archive.close()
    return wb
