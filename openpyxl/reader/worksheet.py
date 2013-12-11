# file openpyxl/reader/worksheet.py

# Copyright (c) 2010-2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

"""Reader for a single worksheet."""

# Python stdlib imports

# compatibility imports
from openpyxl.shared.compat import BytesIO, StringIO
from openpyxl.shared.compat import iterparse

# package imports
from openpyxl.cell import get_column_letter
from openpyxl.shared.xmltools import fromstring
from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.worksheet import Worksheet, ColumnDimension, RowDimension
from openpyxl.shared.ooxml import SHEET_MAIN_NS
from openpyxl.style import Color
from openpyxl.styles.formatting import ConditionalFormatting

def _get_xml_iter(xml_source):

    if not hasattr(xml_source, 'name'):
        if hasattr(xml_source, 'decode'):
            return BytesIO(xml_source)
        else:
            return BytesIO(xml_source.encode('utf-8'))
    else:
        xml_source.seek(0)
        return xml_source

def read_dimension(xml_source):

    source = _get_xml_iter(xml_source)

    it = iterparse(source)

    smax_col = None
    smax_row = None
    smin_col = None
    smin_row = None

    for event, element in it:

        if element.tag == '{%s}dimension' % SHEET_MAIN_NS:
            ref = element.get('ref')

            if ':' in ref:
                min_range, max_range = ref.split(':')
            else:
                min_range = max_range = ref

            min_col, min_row = coordinate_from_string(min_range)
            max_col, max_row = coordinate_from_string(max_range)

            return min_col, min_row, max_col, max_row

        if element.tag == '{%s}c' % SHEET_MAIN_NS:
            # Supposedly the dimension is mandatory, but in practice it can be
            # left off sometimes, if so, observe the max/min extants and return
            # those instead.
            col, row = coordinate_from_string(element.get('r'))
            if smin_row is None:
                # initialize the observed max/min values
                smin_col = smax_col = col
                smin_row = smax_row = row
            else:
                # Keep track of the seen max and min (fallback if there's no dimension)
                smin_col = min(smin_col, col)
                smin_row = min(smin_row, row)
                smax_col = max(smax_col, col)
                smax_row = max(smax_row, row)
        else:
            element.clear()

    return smin_col, smin_row, smax_col, smax_row

def filter_cells(pair):
    (event, element) = pair

    return element.tag == '{%s}c' % SHEET_MAIN_NS


class WorkSheetParser(object):

    def __init__(self, ws, xml_source, string_table, style_table, color_index=None):
        self.ws = ws
        self.source = xml_source
        self.string_table = string_table
        self.style_table = style_table
        self.color_index = color_index
        self.guess_types = ws.parent._guess_types
        self.data_only = ws.parent.data_only

    def parse(self):
        stream = _get_xml_iter(self.source)
        it = iterparse(stream)

        dispatcher = {
            '{%s}c' % SHEET_MAIN_NS: self.parse_cell,
            '{%s}mergeCells' % SHEET_MAIN_NS: self.parse_merge,
            '{%s}cols' % SHEET_MAIN_NS: self.parse_column_dimensions,
            '{%s}sheetData' % SHEET_MAIN_NS: self.parse_row_dimensions,
            '{%s}printOptions' % SHEET_MAIN_NS: self.parse_print_options,
            '{%s}pageMargins' % SHEET_MAIN_NS: self.parse_margins,
            '{%s}pageSetup' % SHEET_MAIN_NS: self.parse_page_setup,
            '{%s}headerFooter' % SHEET_MAIN_NS: self.parse_header_footer,
            '{%s}conditionalFormatting' % SHEET_MAIN_NS: self.parser_conditional_formatting
                      }
        for event, element in it:
            tag_name = element.tag
            if tag_name in dispatcher:
                dispatcher[tag_name](element)


    def parse_cell(self, element):
        value = element.findtext('{%s}v' % SHEET_MAIN_NS)
        formula = element.find('{%s}f' % SHEET_MAIN_NS)

        coordinate = element.get('r')
        style_id = element.get('s')
        if style_id is not None:
            self.ws._styles[coordinate] = self.style_table.get(int(style_id))

        if value is not None:
            data_type = element.get('t', 'n')
            if data_type == Cell.TYPE_STRING:
                value = self.string_table.get(int(value))
            if formula is not None and not self.data_only:
                if formula.text:
                    value = "=" + str(formula.text)
                else:
                    value = "="
                formula_type = formula.get('t')
                if formula_type:
                    self.ws.formula_attributes[coordinate] = {'t': formula_type}
                    if formula.get('si'):  # Shared group index for shared formulas
                        self.ws.formula_attributes[coordinate]['si'] = formula.get('si')
                    if formula.get('ref'):  # Range for shared formulas
                        self.ws.formula_attributes[coordinate]['ref'] = formula.get('ref')
            if not self.guess_types and formula is None:
                self.ws.cell(coordinate).set_explicit_value(value=value, data_type=data_type)
            else:
                self.ws.cell(coordinate).value = value

        # to avoid memory exhaustion, clear the item after use
        element.clear()


    def parse_merge(self, element):
        for mergeCell in element.findall('{%s}mergeCell' % SHEET_MAIN_NS):
            self.ws.merge_cells(mergeCell.get('ref'))


    def parse_column_dimensions(self, element):
        colNodes = element.findall('{%s}col' % SHEET_MAIN_NS)
        for col in colNodes:
            min = int(col.get('min')) if col.get('min') else 1
            max = int(col.get('max')) if col.get('max') else 1
            # Ignore ranges that go up to the max column 16384.  Columns need to be extended to handle
            # ranges without creating an entry for every single one.
            if max != 16384:
                for colId in range(min, max + 1):
                    column = get_column_letter(colId)
                    width = col.get("width")
                    auto_size = col.get('bestFit') == '1'
                    visible = col.get('hidden') != '1'
                    outline = col.get('outlineLevel') or 0
                    collapsed = col.get('collapsed') == '1'
                    style_index =  self.style_table.get(int(col.get('style', 0)))
                    if column not in self.ws.column_dimensions:
                        new_dim = ColumnDimension(index=column,
                                                  width=width, auto_size=auto_size,
                                                  visible=visible, outline_level=outline,
                                                  collapsed=collapsed, style_index=style_index)
                        self.ws.column_dimensions[column] = new_dim


    def parse_row_dimensions(self, element):
        rowNodes = element.findall('{%s}row' % SHEET_MAIN_NS)
        for row in rowNodes:
            rowId = int(row.get('r'))
            if rowId not in self.ws.row_dimensions:
                self.ws.row_dimensions[rowId] = RowDimension(rowId)
            ht = row.get('ht')
            if ht is not None:
                self.ws.row_dimensions[rowId].height = float(ht)


    def parse_print_options(self, element):
        hc = element.get('horizontalCentered')
        if hc is not None:
            self.ws.page_setup.horizontalCentered = hc
        vc = element.get('verticalCentered')
        if vc is not None:
            self.ws.page_setup.verticalCentered = vc


    def parse_margins(self, element):
        for key in ("left", "right", "top", "bottom", "header", "footer"):
            value = element.get(key)
            if value is not None:
                setattr(self.ws.page_margins, key, float(value))


    def parse_page_setup(self, element):
        for key in ("orientation", "paperSize", "scale", "fitToPage",
                    "fitToHeight", "fitToWidth", "firstPageNumber",
                    "useFirstPageNumber"):
            value = element.get(key)
            if value is not None:
                setattr(self.ws.page_setup, key, value)


    def parse_header_footer(self, element):
        oddHeader = element.find('{%s}oddHeader' % SHEET_MAIN_NS)
        if oddHeader is not None and oddHeader.text is not None:
            self.ws.header_footer.setHeader(oddHeader.text)
        oddFooter = element.find('{%s}oddFooter' % SHEET_MAIN_NS)
        if oddFooter is not None and oddFooter.text is not None:
            self.ws.header_footer.setFooter(oddFooter.text)


    def parser_conditional_formatting(self, element):
        conditionalFormattingNodes = element.findall('{%s}conditionalFormatting' % SHEET_MAIN_NS)
        rules = {}
        for cf in conditionalFormattingNodes:
            if not cf.get('sqref'):
                # Potentially flag - this attribute should always be present.
                continue
            range_string = cf.get('sqref')
            cfRules = cf.findall('{%s}cfRule' % SHEET_MAIN_NS)
            rules[range_string] = []
            for cfRule in cfRules:
                if not cfRule.get('type') or cfRule.get('type') == 'dataBar':
                    # dataBar conditional formatting isn't supported, as it relies on the complex <extLst> tag
                    continue
                rule = {'type': cfRule.get('type')}
                for attr in ConditionalFormatting.rule_attributes:
                    if cfRule.get(attr) is not None:
                        rule[attr] = cfRule.get(attr)

                formula = cfRule.findall('{%s}formula' % SHEET_MAIN_NS)
                for f in formula:
                    if 'formula' not in rule:
                        rule['formula'] = []
                    rule['formula'].append(f.text)

                colorScale = cfRule.find('{%s}colorScale' % SHEET_MAIN_NS)
                if colorScale is not None:
                    rule['colorScale'] = {'cfvo': [], 'color': []}
                    cfvoNodes = colorScale.findall('{%s}cfvo' % SHEET_MAIN_NS)
                    for node in cfvoNodes:
                        cfvo = {}
                        if node.get('type') is not None:
                            cfvo['type'] = node.get('type')
                        if node.get('val') is not None:
                            cfvo['val'] = node.get('val')
                        rule['colorScale']['cfvo'].append(cfvo)
                    colorNodes = colorScale.findall('{%s}color' % SHEET_MAIN_NS)
                    for color in colorNodes:
                        c = Color(Color.BLACK)
                        if color_index and color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                            c.index = color_index[int(color.get('indexed'))]
                        if color.get('theme') is not None:
                            if color.get('tint') is not None:
                                c.index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                            else:
                                c.index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                        elif color.get('rgb'):
                            c.index = color.get('rgb')
                        rule['colorScale']['color'].append(c)

                iconSet = cfRule.find('{%s}iconSet' % SHEET_MAIN_NS)
                if iconSet is not None:
                    rule['iconSet'] = {'cfvo': []}
                    for iconAttr in ConditionalFormatting.icon_attributes:
                        if iconSet.get(iconAttr) is not None:
                            rule['iconSet'][iconAttr] = iconSet.get(iconAttr)
                    cfvoNodes = iconSet.findall('{%s}cfvo' % SHEET_MAIN_NS)
                    for node in cfvoNodes:
                        cfvo = {}
                        if node.get('type') is not None:
                            cfvo['type'] = node.get('type')
                        if node.get('val') is not None:
                            cfvo['val'] = node.get('val')
                        rule['iconSet']['cfvo'].append(cfvo)

                rules[range_string].append(rule)
        if len(rules):
            self.ws.conditional_formatting.setRules(rules)


def fast_parse(ws, xml_source, string_table, style_table, color_index=None):

    parser = WorkSheetParser(ws, xml_source, string_table, style_table, color_index)
    parser.parse()


from openpyxl.reader.iter_worksheet import IterableWorksheet

def read_worksheet(xml_source, parent, preset_title, string_table,
                   style_table, color_index=None, workbook_name=None, sheet_codename=None, keep_vba=False):
    """Read an xml worksheet"""
    if workbook_name and sheet_codename:
        ws = IterableWorksheet(parent, preset_title, workbook_name,
                sheet_codename, xml_source, string_table)
    else:
        ws = Worksheet(parent, preset_title)
        fast_parse(ws, xml_source, string_table, style_table, color_index)
    if keep_vba:
        ws.xml_source = xml_source
    return ws
