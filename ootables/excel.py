import re
import string
import zipfile as zf

from ootables import core


def col_to_int(s):
    return int(''.join([str(string.ascii_uppercase.index(c)) for c in s]))


class ExcelRange:
    def __init__(self, range):
        self.__range = range
        self.__set_props(range)

    @property
    def range(self):
        return self.__range

    @property
    def start(self):
        return self.__start

    @property
    def end(self):
        return self.__end

    @property
    def left_bound(self):
        return self.__start_col

    @property
    def left_bound_n(self):
        return self.__start_col_n

    @property
    def right_bound(self):
        return self.__end_col

    @property
    def right_bound_n(self):
        return self.__end_col_n

    @property
    def upper_bound(self):
        return self.__start_row

    @property
    def lower_bound(self):
        return self.__end_row

    def __parse_loc(self, s):
        col, row = str(), str()
        for c in s:
            try:
                int(c)
                row += c
            except ValueError:
                col += c
        return col, row

    def __set_props(self, r):
        self.__start, self.__end = r.split(':')
        self.__start_col, self.__start_row = self.__parse_loc(self.__start)
        self.__end_col, self.__end_row = self.__parse_loc(self.__end)
        self.__start_col_n = col_to_int(self.__start_col)
        self.__end_col_n = col_to_int(self.__end_col)


class ExcelCell:
    def __init__(self, index, value):    # , type=None):
        self.__index = index
        self.__set_row_col(index)
        # self.__type = type
        self.__set_value(value)

    def __repr__(self):
        return f"ExcelCell(index='{self.__index}', value='{self.__value}')"

    @property
    def index(self):
        return self.__index

    @property
    def row(self):
        return self.__row

    @property
    def col(self):
        return self.__col

    @property
    def col_n(self):
        return self.__col_n

    def __set_row_col(self, s):
        self.__row, self.__col = str(), str()
        for c in s:
            try:
                int(c)
                self.__row += c
            except ValueError:
                self.__col += c
        self.__col_n = col_to_int(self.__col)

    # @property
    # def type(self):
    #     return self.__type

    @property
    def value(self):
        return self.__value

    def __set_value(self, value):
        # Update this to change the type of the value to a Python type?
        self.__value = value


class ExcelRow:
    # An ExcelRow is a list of ExcelCells
    def __init__(self, index, span, cells):
        self.__index = index
        self.__span = span
        self.__cells = cells

    @property
    def index(self):
        return self.__index

    @property
    def span(self):
        return self.__span

    @property
    def cells(self):
        return self.__cells


class ExcelColumn:
    def __init__(self, id, name, cells=list(), values=list()):
        self.__id = id
        self.__name = name
        self.values = values

    @property
    def id(self):
        return self.__id

    @property
    def name(self):
        return self.__name

    @property
    def cells(self):
        return self.__cells

    @cells.setter
    def cells(self, cells):
        self.__cells = cells

    @property
    def values(self):
        return self.__values

    @values.setter
    def values(self, values):
        self.__values = values


class ExcelTable:
    def __init__(self, name, display_name, range, rows, columns):
        self.__name = name
        self.__display_name = display_name
        self.__range = range
        self.__rows = rows
        self.__header = list()
        self.__set_cols(columns)
        self.__set_data()

    @property
    def name(self):
        return self.__name

    @property
    def display_name(self):
        return self.__display_name

    @property
    def range(self):
        return self.__range

    @property
    def header(self):
        return self.__header

    @property
    def rows(self):
        return self.__rows

    @property
    def cols(self):
        return self.__cols

    def __set_cols(self, cols):
        # cols = list of ExcelColumns
        self.__cols = cols
        # How to tell if the table actually has a header? For now, this will
        # set the header
        matches = list()
        for i in range(len(cols)):
            if self.__rows[0].cells[i].value == cols[i].name:
                self.__header.append(cols[i].name)
                matches.append(True)
            else:
                self.__header.append(f'Column{i}')
                matches.append(False)
        if sum(matches) == len(matches):
            self.__first_row = 1
        else:
            self.__first_row = 0

        # For each column, get its data from the rows
        col_cells = dict()
        col_values = dict()
        for i in range(len(self.__header)):
            col_cells[cols[i].name] = list()
            col_values[cols[i].name] = list()
        for r in self.__rows[self.__first_row:]:
            for i in range(len(r.cells)):
                col_cells[cols[i].name].append(r.cells[i])
                col_values[cols[i].name].append(r.cells[i].value)
        for c in cols:
            c.cells = col_cells[c.name]
            c.values = col_values[c.name]

    @property
    def data(self):
        return self.__data

    def __set_data(self):
        self.__data = list()
        for r in self.__rows[self.__first_row:]:
            row_dict = dict()
            for i in range(len(r.cells)):
                row_dict[self.__header[i]] = r.cells[i].value
            self.__data.append(row_dict)


class SharedString:
    def __init__(self, value):
        self.__v = value

    def __repr__(self):
        return f"SharedString('{self.__v}')"

    def __str__(self):
        return self.__v

    @property
    def value(self):
        return self.__v


class Sheet(core.CoreObject):
    def __init__(self, oofile, id, name, target, strings):
        d, f = target.split('/')
        self.__obj_file = f'xl/{target}'
        self.__rel_file = f'xl/{d}/_rels/{f}.rels'
        super().__init__(oofile, self.__obj_file, self.__rel_file)
        with self._oofile.open(self.__obj_file) as f:
            self.__xml_doc = core.XMLDoc(f.read())
        self.__id = id
        self.__name = name
        self.__strings = strings
        self.__set_data(core.get_elements(self.__xml_doc, 'row'))
        self.__set_tables(core.get_elements(self.__xml_doc, 'tablePart'))

    def __repr__(self):
        return f"Sheet(name='{self.__name}')"

    @property
    def id(self):
        return self.__id

    @property
    def name(self):
        return self.__name

    @property
    def data(self):
        return self.__data

    def __set_data(self, row_els):
        self.__data = list()
        for e in row_els:
            cells = list()
            for c in e.getchildren():
                value = c.getchildren()[0].text
                if c.get('t') == 's':
                    value = self.__strings[int(value)].value
                cells.append(ExcelCell(index=c.get('r'), value=value))
            self.__data.append(ExcelRow(e.get('r'), e.get('spans'), cells))

    @property
    def tables(self):
        return self.__tables

    def __set_tables(self, table_elements):
        self.__tables = list()
        for e in table_elements:
            rel_id_key = list(filter(
                lambda k: re.match('\{.*\}id', k),
                e.keys()
            ))[0]
            t_rel = list(filter(
                lambda r: r.id == e.get(rel_id_key),
                self._rels
            ))[0]
            with self._oofile.open(
                    f"xl/tables/{t_rel.target.split('/')[-1]}") as f:
                t_xml_doc = core.XMLDoc(f.read())
            t_el = core.get_elements(t_xml_doc, 'table')[0]
            # Find the range of the table here
            # Add the cells in that range to the Table, all the info about the
            # table is provided to the Table() when initialized. The Table
            # class should not be anything more than a home-grown API that
            # provides a consistent interface regardless of where the data
            # comes from. In other words, from the perspective of OOTables, it
            # doesn't matter if the table came from Excel, Word, or PowerPoint
            # because each one will have the same properties and methods
            t_rng = ExcelRange(t_el.get('ref'))
            t_data = list()
            for r in self.__data:
                t_row_cells = list()
                if r.index >= t_rng.upper_bound \
                        and r.index <= t_rng.lower_bound:
                    for c in r.cells:
                        if c.col_n >= t_rng.left_bound_n \
                                and c.col_n <= t_rng.right_bound_n:
                            t_row_cells.append(c)
                        else:
                            continue
                    t_data.append(ExcelRow(
                        r.index, r.span, t_row_cells
                    ))
                else:
                    continue

            # Get the columns
            t_cols = list(filter(
                lambda e: re.match('\{.*\}tableColumns$', e.tag),
                t_el
            ))[0].getchildren()
            t_cols = [
                ExcelColumn(c.get('id'), c.get('name'))
                for c in t_cols
            ]
            self.__tables.append(ExcelTable(
                t_el.get('name'), t_el.get('displayName'), t_el.get('ref'),
                t_data, t_cols
            ))


class Book(core.CoreObject):
    def __init__(self, filename):
        self.__filename = filename
        self.__xml_path = 'xl/workbook.xml'
        self.__rel_xml_path = 'xl/_rels/workbook.xml.rels'
        super().__init__(
            zf.ZipFile(filename), self.__xml_path, self.__rel_xml_path
        )
        with self._oofile.open('xl/workbook.xml') as f:
            self.__xml_doc = core.XMLDoc(f.read())

        shared_str_rels = list(filter(
            lambda r: r.type == 'sharedStrings', self._rels
        ))
        if len(shared_str_rels) > 0:
            self.__set_shared_strings(shared_str_rels)

        self.__set_sheets()

    def __repr__(self):
        return f"Book(name='{self.__filename}')"

    @property
    def filename(self):
        return self.__filename

    @property
    def xml_doc(self):
        return self.__xml_doc

    @property
    def shared_strings(self):
        return self.__shared_strings

    def __set_shared_strings(self, rels):
        with self._oofile.open(f'xl/{rels[0].target}') as f:
            doc = core.XMLDoc(f.read())
        self.__shared_strings = tuple([
            SharedString(s.text)
            for s in core.get_elements(doc, 't')
        ])

    @property
    def sheets(self):
        return self.__sheets

    def __set_sheets(self):
        self.__sheets = list()
        # with self._oofile.open(self.__xml_path) as f:
        #     doc = core.XMLDoc(f.read())
        sheets = core.get_elements(self.__xml_doc, 'sheet')
        for i in range(len(sheets)):
            key = list(filter(
                lambda k: re.match('\{.*\}id$', k), sheets[i].keys()
            ))[0]
            rel_id = sheets[i].get(key)
            target = list(filter(
                    lambda r: rel_id == r.id, self._rels
            ))[0].target
            sheet = Sheet(
                self._oofile, sheets[i].get('sheetId'),
                sheets[i].get('name'), target, self.__shared_strings
            )
            self.__sheets.append(sheet)
