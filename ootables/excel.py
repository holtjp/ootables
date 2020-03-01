import re
import zipfile as zf

from ootables import core


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

    def __repr__(self):
        return f"Sheet(name='{self.__name}')"

    @property
    def id(self):
        return self.__id

    @property
    def name(self):
        return self.__name


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
        with self._oofile.open(self.__xml_path) as f:
            doc = core.XMLDoc(f.read())
        sheets = core.get_elements(doc, 'sheet')
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
