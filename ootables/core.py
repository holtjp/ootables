import io
import re

from lxml import etree


class XMLDoc:
    def __init__(self, xml_bytes: bytes):
        self.xml_bytes = xml_bytes
        self.xml = xml_bytes.decode()
        self.element_tree = etree.parse(io.BytesIO(xml_bytes))

    @property
    def xml_bytes(self):
        return self.__bytes_xml

    @xml_bytes.setter
    def xml_bytes(self, value: bytes):
        self.__bytes_xml = value

    @property
    def xml(self):
        return self.__xml

    @xml.setter
    def xml(self, value: str):
        self.__xml = value

    @property
    def element_tree(self):
        return self.__element_tree

    @element_tree.setter
    def element_tree(self, value: etree._ElementTree):
        self.__element_tree = value


def get_elements(doc: XMLDoc, element_name: str) -> list:
    elements = list()
    for el in doc.element_tree.iter():
        if re.match('\{.*\}' + element_name + '$', el.tag):
            elements.append(el)
    return elements


class Relationship:
    def __init__(self, id, type_url, target):
        self.__id = id
        self.__target = target
        self.type = type_url

    def __repr__(self):
        return f"Relationship(id='{self.__id}', type='{self.__type}', " \
            + f"target='{self.__target}')"

    @property
    def id(self):
        return self.__id

    @property
    def target(self):
        return self.__target

    @property
    def type(self):
        return self.__type

    @type.setter
    def type(self, url: str):
        self.__type = url.split('/')[-1]


class CoreObject:
    # To represent objects with relationships (i.e. Workbooks, Worksheets,
    # Documents, Presentations). In the compressed directory, each one of these
    # will have its own XML file and a _rels directory
    #
    # In Excel, the OOCoreObjects are Workbooks and Worksheets.
    # In Word, they are Documents
    # In PowerPoint, they are Presentations and Slides
    #
    # To avoid rewriting code, this must exist for inheritance
    def __init__(self, oofile, xml_path, rel_xml_path):   # application, name):
        self._oofile = oofile

        # self.application = application
        # self.name = name
        # ^^^ Create getters and setters for these to validate user input
        # apps: Excel = xl, Word = word, PowerPoint = ppt
        # names:
        #   workbook, sheetN
        #   document
        #   presentation, slideN
        self.__set_rels(rel_xml_path)

    # @property
    # def application(self):
    #     return self.__app

    # @application.setter
    # def application(self, value):
    #     self.__app = value
    #     # self.__xml_path_prefix = f'{value}/'
    #     # self.__rel_xml_path_prefix = f'{value}/_rels/'

    # @property
    # def name(self):
    #     return self.__name

    # @name.setter
    # def name(self, value):
    #     self.__name = value
    #     # self.__xml_path = f'{self.__xml_path_prefix}{name}.xml'
    #     # self.__rel_xml_path = f'{self.__rel_xml_path_prefix}{name}.xml.rel'

    @property
    def relationships(self):
        return self._rels

    def __set_rels(self, filepath):
        # Open the relationships file and iterate through it to create a list
        # of Relationship objects
        # rels = list()
        with self._oofile.open(filepath) as f:
            doc = XMLDoc(f.read())
        self._rels = [
            Relationship(r.get('Id'), r.get('Type'), r.get('Target'))
            for r in get_elements(doc, 'Relationship')
        ]
