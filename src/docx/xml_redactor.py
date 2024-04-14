from docx import Document
from docx.oxml import parse_xml
from src.exceptions.docx_exceptions import NotSupportedFormat, NumberingIsNotExists
import zipfile
import os
import shutil
import tempfile
import atexit
import xml.etree.ElementTree as ET
from src.docx.enum.schemas import schemas
from typing import Union


class XMLRedactor:
    """
    For correct work this class you should edit xml.etree.ElementTree
    replace 862 line on : "return qnames, _namespace_map"
    Else DOCX-file after editing will contain wrong namespace
    """
    def __init__(self, path: str):
        self.path: str = path
        self.temp_dir: str = tempfile.mkdtemp()
        atexit.register(self.__rm_temp_dir)
        self.__extract_files()

    def __rm_temp_dir(self):
        if self.temp_dir:
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def __extract_files(self):
        with zipfile.ZipFile(self.path, 'r') as zr:
            zr.extractall(self.temp_dir)

    def delete_comment_by_id(self, comment_id: Union[int, str]):
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self.temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        for para in root.iter(f'{{{schemas.w}}}p'):
            for element in para:
                if element.tag == f'{{{schemas.w}}}commentRangeStart':
                    if element.attrib[f'{{{schemas.w}}}id'] == str(comment_id):
                        para.remove(element)
                if element.tag == f'{{{schemas.w}}}commentRangeEnd':
                    if element.attrib[f'{{{schemas.w}}}id'] == str(comment_id):
                        para.remove(element)
                if element.tag == f'{{{schemas.w}}}r':
                    for sub_elem in element:
                        if sub_elem.tag == f'{{{schemas.w}}}commentReference':
                            if sub_elem.attrib[f'{{{schemas.w}}}id'] == str(comment_id):
                                para.remove(element)
        tree.write(os.path.join(self.temp_dir, 'word', 'document.xml'), xml_declaration=True,
                   encoding='unicode')

    def edit_comment_by_id(self, comment_id: Union[int, str], comment_text: str, new_author: str = None):
        para_count = 0
        remove_queue = {}
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self.temp_dir, 'word', 'comments.xml'))
        root = tree.getroot()
        for comment in root:
            if comment.attrib[f'{{{schemas.w}}}id'] == str(comment_id):
                if new_author is not None:
                    comment.attrib[f'{{{schemas.w}}}author'] = new_author
                for paragraph in comment:
                    if para_count == 1:
                        remove_queue[paragraph] = comment
                    else:
                        para_count += 1
                        for elem in paragraph:
                            if elem.tag == f'{{{schemas.w}}}r':
                                for sub_elem in elem:
                                    if sub_elem.tag == f'{{{schemas.w}}}t':
                                        sub_elem.text = comment_text

        for para, comm in remove_queue.items():
            comm.remove(para)

        tree.write(os.path.join(self.temp_dir, 'word', 'comments.xml'), xml_declaration=True,
                   encoding='unicode')


    def save(self, path: str = None):
        if path is None:
            path = self.path
        with zipfile.ZipFile(path, 'w') as zw:
            for folder, sub_folder, file_names in os.walk(self.temp_dir):
                for file_name in file_names:
                    file_path = os.path.join(folder, file_name)
                    zw.write(file_path, os.path.relpath(file_path, self.temp_dir))

    def get_document(self):
        xml_path = os.path.join(self.temp_dir, 'word', 'document.xml')
        tree = ET.parse(xml_path)
        root = tree.getroot()
        return ET.tostring(root, encoding='unicode')


class XMLRedactorNumbering:
    """
    At the moment the class works only with word/numbering.xml
    """
    def __init__(self, path: str) -> None:
        if len(path) < 5 or len(path) > 255:
            raise NameError(
                f'\'{path}\' can\'t be open'
            )
        if path[-5:] != '.docx':
            raise NotSupportedFormat(
                path
            )
        self.path: str = path
        self.document: Document = Document(path)
        self.autosave: bool = False
        try:
            self.numbering_xml_file = self.document.part.numbering_part
            self.numbering_xml_tree = self.numbering_xml_file._element
        except KeyError as _ke:
            raise NumberingIsNotExists(
                path
            )

    def save(self, path: str = None):
        if path is None:
            self.document.save(self.path)
        else:
            self.document.save(path)

    @property
    def __get_schema_numId(self) -> str:
        return f'{{{schemas["w"]}}}numId'

    @property
    def __get_schema_abstractNumId(self) -> str:
        return f'{{{schemas["w"]}}}abstractNumId'

    @property
    def __get_new_numId(self) -> int:
        num_elements = self.numbering_xml_tree.findall(
            './/w:num', namespaces=self.numbering_xml_file._element.nsmap
        )
        return 1 + max(int(num.get(self.__get_schema_numId)) for num in num_elements)

    @property
    def __get_new_abstractId(self) -> int:
        anum_elements = self.numbering_xml_tree.findall(
            './/w:abstractNum', namespaces=self.numbering_xml_file._element.nsmap
        )
        return 1 + max(int(anum.get(self.__get_schema_abstractNumId)) for anum in anum_elements)

    def add_new_num(self) -> None:
        """
        !!! Непонятно для чего durableId
        """
        num_element = parse_xml(f'''
            <w:num xmlns:w="{schemas["w"]}" xmlns:w16cid="{schemas["w16cid"]}" w:numId="{self.__get_new_numId}" w16cid:durableId="0">
                <w:abstractNumId w:val="{self.__get_new_abstractId}"/>
            </w:num>
        ''')
        self.numbering_xml_tree.append(num_element)
        if self.autosave:
            self.document.save(self.path)

    def add_new_abstractNum(self) -> None:
        """
        !!! Много аттрибутов, непонятно какие нужны и за что отвечают
        !!! Проблема с bullet list, как различать различные символы, пока только о
        """
        lvl_atr_dec = '''<w:start w:val="1"/>
                        <w:numFmt w:val="decimal"/>
                        <w:lvlText w:val="%1."/>
                        <w:lvlJc w:val="left"/>'''
        lvl_atr_bullet = '''<w:start w:val="1"/>
                        <w:numFmt w:val="bullet"/>
                        <w:lvlText w:val="o"/>
                        <w:lvlJc w:val="left"/>'''
        abstract_num_element = parse_xml(f'''
                <w:abstractNum xmlns:w="{schemas["w"]}" w:abstractNumId="{self.__get_new_abstractId}">
                    <w:nsid w:val="0"/>
                    <w:multiLevelType w:val="hybridMultilevel"/>
                    <w:tmpl w:val="0"/>
                    <w:lvl w:ilvl="0">
                        <w:start w:val="1"/>
                        <w:numFmt w:val="decimal"/>
                        <w:lvlText w:val="%1."/>
                        <w:lvlJc w:val="left"/>
                        <w:pPr>
                            <w:ind w:left="720" w:hanging="360"/>
                        </w:pPr>
                        <w:rPr>
                            <w:rFonts w:ascii="default"/>
                        </w:rPr>
                    </w:lvl>
                </w:abstractNum>
            ''')
        self.numbering_xml_tree.append(abstract_num_element)
        if self.autosave:
            self.document.save(self.path)


example_abstract_decimal = '''
    <w:abstractNum w:abstractNumId="2">
        <w:nsid w:val="0"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="0"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="default"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
'''

example_abstract_bullet = '''
    <w:abstractNum w:abstractNumId="3">
        <w:nsid w:val="0"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="0"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val=""/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Symbol" w:eastAsiaTheme="minorHAnsi" w:hAnsi="Symbol" w:cstheme="minorBidi"
                          w:hint="default"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
'''

example_num = '''
    <w:num w:numId="3" w16cid:durableId="0">
        <w:abstractNumId w:val="2"/>
    </w:num>
'''
