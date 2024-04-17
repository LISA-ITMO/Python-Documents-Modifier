from docx import Document
from docx.oxml import parse_xml
from src.exceptions.docx_exceptions import NotSupportedFormat, NumberingIsNotExists
import zipfile
import os
import shutil
import tempfile
import atexit
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
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

    def add_new_abstract_and_num(self, is_decimal: bool = True) -> int:
        """
        Return id of new Num-object
        """
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self.temp_dir, 'word', 'numbering.xml'))
        root = tree.getroot()
        max_abstract = -1
        max_num = 0
        for abstract_num in root.iter(f'{{{schemas.w}}}abstractNum'):
            max_abstract = max(max_abstract, int(abstract_num.attrib[f'{{{schemas.w}}}abstractNumId']))
        for num in root.iter(f'{{{schemas.w}}}num'):
            max_num = max(max_num, int(num.attrib[f'{{{schemas.w}}}numId']))
        insert_abstract = ET.Element(f'{{{schemas.w}}}abstractNum', {
            f'{{{schemas.w}}}abstractNumId': str(max_abstract + 1),
            f'{{{schemas.w15}}}restartNumberingAfterBreak': '0'
        })
        insert_abstract.append(ET.Element(f'{{{schemas.w}}}multiLevelType', {
            f'{{{schemas.w}}}val': 'hybridMultilevel'
        }))
        insert_abstract.append(self.__create_ilvl(is_decimal))
        insert_num = ET.Element(f'{{{schemas.w}}}num', {
            f'{{{schemas.w}}}numId': str(max_num + 1),
            f'{{{schemas.w16cid}}}durableId': '0'
        })
        insert_num.append(ET.Element(f'{{{schemas.w}}}abstractNumId', {
            f'{{{schemas.w}}}val': str(max_abstract + 1)
        }))
        root.insert(len(root.findall(f'.//{{{schemas.w}}}abstractNum')), insert_abstract)
        root.append(insert_num)

        tree.write(os.path.join(self.temp_dir, 'word', 'numbering.xml'), xml_declaration=True,
                   encoding='unicode')
        return max_num + 1

    def __create_ilvl(self, is_decimal: bool) -> Element:
        i_lvl = ET.Element(f'{{{schemas.w}}}lvl', {f'{{{schemas.w}}}ilvl': '0'})
        i_lvl.append(ET.Element(f'{{{schemas.w}}}start', {f'{{{schemas.w}}}val': '1'}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}numFmt',
                                {f'{{{schemas.w}}}val': 'decimal' if is_decimal else 'bullet'}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}lvlText',
                                {f'{{{schemas.w}}}val': '%1.' if is_decimal else ''}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}lvlJc', {f'{{{schemas.w}}}val': 'left'}))
        temp_lem = ET.Element(f'{{{schemas.w}}}pPr')
        temp_lem.append(ET.Element(
            f'{{{schemas.w}}}ind',
            {f'{{{schemas.w}}}left': '720', f'{{{schemas.w}}}hanging': '360'}
        ))
        i_lvl.append(temp_lem)
        temp_lem = ET.Element(f'{{{schemas.w}}}rPr')
        temp_lem.append(ET.Element(f'{{{schemas.w}}}rFonts',
                                   {f'{{{schemas.w}}}hint': 'default'}))
        i_lvl.append(temp_lem)
        return i_lvl
