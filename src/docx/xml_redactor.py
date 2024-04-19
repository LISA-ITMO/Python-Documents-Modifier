import zipfile
import os
import shutil
import tempfile
import atexit
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from typing import Union, Tuple
from src.exceptions.docx_exceptions import ParagraphDoesNotContainNumPr, ILvlDoesNotExist
from src.docx.enum.schemas import schemas


class XMLRedactor:
    """
    For correct work this class you should edit xml.etree.ElementTree
    replace 862 line on : "return qnames, _namespace_map"
    Else DOCX-file after editing will contain wrong namespace
    """
    def __init__(self, path: str):
        self.path: str = path
        self._temp_dir: str = tempfile.mkdtemp()
        self.encoding: str = 'utf-8'
        atexit.register(self.__rm_temp_dir)
        self.__extract_files()

    def __rm_temp_dir(self):
        if self._temp_dir:
            shutil.rmtree(self._temp_dir, ignore_errors=True)

    def __extract_files(self):
        with zipfile.ZipFile(self.path, 'r') as zr:
            zr.extractall(self._temp_dir)

    def __get_ilvl_marking(self) -> dict[int, Tuple[int, int]]:
        return {
            0: (720, 360),
            1: (1440, 360),
            2: (2160, 360),
            3: (2880, 360),
            4: (3600, 360),
            5: (4320, 360),
            6: (5040, 360),
            7: (5760, 360),
            8: (6480, 360)
        }

    def delete_comment_by_id(self, comment_id: Union[int, str]):
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
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
        tree.write(os.path.join(self._temp_dir, 'word', 'document.xml'), xml_declaration=True,
                   encoding=self.encoding)

    def edit_comment_by_id(self, comment_id: Union[int, str], comment_text: str, new_author: str = None):
        para_count = 0
        remove_queue = {}
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'comments.xml'))
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

        tree.write(os.path.join(self._temp_dir, 'word', 'comments.xml'), xml_declaration=True,
                   encoding=self.encoding)


    def save(self, path: str = None):
        if path is None:
            path = self.path
        with zipfile.ZipFile(path, 'w') as zw:
            for folder, sub_folder, file_names in os.walk(self._temp_dir):
                for file_name in file_names:
                    file_path = os.path.join(folder, file_name)
                    zw.write(file_path, os.path.relpath(file_path, self._temp_dir))

    def add_new_abstract_and_num(self, list_style: bool = True, levels_count: int = 1,
                                 bullet_symbol: str = '•') -> int:
        """
        This function added new abstractNum and num
        list_style -- True: decimal, False: bullet
        Return id of new Num-object
        """
        for k, v in schemas.to_namespace.items():
            ET.register_namespace(k, v)
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'numbering.xml'))
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
        for i in range(levels_count):
            insert_abstract.append(self.__create_ilvl(list_style, i, bullet_symbol))
        insert_num = ET.Element(f'{{{schemas.w}}}num', {
            f'{{{schemas.w}}}numId': str(max_num + 1),
            f'{{{schemas.w16cid}}}durableId': '0'
        })
        insert_num.append(ET.Element(f'{{{schemas.w}}}abstractNumId', {
            f'{{{schemas.w}}}val': str(max_abstract + 1)
        }))
        root.insert(len(root.findall(f'.//{{{schemas.w}}}abstractNum')), insert_abstract)
        root.append(insert_num)

        tree.write(os.path.join(self._temp_dir, 'word', 'numbering.xml'), xml_declaration=True,
                   encoding=self.encoding)
        return max_num + 1

    def __create_ilvl(self, is_decimal: bool, i_lvl: int = 0, bullet_symbol: str = '•') -> Element:
        if not (0 <= i_lvl <= 8):
            raise ILvlDoesNotExist(i_lvl)
        params = self.__get_ilvl_marking()[i_lvl]
        i_lvl = ET.Element(f'{{{schemas.w}}}lvl', {f'{{{schemas.w}}}ilvl': str(i_lvl)})
        i_lvl.append(ET.Element(f'{{{schemas.w}}}start', {f'{{{schemas.w}}}val': '1'}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}numFmt',
                                {f'{{{schemas.w}}}val': 'decimal' if is_decimal else 'bullet'}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}lvlText',
                                {f'{{{schemas.w}}}val': '%1.' if is_decimal else bullet_symbol}))
        i_lvl.append(ET.Element(f'{{{schemas.w}}}lvlJc', {f'{{{schemas.w}}}val': 'left'}))
        temp_lem = ET.Element(f'{{{schemas.w}}}pPr')
        temp_lem.append(ET.Element(
            f'{{{schemas.w}}}ind',
            {f'{{{schemas.w}}}left': str(params[0]), f'{{{schemas.w}}}hanging': str(params[1])}
        ))
        i_lvl.append(temp_lem)
        temp_lem = ET.Element(f'{{{schemas.w}}}rPr')
        temp_lem.append(ET.Element(f'{{{schemas.w}}}rFonts',
                                   {f'{{{schemas.w}}}hint': 'default'}))
        i_lvl.append(temp_lem)
        return i_lvl

    def edit_list_style_by_paraIds(self, paraIds: Union[list[str], str],
                                   new_num: Union[str, int]):
        """
        Replace the numId to newNum of all specified paragraphs
        :param paraIds: One or many paraIds, which we replace numId
        :param new_num: New numId
        :return: None
        """
        if isinstance(paraIds, str):
            paraIds = [paraIds]
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        for paragraph in root.iter(f'{{{schemas.w}}}p'):
            if paragraph.attrib[f'{{{schemas.w14}}}paraId'] in paraIds:
                try:
                    cnt = 0
                    for num in paragraph.iter(f'{{{schemas.w}}}numId'):
                        num.attrib[f'{{{schemas.w}}}val'] = str(new_num)
                        cnt += 1
                    if cnt == 0:
                        raise KeyError()
                except KeyError as _ke:
                    raise ParagraphDoesNotContainNumPr(paragraph.attrib[f'{{{schemas.w14}}}paraId'])
        tree.write(os.path.join(self._temp_dir, 'word', 'document.xml'), xml_declaration=True,
                   encoding='utf-8')
