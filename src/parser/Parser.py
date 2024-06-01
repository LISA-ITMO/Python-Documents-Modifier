import json
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from src.docx.enum.schemas import schemas
from src.parser.Elem import Elem
from typing import List, Dict


class Parser:
    """
    Class, that parses DOCX-document and saves it in json-format
    :param self.path: path to file
    :param self._temp_dir: path to the temp directory, where the unpacked DOCX-document is stored
    """
    def __init__(self, path: str):
        self.path = path
        self._temp_dir = tempfile.mkdtemp()
        self._extract_files()

    def _extract_files(self):
        """
        Unpacks DOCX-document in temp directory
        """
        with zipfile.ZipFile(self.path, 'r') as zr:
            zr.extractall(self._temp_dir)

    def parse(self):
        """
        Parses document.xml part and returns content and flag (potentially damaged)
        """
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        paragraphs = []
        for element in root.iter(f'{{{schemas.w}}}sdt'):
            for content in element.iter(f'{{{schemas.w}}}sdtContent'):
                for para in content.iter(f'{{{schemas.w}}}p'):
                    paragraphs.append(para)
        return self._parse_paragraphs(paragraphs)

    def parse_paragraphs_from_anchor(self, anchor_id: str,
                                     list_view: bool = True):
        """
        Parses text, that attached to a chapter by anchor_id
        """
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()

        check = False
        text = []
        for p in root.iter(f"{{{schemas.w}}}p"):
            if p.findall(f'{{{schemas.w}}}bookmarkStart') and check:
                return text if list_view else '\n'.join(text)
            if check:
                text.append(self._parse_text_from_anchor(p))
            if p.findall(f'{{{schemas.w}}}bookmarkStart[@{{{schemas.w}}}name="{anchor_id}"]'):
                check = True

    def _parse_text_from_anchor(self, paragraph: Element):
        return ''.join([i for i in paragraph.itertext()])

    def _parse_paragraphs(self, paragraphs: List[Element]):
        struct = []
        potentially_damage = False
        for para in paragraphs:
            for pPr in para.iter(f'{{{schemas.w}}}pPr'):
                for style in pPr.iter(f'{{{schemas.w}}}pStyle'):
                    val = style.attrib[f'{{{schemas.w}}}val']
                    hyperlink = para.iter(f'{{{schemas.w}}}hyperlink')
                    hyperlink_id = None
                    for i in hyperlink:
                        hyperlink_id = i.attrib[f"{{{schemas.w}}}anchor"]
                    if not val.isdigit():
                        continue
                    elif int(val) // 10 == 1:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            struct.append(Elem(str(len(struct) + 1), text.text, hyperlink_id))
                            break
                    elif int(val) // 10 == 2:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            if len(struct) == 0:
                                struct.append(Elem(str(len(struct) + 1), text.text, hyperlink_id))
                                potentially_damage = True
                                break
                            else:
                                num = f'{str(len(struct))}.{str(len(struct[-1]) + 1)}'
                                struct[-1].append(Elem(num, text.text, hyperlink_id))
                                break
                    elif int(val) // 10 == 3:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            if len(struct) == 0:
                                potentially_damage = True
                                struct.append(Elem(str(len(struct) + 1), text.text, hyperlink_id))
                                break
                            elif len(struct[-1]) == 0:
                                potentially_damage = True
                                num = f'{str(len(struct))}.{str(len(struct[-1]) + 1)}'
                                struct[-1].append(Elem(num, text.text, hyperlink_id))
                                break
                            else:
                                num = (f'{str(len(struct))}.{str(len(struct[-1]))}.'
                                       f'{str(len(struct[-1].sub_elements) + 1)}')
                                struct[-1].sub_elements[-1].append(Elem(num, text.text, hyperlink_id))
                                break
        return struct, potentially_damage

    def save(self, struct: Dict, path: str = None):
        """
        Save DOCX-document as JSON
        """
        if path is None:
            path = self.path
        with open(f'{path[:-5]}.json', 'w', encoding='utf-8') as json_file:
            json.dump(struct, json_file, indent=4, default=_handle_none,
                      ensure_ascii=False)

    def get_other_text(self):
        """
        Gets other text from a document
        """
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        all_text = ''
        for para in root.findall(f'.//{{{schemas.w}}}body//{{{schemas.w}}}p'):
            for t in para.iter(f'{{{schemas.w}}}t'):
                all_text += f'{t.text} '
        return all_text


def _handle_none(obj):
    if obj is None:
        return "null"
    return obj
