import json
import os
import tempfile
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from src.docx.enum.schemas import schemas
from typing import List, Optional, Dict
from os.path import join, dirname, abspath
import logging


class Elem:
    def __init__(self, num: str, text: str, anchor_id: str = None):
        self.num: str = num
        self.text: str = text
        self.sub_elements: Optional[List[Elem]] = None
        self.anchor_id: str = anchor_id

    def append(self, elem):
        if self.sub_elements is None:
            self.sub_elements = []
        self.sub_elements.append(elem)

    def __len__(self):
        return 0 if self.sub_elements is None else len(self.sub_elements)


class Parser:
    def __init__(self, path: str):
        self.path = path
        self._temp_dir = tempfile.mkdtemp()
        self._extract_files()

    def _extract_files(self):
        with zipfile.ZipFile(self.path, 'r') as zr:
            zr.extractall(self._temp_dir)

    def parse(self):
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
        if path is None:
            path = self.path
        with open(f'{path[:-5]}.json', 'w', encoding='utf-8') as json_file:
            json.dump(struct, json_file, indent=4, default=handle_none,
                      ensure_ascii=False)

    def get_other_text(self):
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        all_text = ''
        for para in root.findall(f'.//{{{schemas.w}}}body//{{{schemas.w}}}p'):
            for t in para.iter(f'{{{schemas.w}}}t'):
                all_text += f'{t.text} '
        return all_text


def struct_to_dict(elements: List[Elem], p: Parser):
    if elements is None:
        return

    def convert_element_to_dict(elem: Elem, p: Parser):
        elem_dict = {
            "num": elem.num,
            "title": elem.text,
            'text': p.parse_paragraphs_from_anchor(elem.anchor_id)
        }
        if elem.sub_elements:
            elem_dict["sub_elements"] = [convert_element_to_dict(sub_elem, p) for sub_elem in elem.sub_elements]
        return elem_dict
    return [convert_element_to_dict(elem, p) for elem in elements]


def handle_none(obj):
    if obj is None:
        return "null"
    return obj


def iter_docs():
    root = str(join(dirname(abspath(__file__))))
    for filename in os.listdir(join(root, 'documents_for_extract')):
        print(filename)


def parsing_documents():
    root_in = join(str(join(dirname(abspath(__file__)))), 'documents_for_extract')
    root_out = join(str(join(dirname(abspath(__file__)))), 'test_out')
    i = 1
    c = len(os.listdir(root_in))
    for filename in os.listdir(root_in):
        try:
            p = Parser(join(root_in, filename))
            s, pot = p.parse()
            n = {'potentially_damage': pot, 'table_of_content': struct_to_dict(s, p),
                    'other_text': p.get_other_text()}
            p.save(n, join(root_out, filename))
        except Exception as _e:
            logging.warn(_e)
        finally:
            sys.stdout.write(f"\rProgress: [{i} // {c}]")
            sys.stdout.flush()
            i += 1


if __name__ == '__main__':
    parsing_documents()
