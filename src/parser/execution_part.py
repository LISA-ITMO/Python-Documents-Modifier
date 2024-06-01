import os
import sys
from src.parser.Elem import Elem
from src.parser.Parser import Parser
from typing import List
from os.path import join, dirname, abspath
import logging


def struct_to_dict(elements: List[Elem], p: Parser):
    """
    Converts a list of chapter-elements to dict
    :param elements: list of elements
    :param p: parser
    """
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


def parsing_documents():
    """
    Iterations all DOCX-documents in root_in, converts everything to JSON-format and saves it in root_out
    """
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
