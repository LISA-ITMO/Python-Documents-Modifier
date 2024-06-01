import unittest
import json
import os
from os.path import join, dirname, abspath
from src.parser.parser import *


class TestDocxParser(unittest.TestCase):
    def test_out_json(self):
        path_in = join(dirname(dirname(abspath(__file__))), 'documents_for_extract')
        path_result = join(dirname(dirname(abspath(__file__))), 'test_out')
        for filename in os.listdir(path_in):
            p = Parser(join(path_in, filename))
            struct, pot = p.parse()
            json_out = {'potentially_damage': pot, 'table_of_content': struct_to_dict(struct, p),
                        'other_text': p.get_other_text()}
            with open(join(path_result, filename.replace('.docx', '.json')),
                      'r', encoding='utf-8') as json_reader:
                data_dict = json.load(json_reader)
                if data_dict != json_out:
                    raise Exception()


if __name__ == '__main__':
    unittest.main()
