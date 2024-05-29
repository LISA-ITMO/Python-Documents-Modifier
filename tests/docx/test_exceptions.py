import unittest
from src.exceptions.docx_exceptions import *
from src.docx.docx_redactor import DOCXRedactor
from src.docx.xml_redactor import XMLRedactor
from src.docx.enum.color import Color
from src.docx.enum.schemas import schemas
from src.docx.enum.ListStyle import ListStyle
from os.path import join, dirname, abspath


class TestExceptions(unittest.TestCase):
    root = join(dirname(dirname(abspath(__file__))), 'docx')

    def test_not_supported_format(self):
        with self.assertRaises(NotSupportedFormat):
            DOCXRedactor(join(self.root, 'file.txt'))

    def test_supported_format(self):
        with self.assertRaises(FileNotFoundError):
            DOCXRedactor(join(self.root, 'not_exists_file.docx'))

    def test_correct_file(self):
        self.assertIsInstance(DOCXRedactor(join(self.root, 'file.docx')), DOCXRedactor)

    def test_impossible_color(self):
        with self.assertRaises(ImpossibleColor):
            Color(260, 5, 5)

    def test_possible_color(self):
        Color(128, 0, 255)

    def test_paragraph_not_found(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        with self.assertRaises(ParagraphNotFound):
            doc.edit_style_by_id('not_exist_paraId', size=14)

    def test_exist_paragraph(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        para_id = doc.all_para_attributes()[0].values()[0]
        doc.edit_style_by_id(para_id, size=15)
        doc.save('output_file.docx')

    def test_paragraph_does_not_contain_numPr(self):
        xr = XMLRedactor(join(self.root, 'file.docx'))
        with self.assertRaises(ParagraphDoesNotContainNumPr):
            first_paraId = xr.get_all_para_attributes()[0][f'{{{schemas.w14}}}paraId']
            xr.edit_list_style_by_paraIds(first_paraId, 0)

    def test_paragraph_contain_numPr(self):
        xr = XMLRedactor(join(self.root, 'file.docx'))
        second_paraId = xr.get_all_para_attributes()[1][f'{{{schemas.w14}}}paraId']
        xr.edit_list_style_by_paraIds(second_paraId, 0)
        xr.save('output_file.docx')

    def test_ilvl_does_not_exist(self):
        xr = XMLRedactor(join(self.root, 'file.docx'))
        with self.assertRaises(ILvlDoesNotExist):
            xr.add_new_abstract_and_num(ListStyle.bullet, 100, '*')

    def test_ilvl_exist(self):
        xr = XMLRedactor(join(self.root, 'file.docx'))
        xr.add_new_abstract_and_num(ListStyle.bullet, 8, '*')
        xr.save('output_file.docx')

    def test_docx_does_not_contain_xml_file(self):
        xr = XMLRedactor(join(self.root, 'file.docx'))
        with self.assertRaises(FileDoesNotContainXMLFile):
            xr.edit_comment_by_id('0', 'text', 'author')


if __name__ == '__main__':
    unittest.main()
