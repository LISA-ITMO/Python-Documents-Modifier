import unittest
from src.docx.exceptionsdocx import *
from src.docx.docxredactor import DOCXRedactor
from src.docx.enum.color import Color
from docx.opc.exceptions import PackageNotFoundError


class TestExceptions(unittest.TestCase):
    def test_not_supported_format(self):
        with self.assertRaises(NotSupportedFormat):
            DOCXRedactor('../docx/file.txt')

    def test_supported_format(self):
        with self.assertRaises(PackageNotFoundError):
            DOCXRedactor('../docx/not_exist_file.docx')

    def test_correct_file(self):
        self.assertIsInstance(DOCXRedactor('../docx/file.docx'), DOCXRedactor)

    def test_impossible_color(self):
        with self.assertRaises(ImpossibleColor):
            Color(260, 5, 5)

    def test_possible_color(self):
        Color(128, 0, 255)

    def test_paragraph_not_found(self):
        doc = DOCXRedactor('../docx/file.docx')
        with self.assertRaises(ParagraphNotFound):
            doc.edit_style_by_id('not_exist_paraId', size=14)

    def test_exist_paragraph(self):
        doc = DOCXRedactor('../docx/file.docx')
        para_id = doc.all_para_attributes()[0].values()[0]
        doc.edit_style_by_id(para_id, size=15)
        doc.save()


if __name__ == '__main__':
    unittest.main()
