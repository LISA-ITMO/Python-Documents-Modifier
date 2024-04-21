import unittest
from src.docx.docx_redactor import DOCXRedactor
from src.docx.enum.ListStyle import ListStyle
from os.path import join, dirname, abspath


class TestDocxRedactor(unittest.TestCase):
    root = join(dirname(dirname(abspath(__file__))), 'docx')

    def test_create_object(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))

    def test_correct_save(self):
        DOCXRedactor(join(self.root, 'file.docx')).save('output_file.docx')

    def test_add_comment_by_id(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        paraId = doc.all_para_attributes()[0].values()[0]
        doc.add_comment_by_id(paraId, 'comment', 'author')
        doc.save('output_file.docx')

    def test_edit_comment_by_id(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        paraId = doc.all_para_attributes()[0].values()[0]
        doc.add_comment_by_id(paraId, 'comment', 'author')
        doc.save('output_file.docx')
        doc = DOCXRedactor(join(self.root, 'output_file.docx'))
        doc.edit_comment_by_id('0', 'new_comment', 'new_author')
        doc.save()

    def test_delete_comment_by_id(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        paraId = doc.all_para_attributes()[0].values()[0]
        doc.add_comment_by_id(paraId, 'comment', 'author')
        doc.save('output_file.docx')
        doc = DOCXRedactor(join(self.root, 'output_file.docx'))
        doc.delete_comment_by_id('0')
        doc.save()

    def test_edit_list_style(self):
        doc = DOCXRedactor(join(self.root, 'file.docx'))
        paraId = doc.all_para_attributes()[1].values()[0]
        doc.edit_list_style_by_paraIds(paraId, ListStyle.bullet, '*')
        doc.save('output_file.docx')


if __name__ == '__main__':
    unittest.main()
