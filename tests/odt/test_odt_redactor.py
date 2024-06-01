import unittest
from os.path import join, dirname, abspath
from src.odt.odtredactor import ODTRedactor

class TestODTRedactor(unittest.TestCase):
    root = join(dirname(dirname(abspath(__file__))), 'odt')

    def test_create_object(self):
        file = ODTRedactor(
            path_input=join(self.root, 'test_file.odt'),
            path_output=join(self.root, 'test_file_edited.odt')
        )

    def test_add_comment_by_text(self):
        file = ODTRedactor(
            path_input=join(self.root, 'test_file.odt'),
            path_output=join(self.root, 'test_file_edited.odt')
        )

        file.add_comment_by_text(
            text="Ночь на перекрёстке",
            text_annotation="test annotation",
            author="Lisa"
        )

        file.save_file()

    def test_correct_save(self):
        file = ODTRedactor(
            path_input=join(self.root, 'test_file.odt'),
            path_output=join(self.root, 'test_file_edited.odt')
        )
        file.save_file()

    def test_edit_style_by_text(self):
        file = ODTRedactor(
            path_input=join(self.root, 'test_file.odt'),
            path_output=join(self.root, 'test_file_edited.odt')
        )
        file.edit_style_by_text(
            text='режиссёра',
            font_name="Arial",
            font_size=10,
            color="#ff0000",
        )
        file.save_file()



if __name__ == '__main__':
    unittest.main()