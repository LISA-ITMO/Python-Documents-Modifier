import sys, os

parent_dir = os.path.abspath(os.path.join(os.path.dirname('.'), "."))
sys.path.append(parent_dir)

from src.ODTRedactor.odtredactor import ODTRedactor

if __name__ == '__main__':
    file = ODTRedactor(
        path_input=r'examples/example.odt',
        path_output=r'examples/example_edited.odt'
    )
    file.edit_style_by_text("Фильм стал", font_name="Arial", font_size=12, color="#ff0000")
    file.add_comment_by_text("Фильм", text_annotation="example annotation")