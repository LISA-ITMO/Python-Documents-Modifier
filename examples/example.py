import sys, os

parent_dir = os.path.abspath(os.path.join(os.path.dirname('.'), "."))
sys.path.append(parent_dir)

from src.ODTRedactor.odtredactor import ODTRedactor

if __name__ == '__main__':
    file = ODTRedactor('examples/example.odt', 'examples/editedExample.odt')

    file.add_comment_by_text(
        text="Пример некоторого текста",
        text_annotation="Некоторая аннотация",
    )