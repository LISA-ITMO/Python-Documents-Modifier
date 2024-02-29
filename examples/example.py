import sys, os

parent_dir = os.path.abspath(os.path.join(os.path.dirname('.'), "."))
sys.path.append(parent_dir)

from src.odtredactor import ODTRedactor

if __name__ == '__main__':
    file = ODTRedactor('examples/example.odt')

    file.add_annotation(
        text="Пример некоторого текста",
        ann_text="Некоторая аннотация",
    )