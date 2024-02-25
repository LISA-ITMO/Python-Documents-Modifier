from odf import text as t, office
from odf.element import Element
from odf.opendocument import load
from odf.namespaces import DCNS

from typing import Union

class ODTRedactor:
    def __init__(self, path) -> None:
        self.path = path
        self.file = None
        self.load_file()

    def load_file(self) -> None:
        self.file = load(self.path)

    def save_file(self) -> None:
        self.file.save(self.path)

    def add_author(self, name, **args) -> Element:
        return Element(qname = (DCNS, 'creator'), text=name, **args)

    def annotation(self, text, ann_text, author=None) -> None:
        assert all(isinstance(i, Union[str, None]) for i in [text, ann_text, author]), "Аргументы должны быть строковыми."
        for para in self.file.getElementsByType(t.P):
            if text in str(para):
                annot = office.Annotation()
                annot.addElement(t.P(text=ann_text))
                if author:
                    annot.addElement(self.add_author(name=author))
                para.addElement(annot)

        self.save_file()