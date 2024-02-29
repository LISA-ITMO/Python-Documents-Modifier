from odf import text as t, office
from odf.element import Element
from odf.opendocument import load
from odf.namespaces import DCNS
from odf.style import TextProperties


from typing import Union

class ODTRedactor:
    """Class for working with ODT documents"""
    def __init__(self, path) -> None:
        self.path = path
        self.file = None
        self.load_file()

    def load_file(self) -> None:
        """Function for loading a file by path"""
        self.file = load(self.path)

    def save_file(self) -> None:
        """Function for saving a file"""
        self.file.save(self.path)

    def _add_author(self, name: str, **args) -> Element:
        """Function to add an author to a text element

        Attributes
        ----------
        name: :class:`str`
            Author's name
        """
        return Element(qname = (DCNS, 'creator'), text=name, **args)

    def add_annotation(self, text: str, ann_text: str, author: str = None) -> None:
        """Function for adding annotations to text based on a passage

        Attributes
        ----------
        text: :class:`str`
            A piece of text to which an annotation will be added
        ann_text: :class:`str`
            Annotation text
        author: :class:`None` | :class:`str`
            Author's name. By default the author is not added
        """

        assert all(isinstance(i, str) for i in [text, ann_text]) and isinstance(author, Union[str, None]), "Arguments must be strings."

        for para in self.file.getElementsByType(t.P):
            if text in str(para):
                annot = office.Annotation()
                annot.addElement(t.P(text=ann_text))
                if author:
                    annot.addElement(self._add_author(name=author))
                para.addElement(annot)

        self.save_file()
        
    def show_xml(self):
        print(self.file.xml())



                        
redactor = ODTRedactor("example.odt")
redactor.add_annotation('asadtge', 'some annot')
redactor.show_xml()