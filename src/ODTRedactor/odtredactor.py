import zipfile
import xml.etree.ElementTree as ET
from odf.namespaces import nsdict

import src.ODTRedactor.models as models
from src.ODTRedactor.styles import Styles


class ODTRedactor:
    """Class for working with ODT documents"""
    def __init__(self, path_input: str, path_output: str) -> None:
        self.path_input = path_input
        self.path_output = path_output
        self.data = None
        self.stringroot = None
        self.load_file()

    def load_file(self) -> None:
        """Function for loading a file by path"""
        with zipfile.ZipFile(self.path_input, 'r') as zip_ref:
            file_data = {}
            for file in zip_ref.namelist():
                with zip_ref.open(file) as f:
                    file_data[file] = f.read()
        self.data = file_data
        self.stringroot = ET.fromstring(self.data['content.xml'].decode('UTF-8'))

        for i in nsdict:
            if f"xmlns:{nsdict[i]}" in self.stringroot.attrib or i in self.stringroot.attrib.values(): continue
            self.stringroot.attrib[f"xmlns:{nsdict[i]}"] = i

    def save_file(self) -> None:
        """Function for saving a file"""

        self.data['content.xml'] = bytes(ET.tostring(self.stringroot, encoding="UTF-8", xml_declaration=True, default_namespace=None))

        with zipfile.ZipFile(self.path_output, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file, data in self.data.items():
                zip_file.writestr(file, data)
    
    def save_xml(self) -> None:
        with open('xml.xml', 'wb') as file:
            file.write(self.data['content.xml'])

    def add_comment_by_text(self, text: str, text_annotation: str, author: str = "Lisa") -> None:
        for paragraph in self.stringroot.iter(Styles.PARAGRAPH):
            if text in self.__get_text_from_children(paragraph):
                annotation = models.Annotation(
                    text_annotation=text_annotation,
                    author=author 
                )

                paragraph.insert(0, annotation.get_start())

                paragraph.append(
                    annotation.get_end()
                )

                break

        self.save_file()
        # self.save_xml()

    def to_text(self) -> None:
        text = []

        for i in self.stringroot.iter(Styles.PARAGRAPH):
            text.append(self._get_text_from_children(i))

        return '\n'.join(text)
    
    @staticmethod
    def __clear_paragraph(paragraph) -> None:
        paragraph.text = None
        for i in paragraph.iter(Styles.SPAN):
            i.text = None
            
    @staticmethod
    def __get_text_from_children(paragraph) -> str:
        return ' '.join(filter(None, [paragraph.text] + [i.text for i in paragraph.iter(Styles.SPAN)]))

