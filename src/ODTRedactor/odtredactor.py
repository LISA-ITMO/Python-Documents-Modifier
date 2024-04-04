import zipfile
import xml.etree.ElementTree as ET
from odf.namespaces import nsdict

import src.ODTRedactor.models as models 

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
        """Saving content.xml to an xml file"""

        with open('xml.xml', 'wb') as file:
            file.write(self.data['content.xml'])

    def add_comment_by_text(self, text: str, text_annotation: str, author: str = "Lisa") -> None:
        """Add a comment to the text using text search"""

        for paragraph in self.stringroot.iter(models.Styles.PARAGRAPH):
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

    def edit_style_by_text(self, text: str, font_name: str, font_size: int, color, name=None, family: str = "text") -> None:   
        """Edit a paragraph of text using text search"""     
        automatic_styles = self.stringroot.find(models.Styles.AUTOMATIC_STYLE)

        for paragraph in self.stringroot.iter(models.Styles.PARAGRAPH):
            if text in self.__get_text_from_children(paragraph):
                for span in paragraph.iter(tag=models.Styles.SPAN):
                    style = models.Style(
                        data=self.stringroot,
                        font_name=font_name,
                        font_size=font_size,
                        color=color,
                        parent_style_name=span.attrib[models.NameSpaces.TEXT+'style-name'],
                        name=name,
                        family=family,
                        text_position=None
                    )

                    parent_style = self.__get_style(span.attrib[models.NameSpaces.TEXT+'style-name'])
                    properties = next(parent_style.iter(models.NameSpaces.STYLE+'text-properties'))
                    if models.NameSpaces.STYLE+'text-position' in properties.keys():
                        style.text_position = properties.attrib[models.Styles.TEXT_POSITION]

                    automatic_styles.append(
                        style.create_element_style()
                    )

                    span.attrib[models.NameSpaces.TEXT+'style-name'] = style.name

                break


        self.save_file()

    def to_text(self) -> None:
        """Return the text of the entire document"""
        text = []

        for i in self.stringroot.iter(models.Styles.PARAGRAPH):
            text.append(self.__get_text_from_children(i))

        return '\n'.join(text)
    
    def __get_style(self, name: str) -> ET.Element:
        automatic_styles = self.stringroot.find(models.Styles.AUTOMATIC_STYLE)
        for i in automatic_styles.iter(models.Styles.STYLE):
            if i.attrib[models.NameSpaces.STYLE+'name'] == name:
                return i
    
    @staticmethod
    def __clear_paragraph(paragraph) -> None:
        paragraph.text = None
        for i in paragraph.iter(models.Styles.SPAN):
            i.text = None
            
    @staticmethod
    def __get_text_from_children(paragraph) -> str:
        return ' '.join(filter(None, [paragraph.text] + [i.text for i in paragraph.iter(models.Styles.SPAN)]))

