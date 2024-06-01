from src.odt.enum.styles import Styles
from src.odt.enum.style import Style
from src.odt.enum.annotation import Annotation
from src.odt.enum.namespaces import NameSpaces

import zipfile
import xml.etree.ElementTree as ET

from typing import Union

class ODTRedactor:
    """Class for working with ODT documents"""
    def __init__(self, path_input: str, path_output: str, encoding: str = "UTF-8", datafile: str = 'content.xml') -> None:
        self.path_input = path_input
        self.path_output = path_output
        self.data = self.__load_file(path_input)
        self.stringroot = self.__get_root(self.data, filename=datafile, encoding=encoding)

    def save_file(self) -> None:
        """Function for saving a file"""

        self.data['content.xml'] = bytes(ET.tostring(self.stringroot, encoding="UTF-8", xml_declaration=True, default_namespace=None))

        with zipfile.ZipFile(self.path_output, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file, data in self.data.items():
                zip_file.writestr(file, data)

    def add_comment_by_text(self, text: str, text_annotation: str, author: str = "Lisa") -> None:
        """
        :param text: The text by which the paragraph to be commented on will be found
        :param text_annotation: Annotation contents
        :param author: Name of the author of the annotation
        """

        for paragraph in self.stringroot.iter(Styles.PARAGRAPH):
            if text in self.__get_text_from_children(paragraph):
                annotation = Annotation(
                    text_annotation=text_annotation,
                    author=author 
                )

                paragraph.insert(0, annotation.get_start())

                paragraph.append(
                    annotation.get_end()
                )

                break

    def delete_comment_by_id(self, name_id: Union[int, str]) -> None:
        """
        :param name_id: name_id of the annotation to be deleted
        """
        name_id = str(name_id)

        for p in self.stringroot.iter(Styles.PARAGRAPH):
            for i in p.iter(Styles.ANNOTATION):
                if name_id == i.attrib[NameSpaces.XML+'id']:
                    id = i.attrib[NameSpaces.OFFICE+'name']
                    p.remove(i)
                    break

        for p in self.stringroot.iter(Styles.PARAGRAPH):
            for i in p.iter(Styles.SPAN):
                ann =  i.find(Styles.ANNOTATION_END)
                if isinstance(ann, ET.Element) and ann.attrib[NameSpaces.OFFICE+'name'] == id:
                    p.remove(i)
                    break

    def edit_style_by_text(self, text: str, font_name: str, font_size: int, color, style_name=None, family: str = "text") -> None:   
        """
        :param text: the text for which the paragraph will be searched
        :param font_name: new name of the font
        :param font_size: new font size
        :param color: new paragraph color
        :param name: The name of the new style. 
        In case the argument is not passed, it is generated automatically
        :param family: The family of the style from which the main features of the style are inherited
        """
        automatic_styles = self.stringroot.find(Styles.AUTOMATIC_STYLE)

        for paragraph in self.stringroot.iter(Styles.PARAGRAPH):
            if text in self.__get_text_from_children(paragraph):
                for span in paragraph.iter(tag=Styles.SPAN):
                    style = Style(
                        data=self.stringroot,
                        font_name=font_name,
                        font_size=font_size,
                        color=color,
                        parent_style_name=span.attrib[Styles.STYLE_NAME], 
                        name=style_name,
                        family=family,
                        text_position=None
                    )

                    parent_style = self.__get_style(span.attrib[Styles.STYLE_NAME])
                    properties = next(parent_style.iter(Styles.TEXT_PROPERTIES))
                    if Styles.TEXT_POSITION in properties.keys():
                        style.text_position = properties.attrib[Styles.TEXT_POSITION]

                    automatic_styles.append(
                        style.create_element_style()
                    )

                    span.attrib[Styles.STYLE_NAME] = style.name

                break
    
    def __get_style(self, name: str) -> ET.Element:
        """A helper method that returns a style object by name"""
        automatic_styles = self.stringroot.find(Styles.AUTOMATIC_STYLE)
        for i in automatic_styles.iter(Styles.STYLE):
            if i.attrib[NameSpaces.STYLE+'name'] == name:
                return i
            
    @staticmethod
    def __get_root(data, filename, encoding):
        return ET.fromstring(data[filename].decode(encoding))
            
    @staticmethod
    def __load_file(path) -> None:
        """Function for loading a file by path"""
        with zipfile.ZipFile(path, 'r') as zip_ref:
            file_data = {}
            for file in zip_ref.namelist():
                with zip_ref.open(file) as f:
                    file_data[file] = f.read()
        return file_data
            
    @staticmethod
    def __get_text_from_children(paragraph) -> str:
        """An helper method that returns text from children objects"""
        return ' '.join(filter(None, [paragraph.text] + [i.text for i in paragraph.iter(Styles.SPAN)]))

