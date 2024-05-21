from src.odt.enum.styles import Styles
from src.odt.enum.namespaces import NameSpaces

import xml.etree.ElementTree as ET

class Style:
    """A class for generating a style in a document"""
    def __init__(self, data: ET, font_name: str, font_size: int, color: str, parent_style_name: str, family: str, name: str, text_position: str) -> None:
        self.data: ET.Element = data 
        self.font_name = font_name
        self.font_size = str(font_size) + 'pt'
        self.color = color
        self.name = name if name else self.get_style_name()
        self.parent_style_name = parent_style_name
        self.family = family
        self.text_position = text_position

    def get_style_name(self) -> str:
        automatic_styles = self.data.find(Styles.AUTOMATIC_STYLE)
        styles = automatic_styles.findall(Styles.STYLE)
        return "T" + str(len(styles) + 1)
    
    def create_element_style(self) -> ET.Element:
        style = ET.Element(NameSpaces.STYLE+'style')
        style.attrib[NameSpaces.STYLE+'name'] = self.name
        style.attrib[NameSpaces.STYLE+'parent-style-name'] = self.parent_style_name
        style.attrib[NameSpaces.STYLE+'family'] = self.family

        text_properties = ET.SubElement(style, NameSpaces.STYLE+'text-properties')
        text_properties.attrib[NameSpaces.STYLE+'font-name'] = self.font_name
        text_properties.attrib[NameSpaces.STYLE+'font-name-complex'] = self.font_name
        text_properties.attrib[NameSpaces.FO+'color'] = self.color
        text_properties.attrib[NameSpaces.FO+'font-size'] = self.font_size
        text_properties.attrib[NameSpaces.STYLE+'font-size-asian'] = self.font_size
        text_properties.attrib[NameSpaces.STYLE+'font-size-complex'] = self.font_size
        if self.text_position: text_properties.attrib[Styles.TEXT_POSITION] = self.text_position

        return style