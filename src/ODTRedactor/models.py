import random
from datetime import datetime
import xml.etree.ElementTree as ET


class Annotation:
    def __init__(self, text_annotation: str, author: str) -> None:
        self.text_annotation = text_annotation
        self.author = author
        self.id = random.randint(100_000_000, 999_999_999)
        self.name = random.randint(100000, 999999)
        self.time = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
        self.initials = self.get_initials()

    def get_initials(self) -> str:
        return ''.join([i[0] for i in self.author.split()])

    def get_end(self) -> ET.Element:
        annotation_end = ET.Element('text:span')
        annotation_end.attrib['text:style-name'] = "Знакпримечания"
        span = ET.SubElement(annotation_end, 'office:annotation-end', {'office:name': str(self.name)})
        return annotation_end

    def get_start(self) -> ET.Element:
        annotation = ET.Element('office:annotation')
        annotation.attrib['office:name'] = str(self.name)
        annotation.attrib['xml:id'] = str(self.id)

        creator = ET.SubElement(annotation, 'dc:creator')
        creator.text = self.author

        date = ET.SubElement(annotation, 'dc:date')
        date.text = self.time

        initials = ET.SubElement(annotation, 'meta:creator-initials')
        initials.text = self.initials

        p = ET.SubElement(annotation, 'text:p', {'text:style-name': 'Текстпримечания'})

        span = ET.SubElement(p, 'text:span', {'text:style-name': 'основнойшрифтабзаца'})
        span.text = self.text_annotation

        p = ET.SubElement(annotation, 'text:p', {'text:style-name': 'P5'})
        return annotation
    
class Style:
    def __init__(self, data: ET, font_name: str, font_size: int, color: str, parent_style_name: str, family: str, name: str, text_position: str) -> None:
        self.data: ET.Element = data 
        self.font_name = font_name
        self.font_size = str(font_size) + 'pt'
        self.color = color
        self.name = name if name else self.get_name()
        self.parent_style_name = parent_style_name
        self.family = family
        self.text_position = text_position

    def get_name(self) -> str:
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
    
class NameSpaces:
    STYLE = '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
    TEXT = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
    FO = '{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}'
    OFFICE = '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
    
class Styles:
    PARAGRAPH = NameSpaces.TEXT + 'p'
    SPAN = NameSpaces.TEXT + 'span'
    A = NameSpaces.TEXT + 'a'
    AUTOMATIC_STYLE = NameSpaces.OFFICE + 'automatic-styles'
    STYLE = NameSpaces.STYLE + 'style'
    TEXT_POSITION = NameSpaces.STYLE + 'text-position'