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
        span = ET.SubElement(p, 'text:span', {'text:style-name': 'T2'})
        span.text = self.text_annotation
        p = ET.SubElement(annotation, 'text:p', {'text:style-name': 'P5'})
        return annotation