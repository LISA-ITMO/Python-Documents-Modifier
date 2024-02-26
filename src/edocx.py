from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from src.exceptions import NotSupportedFormat
from src.utils import Color
from typing import List, Dict


class EDocx:
    """
    Class for edit DOCX document objects
    """
    def __init__(self, path: str) -> None:
        """
        :param path: path to DOCX document
        """
        if path[-5:] != '.docx':
            raise NotSupportedFormat(
                f'\'{path}\' should be .docx'
            )
        self.path: str = path
        self.document: Document = Document(path)
        self.autosave: bool = False

    def save(self) -> None:
        """
        Method for save changes in document
        """
        self.document.save(self.path)

    @property
    def __para_ids(self) -> Dict[str, Paragraph]:
        res = {}
        key = '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
        for para in self.document.paragraphs:
            try:
                res[para.paragraph_format.element.attrib[key]] = para
            except KeyError as _ke:
                continue
        return res

    def add_comment_by_id(self, paraId: str, comment: str, author: str = 'EDocx') -> bool:
        """
        Add comment to paragraph by paraId
        :param paraId: paragraph id
        :param comment: a comment
        :param author: author's nickname
        :return: True - comment has been added, False - paragraph not found
        """
        para_ids = self.__para_ids
        if paraId not in para_ids:
            return False
        else:
            para_ids[paraId].add_comment(comment, author)
            if self.autosave:
                self.save()
            return True

    def all_para_attributes(self) -> List[Dict[str, str]]:
        """
        Get all attributes from all paragraphs
        :return: List of dicts (key: attribute, value: value)
        """
        attributes = []
        for para in self.document.paragraphs:
            attributes.append(para.paragraph_format.element.attrib)
        return attributes

    def edit_font_size_by_id(self, paraId: str, font_size: int = 14) -> bool:
        """
        Edit font size of paragraph by paraId
        :param paraId: paragraph id
        :param font_size: new font size
        :return: True - font size has been changed, False - paragraph not found
        """
        para_ids = self.__para_ids
        if paraId not in para_ids:
            return False
        else:
            for run in para_ids[paraId].runs:
                run.font.size = Pt(font_size)
            if self.autosave:
                self.save()
            return True

    def edit_color_by_id(self, paraId: str, color: Color = Color.black) -> bool:
        """
        Edit text color of paragraph by paraId
        :param paraId: paragraph id
        :param color: text color
        :return: True - text color has been changed, False - paragraph not found
        """
        para_ids = self.__para_ids
        if paraId not in para_ids:
            return False
        else:
            for run in para_ids[paraId].runs:
                run.font.color.rgb = color
            if self.autosave:
                self.save()
            return True
