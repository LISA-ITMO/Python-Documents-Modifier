from docx import Document
from src.exceptions import NotSupportedFormat
from typing import List, Dict


class EDocx:
    """
    Class for edit DOCX document objects
    """
    def __init__(self, path: str):
        """
        :param path: path to DOCX document
        """
        if path[-5:] != '.docx':
            raise NotSupportedFormat(
                f'\'{path}\' should be .docx'
            )
        self.path = path
        self.document = Document(path)

    def add_comment_by_id(self, para_id: str, comment: str, author: str = 'EDocx') -> bool:
        """
        Add comment to paragraph by paraId
        :param para_id: paragraph id
        :param comment: a comment
        :param author: author's nickname
        :return: True - comment has been added, False - paraId not found
        """
        key = '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
        flag = False
        for para in self.document.paragraphs:
            try:
                if para.paragraph_format.element.attrib[key] == para_id:
                    para.add_comment(comment, author)
                    flag = True
                    break
            except KeyError as _ke:  # Skip paragraphs without paraId
                continue
        self.document.save(self.path)
        return flag

    def all_para_attributes(self) -> List[Dict[str, str]]:
        """
        Get all attributes from all paragraphs
        :return: List of dicts (key: attribute, value: value)
        """
        attributes = []
        for para in self.document.paragraphs:
            attributes.append(para.paragraph_format.element.attrib)
        return attributes
