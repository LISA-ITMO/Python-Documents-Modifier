from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from src.docx.enum.color import Color
from src.docx.enum.font_style import FontStyle
from src.docx.enum.underline_style import UnderlineStyle
from src.exceptions.docx_exceptions import NotSupportedFormat, ParagraphNotFound
from typing import List, Dict


class DOCXRedactor:
    """
    Class for edit DOCX document objects
    """
    def __init__(self, path: str) -> None:
        """
        :param path: path to DOCX document
        """
        if len(path) < 5 or len(path) > 255:
            raise NameError(
                f'\'{path}\' can\'t be open'
            )
        if path[-5:] != '.docx':
            raise NotSupportedFormat(
                path
            )
        self.path: str = path
        self.document: Document = Document(path)
        self.autosave: bool = False

    def save(self, path=None) -> None:
        """
        Method for save changes in document
        """
        if path is None:
            self.document.save(self.path)
        else:
            self.document.save(path)

    def __para_ids(self) -> Dict[str, Paragraph]:
        res = {}
        key = self.__get_schema_paraId()
        for para in self.document.paragraphs:
            try:
                res[para.paragraph_format.element.attrib[key]] = para
            except KeyError as _ke:
                continue
        return res

    def add_comment_by_id(self, paraId: str, comment: str, author: str = 'EDocx') -> None:
        """
        Add comment to paragraph by paraId
        :param paraId: paragraph id
        :param comment: a comment
        :param author: author's nickname
        """
        try:
            self.__para_ids()[paraId].add_comment(comment, author)
            if self.autosave:
                self.save()
        except KeyError as _ke:
            raise ParagraphNotFound(
                paraId
            )

    def all_para_attributes(self) -> List[Dict[str, str]]:
        """
        Get all attributes from all paragraphs
        :return: List of dicts (key: attribute, value: value)
        """
        attributes = []
        for para in self.document.paragraphs:
            attributes.append(para.paragraph_format.element.attrib)
        return attributes

    def edit_style_by_id(self, paraId: str,
                         size: int = None, color: Color = None,
                         fontStyle: FontStyle = None,
                         italic: bool = None, bold: bool = None,
                         underline: UnderlineStyle = None) -> None:
        """
        Edit font style of paragraph by paraId
        """
        try:
            for run in self.__para_ids()[paraId].runs:
                self.__edit_font_size(run, size)
                self.__edit_text_color(run, color)
                self.__edit_font_style(run, fontStyle)
                self.__edit_italic(run, italic)
                self.__edit_bold(run, bold)
                self.__edit_underline(run, underline)

            if self.autosave:
                self.save()
        except KeyError as _ke:
            raise ParagraphNotFound(
                paraId
            )

    def __get_schema_paraId(self) -> str:
        return f'{{{self.document.element.nsmap["w14"]}}}paraId'

    def __edit_font_size(self, run: Run, size: int) -> None:
        if size is not None:
            run.font.size = Pt(size)

    def __edit_text_color(self, run: Run, color: Color) -> None:
        if color is not None:
            run.font.color.rgb = color

    def __edit_font_style(self, run: Run, fontStyle: FontStyle) -> None:
        if fontStyle is not None:
            run.font.name = fontStyle

    def __edit_italic(self, run: Run, italic: bool) -> None:
        if italic is not None:
            run.font.italic = italic

    def __edit_bold(self, run: Run, bold: bool) -> None:
        if bold is not None:
            run.font.bold = bold

    def __edit_underline(self, run: Run, underline: UnderlineStyle):
        if underline is not None:
            run.font.underline = underline
