import shutil
import tempfile
from docx import Document
from docx.shared import Pt
from docx.text.run import Run
from src.docx.xml_redactor import XMLRedactor
from src.docx.enum.schemas import schemas
from src.docx.enum.color import Color
from src.docx.enum.font_style import FontStyle
from src.docx.enum.ListStyle import ListStyle
from src.docx.enum.underline_style import UnderlineStyle
from src.exceptions.docx_exceptions import NotSupportedFormat, ParagraphNotFound
from typing import List, Dict, Union


def open_docx_file(func):
    def wrapper(self, *args, **kwargs):
        document = Document(self._temp_file)
        res = func(self, document=document, *args, **kwargs)
        document.save(self._temp_file)
        return res

    return wrapper


def open_xml_redactor(func):
    def wrapper(self, *args, **kwargs):
        xr = XMLRedactor(self._temp_file)
        res = func(self, _xml_redactor=xr, *args, **kwargs)
        xr.save()
        return res

    return wrapper


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
        self._temp_file: str = self.__init_temp_file()

    def __init_temp_file(self) -> str:
        return shutil.copy(self.path, tempfile.mkdtemp())

    def save(self, path=None) -> None:
        """
        Method for save changes in document
        """
        if path is None:
            path = self.path
        shutil.move(self._temp_file, path)

    @open_docx_file
    def add_comment_by_id(self, paraId: str, comment: str, author: str = 'EDocx',
                          document=None) -> None:
        """
        Add comment to paragraph by paraId
        :param paraId: paragraph id
        :param comment: a comment
        :param author: author's nickname
        """
        try:
            for para in document.paragraphs:
                if para.paragraph_format.element.attrib[f'{{{schemas.w14}}}paraId'] == paraId:
                    para.add_comment(comment, author)
        except KeyError as _ke:
            raise ParagraphNotFound(
                paraId
            )

    @open_xml_redactor
    def edit_comment_by_id(self, comment_id: Union[str, int],
                           comment_text: str, new_author: str = None,
                           _xml_redactor: XMLRedactor = None):
        """
        Edit comment by comment_id
        :param _xml_redactor: XMLRedactor-object, automatic generate by decorator
        :param comment_id: id of comment, that should be edited
        :param comment_text: new text of comment
        :param new_author: new author of comment
        """
        _xml_redactor.edit_comment_by_id(comment_id, comment_text, new_author)

    @open_xml_redactor
    def delete_comment_by_id(self, comment_id: Union[str, int], _xml_redactor: XMLRedactor = None):
        """
        :param _xml_redactor: XMLRedactor-object, automatic generate by decorator
        :param comment_id: id of comment, that should be deleted
        """
        _xml_redactor.delete_comment_by_id(comment_id)

    @open_xml_redactor
    def edit_list_style_by_paraIds(self, paraIds: Union[List[str], str],
                                   list_style: bool = ListStyle.decimal,
                                   bullet_symbol: str = '•', _xml_redactor: XMLRedactor = None):
        """
        :param _xml_redactor: XMLRedactor-object, automatic generate by decorator
        Replace the numId to newNum of all specified paragraphs
        :param paraIds: One or many paraIds, which we replace numId
        :param list_style: Style of new List (bullet of decimal)
        :param bullet_symbol: If you select bullet-type of list, this param make bullet-symbol to custom
        """
        new_num = _xml_redactor.add_new_abstract_and_num(list_style=list_style, levels_count=8,
                                                         bullet_symbol=bullet_symbol)
        _xml_redactor.edit_list_style_by_paraIds(paraIds, new_num)

    @open_docx_file
    def all_para_attributes(self, document=None) -> List[Dict[str, str]]:
        """
        Get all attributes from all paragraphs
        :return: List of dicts (key: attribute, value: value)
        """
        attributes = []
        for para in document.paragraphs:
            attributes.append(para.paragraph_format.element.attrib)
        return attributes

    @open_docx_file
    def edit_style_by_id(self, paraId: str,
                         size: int = None, color: Color = None,
                         fontStyle: FontStyle = None,
                         italic: bool = None, bold: bool = None,
                         underline: UnderlineStyle = None,
                         document=None) -> None:
        """
        Edit font style of paragraph by paraId
        """
        flag = False
        for para in document.paragraphs:
            if para.paragraph_format.element.attrib[f'{{{schemas.w14}}}paraId'] == paraId:
                flag = True
                for run in para.runs:
                    self.__edit_font_size(run, size)
                    self.__edit_text_color(run, color)
                    self.__edit_font_style(run, fontStyle)
                    self.__edit_italic(run, italic)
                    self.__edit_bold(run, bold)
                    self.__edit_underline(run, underline)
        if not flag:
            raise ParagraphNotFound(paraId)


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
