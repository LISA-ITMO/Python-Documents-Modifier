
class NotSupportedFormat(Exception):
    """
    The error called when you input a file name of non-docx format.
    """
    def __init__(self, filename: str = 'file you open') -> None:
        super().__init__(f'The \'{filename}\' should be .docx')


class ImpossibleColor(Exception):
    """
    The error called when you try to input rgb-color with values not in range (0, 255)
    """
    def __init__(self, rgb: tuple) -> None:
        super().__init__(f'Color {rgb} is impossible, values should be 0-255')


class ParagraphNotFound(Exception):
    """
    The error called when you input paraId of a non-exists Paragraph
    """
    def __init__(self, paraId: str) -> None:
        super().__init__(f'Paragraph with paraId \'{paraId}\' not found')


class ParagraphDoesNotContainNumPr(Exception):
    """
    The error called when you try to replace num-value for paragraph, which doesn't have num
    """
    def __init__(self, paraId: str) -> None:
        super().__init__(f'Paragraph with paraId \'{paraId}\' does not contain numPr')


class ILvlDoesNotExist(Exception):
    """
    The error called when you try created ilvl with number not from range [0; 8]
    """
    def __init__(self, num: int):
        super().__init__(f'Ilvl with number \'{num}\' can\'t be created (range 0-8)')


class FileDoesNotContainXMLFile(Exception):
    """
    The error called when you try to edit any files in DOCX-file, that doesn't exist in DOCX-archive
    Example: you try edit any comment, but DOCX-archive doesn't contain comments.xml
    """
    def __init__(self, path: str, file: str):
        super().__init__(f'DOCX-archive \'{path}\' doesn\'t contain file \'{file}\'')
