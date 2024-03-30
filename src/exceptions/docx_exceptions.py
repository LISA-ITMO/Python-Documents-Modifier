
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


class NumberingIsNotExists(Exception):
    """
    The error called when you try to open DOCX-file, that doesn't contain
    word/numbering.xml file
    """
    def __int__(self, path: str) -> None:
        super().__init__(f'File \'{path}\' does not contain \'numbering.xml\'')
