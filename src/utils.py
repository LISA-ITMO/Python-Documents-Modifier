from docx.shared import RGBColor
from src.exceptions import ImpossibleColor


class Color(RGBColor):
    """
    RGB Class with some colors
    """
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 255, 0)
    BLUE = RGBColor(0, 0, 255)
    BLACK = RGBColor(0, 0, 0)
    WHITE = RGBColor(255, 255, 255)
    YELLOW = RGBColor(255, 255, 0)
    PINK = RGBColor(255, 0, 255)
    CYAN = RGBColor(0, 255, 255)
    ORANGE = RGBColor(255, 128, 0)

    def __new__(cls, r, g, b):
        try:
            return super(Color, cls).__new__(cls, r, g, b)
        except ValueError as _ve:
            raise ImpossibleColor(
                f'Color ({r}, {g}, {b}) is impossible, values must be 0-255'
            )


class UnderlineStyle:
    """
    Class with styles for underline
    """
    NONE = 0
    """None underline."""

    SINGLE = 1
    """A single line underline."""

    WORDS = 2
    """Underline individual words only."""

    DOUBLE = 3
    """A double line underline."""

    DOTTED = 4
    """Dotted underline."""

    THICK = 6
    """A thick line underline."""

    DASH = 7
    """Dashed underline."""

    DOT_DASH = 9
    """Alternating dots and dashes underline."""

    DOT_DOT_DASH = 10
    """An alternating dot-dot-dash pattern underline."""

    WAVY = 11
    """Wavy underline."""

    DOTTED_HEAVY = 20
    """Heavy dots underline."""

    DASH_HEAVY = 23
    """Heavy dashes underline."""

    DOT_DASH_HEAVY = 25
    """Alternating heavy dots and heavy dashes underline."""

    DOT_DOT_DASH_HEAVY = 26
    """An alternating heavy dot-dot-dash pattern underline."""

    WAVY_HEAVY = 27
    """Heavy wavy underline."""

    DASH_LONG = 39
    """Long dashes underline."""

    WAVY_DOUBLE = 43
    """Double wavy underline."""

    DASH_LONG_HEAVY = 55
    """Long heavy dashes underline."""


class FontStyle:
    """
    Class with font styles
    """
    ARIAL = 'Arial'
    CALIBRI = 'Calibri'
    TIMES_NEW_ROMAN = 'Times New Roman'
    COURIER_NEW = 'Courier New'
    VERDANA = 'Verdana'
    GEORGIA = 'Georgia'
    HELVETICA = 'Helvetica'
    TAHOMA = 'Tahoma'
    TREBUCHET_MS = 'Trebuchet MS'
    IMPACT = 'Impact'
    COMIC_SANS_MS = 'Comic Sans MS'
    SYMBOL = 'Symbol'
    WINGDINGS = 'Wingdings'
