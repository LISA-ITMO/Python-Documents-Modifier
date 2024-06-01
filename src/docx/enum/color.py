from docx.shared import RGBColor
from src.exceptions.docx_exceptions import ImpossibleColor


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
                (r, g, b)
            )
