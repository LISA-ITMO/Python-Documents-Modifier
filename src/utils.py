from docx.shared import RGBColor
from src.exceptions import ImpossibleColor


class Color(RGBColor):
    """
    RGB Class with some colors
    """
    red = RGBColor(255, 0, 0)
    green = RGBColor(0, 255, 0)
    blue = RGBColor(0, 0, 255)
    black = RGBColor(0, 0, 0)
    white = RGBColor(255, 255, 255)
    yellow = RGBColor(255, 255, 0)
    pink = RGBColor(255, 0, 255)
    cyan = RGBColor(0, 255, 255)
    orange = RGBColor(255, 128, 0)

    def __new__(cls, r, g, b):
        try:
            return super(Color, cls).__new__(cls, r, g, b)
        except ValueError as _ve:
            raise ImpossibleColor(
                f'Color ({r}, {g}, {b}) is impossible, values must be 0-255'
            )
