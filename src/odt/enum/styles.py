from src.odt.enum.namespaces import NameSpaces

class Styles:
    """Class with styles"""
    PARAGRAPH = NameSpaces.TEXT + 'p'
    SPAN = NameSpaces.TEXT + 'span'
    A = NameSpaces.TEXT + 'a'
    AUTOMATIC_STYLE = NameSpaces.OFFICE + 'automatic-styles'
    STYLE = NameSpaces.STYLE + 'style'
    STYLE_NAME = NameSpaces.TEXT + 'style-name'
    TEXT_POSITION = NameSpaces.STYLE + 'text-position'
    TEXT_PROPERTIES = NameSpaces.STYLE+'text-properties'
    ANNOTATION = NameSpaces.OFFICE + 'annotation'
    ANNOTATION_END = NameSpaces.OFFICE + 'annotation-end'