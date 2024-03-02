from src.edocx import EDocx
from src.enum.color import Color
from src.enum.font_style import FontStyle
from src.enum.underline_style import UnderlineStyle

if __name__ == '__main__':
    doc = EDocx('example.docx')  # Open example.docx
    doc.autosave = True  # Enable autosave changes

    attrib = doc.all_para_attributes()  # Get all paragraph attributes of a file

    paraId = attrib[1].values()[0]  # Get paraId of first paragraph

    # Add comment to first paragraph with text 'Example comment' by 'Manager' and print result
    doc.add_comment_by_id(paraId, 'Example comment', 'Manager')

    # Edit style of first paragraph and print result
    doc.edit_style_by_id(
        paraId=paraId,
        size=14,
        color=Color.PINK,
        italic=False,
        bold=True,
        underline=UnderlineStyle.WAVY_DOUBLE,
        fontStyle=FontStyle.TIMES_NEW_ROMAN
    )
