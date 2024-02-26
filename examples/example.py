from src.edocx import EDocx
from src.utils import Color

if __name__ == '__main__':
    doc = EDocx('example.docx')  # Open example.docx
    doc.autosave = True  # Enable autosave changes

    attrib = doc.all_para_attributes()  # Get all paragraph attributes of a file

    paraId = attrib[0].values()[0]  # Get paraId of first paragraph

    # Add comment to first paragraph with text 'Example comment' by 'Manager' and print result
    print(doc.add_comment_by_id(paraId, 'Example comment', 'Manager'))

    # Edit font size of first paragraph and print result
    print(doc.edit_font_size_by_id(paraId, 14))

    # Edit text color of first paragraph and print result
    print(doc.edit_color_by_id(paraId, Color.blue))
