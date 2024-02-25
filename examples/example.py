from src.edocx import EDocx

if __name__ == '__main__':
    doc = EDocx('example.docx')  # Open example.docx

    attrib = doc.all_para_attributes()  # Get all paragraph attributes of a file

    paraId = attrib[0].values()[0]  # Get paraId of first paragraph

    # Add comment to first paragraph with text 'Example comment' by 'Manager'
    doc.add_comment_by_id(paraId, 'Example comment', 'Manager')
