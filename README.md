[![python](https://badgen.net/badge/python/3.9|3.10|3.11/blue?icon=python)](https://www.python.org/)
![ITMO](https://raw.githubusercontent.com/aimclub/open-source-ops/43bb283758b43d75ec1df0a6bb4ae3eb20066323/badges/ITMO_badge_rus.svg)
[![codecov](https://codecov.io/gh/LISA-ITMO/Python-Documents-Modifier/graph/badge.svg?token=QA6VQJE7AY)](https://codecov.io/gh/LISA-ITMO/Python-Documents-Modifier)

# Python-Documents-Modifier

___
![itmo](https://camo.githubusercontent.com/b9e4dd42874893b566fbc4c77daa19012408f5b5411a0625bb6b1a8e0212b39f/68747470733a2f2f69746d6f2e72752f66696c652f70616765732f3231332f6c6f676f5f6e615f706c6173686b655f727573736b69795f62656c79792e706e67)
## Description

___
The library for editing DOCX and ODT documents was developed to provide the necessary functionality when checking documents for compliance with standards.

## Features of library

___
The library provides the following features:
* Adding comments ```DOCX & ODT```
* Editing text style ```DOCX & ODT```
* Deleting comments ```DOCX & ODT```
* Editing comments ```DOCX```
* Editing list style ```DOCX```

## Requirements

___
```
lxml==5.1.0
python-docx==1.1.0
typing_extensions==4.9.0
```

## Installation

___
```git clone https://github.com/LISA-ITMO/Python-Documents-Modifier.git```

## Getting started

___
Install the library, open your script and create an instance of the class
```pycon
from src.docx.docx_redactor import DOCXRedactor
from src.odt.odtredactor import ODTRedactor

doc_docx = DOCXRedactor('your_document.docx')
doc_odt = ODTRedactor('your_document.odt')
```

## Examples for using functions

___
1) Add a comment
    ```pycon
    paraId = '00F00080' # DOCX | ID of the paragraph to which you want to add a comment
    text = 'Research has shown that' # ODT | Text to which you want to add a comment
    
    doc_docx.add_comment_by_id(paraId, 'your_comment', 'author')
    doc_odt.add_comment_by_text(text, 'your_comment', 'author')
    ```
2) Edit a comment
    ```pycon
    commentId = '0' # DOCX | ID of the comment you want to edit
    doc_docx.edit_comment_by_id(commentId, 'new_comment_text', 'new_author')
    ```
3) Delete a comment
    ```pycon
    commentId = '0' # DOCX | ID of the comment you want to delete
    nameId = '1' # ODT | NAME_ID of the comment you want to delete
   
    doc_docx.delete_comment_by_id(commentId)
    doc_odt.delete_comment_by_id(nameId)
    ```
4) Change comment style
    ```pycon
    from src.docx.enum.font_style import FontStyle
    from src.docx.enum.underline_style import UnderlineStyle
    from src.docx.enum.color import Color
    paraId = '00F00080' # DOCX | ID of the paragraph to which you want to edit style
    text = 'Scientists established back in 1984 that' # ODT | Text to which you want to edit style
   
    doc_docx.edit_style_by_id
    (
        paraId,
        size=12,
        fontStyle=FontStyle.ARIAL,
        color=Color.RED,
        underline=UnderlineStyle.DOUBLE,
        italic=False,
        bold=True
    )
   
    doc_odt.edit_style_by_text
    (
        text,
        font_name='Arial',
        font_size=12
    )
    ```
5) Edit list style
    ```pycon
    from src.docx.enum.ListStyle import ListStyle
    paraIds = ['00F00080', '11D11171'] # DOCX | IDs of the paragraphs, that contains numPr (included in list)
    doc_docx.edit_list_style_by_paraIds
    (
        paraIds,
        list_style=ListStyle.bullet,
        bullet_symbol='@'
    )
    ```

## Contacts

___
Your contacts. For example:

slavamarcin@yandex.ru\
vlad-tershch@yandex.ru

## Authors

___
[Shafikov Maxim](https://github.com/MrAmfix)\
[Krylov Michael](https://github.com/Inf1nity2483)\
[Tereshchenko Vladislav](https://github.com/Vl-Tershch)\
[Martsinkevich Viacheslav](https://github.com/slavamarcin)