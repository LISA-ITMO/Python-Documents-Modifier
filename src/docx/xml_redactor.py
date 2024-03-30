from docx import Document
from docx.oxml import parse_xml
from src.exceptions.docx_exceptions import NotSupportedFormat, NumberingIsNotExists


class XMLRedactor:
    """
    At the moment the class works only with word/numbering.xml
    """
    def __init__(self, path: str) -> None:
        if len(path) < 5 or len(path) > 255:
            raise NameError(
                f'\'{path}\' can\'t be open'
            )
        if path[-5:] != '.docx':
            raise NotSupportedFormat(
                path
            )
        self.path: str = path
        self.document: Document = Document(path)
        self.autosave: bool = False
        try:
            self.numbering_xml_file = self.document.part.numbering_part
            self.numbering_xml_tree = self.numbering_xml_file._element
        except KeyError as _ke:
            raise NumberingIsNotExists(
                path
            )

    def save(self, path: str = None):
        if path is None:
            self.document.save(self.path)
        else:
            self.document.save(path)

    def __get_schema(self, key: str) -> str:
        return self.document.element.nsmap[key]

    @property
    def __get_schema_numId(self) -> str:
        return f'{{{self.__get_schema("w")}}}numId'

    @property
    def __get_schema_abstractNumId(self) -> str:
        return f'{{{self.__get_schema("w")}}}abstractNumId'

    @property
    def __get_new_numId(self) -> int:
        num_elements = self.numbering_xml_tree.findall(
            './/w:num', namespaces=self.numbering_xml_file._element.nsmap
        )
        return 1 + max(int(num.get(self.__get_schema_numId)) for num in num_elements)

    @property
    def __get_new_abstractId(self) -> int:
        anum_elements = self.numbering_xml_tree.findall(
            './/w:abstractNum', namespaces=self.numbering_xml_file._element.nsmap
        )
        return 1 + max(int(anum.get(self.__get_schema_abstractNumId)) for anum in anum_elements)

    def add_new_num(self) -> None:
        """
        !!! Непонятно для чего durableId
        """
        num_element = parse_xml(f'''
            <w:num xmlns:w="{self.__get_schema("w")}" xmlns:w16cid="{self.__get_schema("w16cid")}" w:numId="{self.__get_new_numId}" w16cid:durableId="0">
                <w:abstractNumId w:val="{self.__get_new_abstractId}"/>
            </w:num>
        ''')
        self.numbering_xml_tree.append(num_element)
        if self.autosave:
            self.document.save(self.path)

    def add_new_abstractNum(self, lvl: int = 0) -> None:
        """
        !!! Много аттрибутов, непонятно какие нужны и за что отвечают
        !!! Проблема с bullet list, как различать различные символы, пока только о
        """
        lvl_atr_dec = '''<w:start w:val="1"/>
                        <w:numFmt w:val="decimal"/>
                        <w:lvlText w:val="%1."/>
                        <w:lvlJc w:val="left"/>'''
        lvl_atr_bullet = '''<w:start w:val="1"/>
                        <w:numFmt w:val="bullet"/>
                        <w:lvlText w:val="o"/>
                        <w:lvlJc w:val="left"/>'''
        abstract_num_element = parse_xml(f'''
                <w:abstractNum xmlns:w="{self.__get_schema("w")}" w:abstractNumId="{self.__get_new_abstractId}">
                    <w:nsid w:val="2A733D56"/>
                    <w:multiLevelType w:val="hybridMultilevel"/>
                    <w:tmpl w:val="47AAD25A"/>
                    <w:lvl w:ilvl="{lvl}">
                        <w:start w:val="1"/>
                        <w:numFmt w:val="decimal"/>
                        <w:lvlText w:val="%1."/>
                        <w:lvlJc w:val="left"/>
                        <w:pPr>
                            <w:ind w:left="720" w:hanging="360"/>
                        </w:pPr>
                        <w:rPr>
                            <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Wingdings"/>
                        </w:rPr>
                    </w:lvl>
                </w:abstractNum>
            ''')
        self.numbering_xml_tree.append(abstract_num_element)
        if self.autosave:
            self.document.save(self.path)

