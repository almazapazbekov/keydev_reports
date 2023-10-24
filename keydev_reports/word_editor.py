import re
from typing import NoReturn

from docx import Document
from docx.text.paragraph import Paragraph as P
from python_docx_replace.paragraph import Paragraph

from .base import BaseEditor


class WordEditor(BaseEditor):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        """
        Конструктор класса. Принимает имя файла и загружает его.
        :param file_name: Имя файла в формате DOCX.
        """
        self.template_document = Document(self.file_name)

    @staticmethod
    def replace_text_in_paragraph(paragraph: P, key: str, value: str) -> NoReturn:
        """
        Статический метод для замены текста по ключу и значению.
        :param paragraph: Текст, в котором производится замена.
        :param key: Ключ, который нужно заменить.
        :param value: Значение для замены.
        """
        if key in paragraph.text:
            inline = paragraph.runs
            for item in inline:
                if key in item.text:
                    item.text = item.text.replace(key, value)

    def docx_replace(self, **kwargs: str) -> NoReturn:
        """
        Метод для замены ключей в документе DOCX на соответствующие значения.
        :param kwargs: Словарь, где ключи - ключи для замены, значения - новые значения.
        """
        pattern = r'\$\{[^}]*\}'
        for p in Paragraph.get_all(self.template_document):
            paragraph = Paragraph(p)
            for key, value in kwargs.items():
                key = f'${{{key}}}'
                paragraph.replace_key(key, str(value))
            for match in re.finditer(pattern, paragraph.p.text):
                key = match.group()
                paragraph.replace_key(key, '')

    def save_word(self, new_file_name=None) -> NoReturn:
        """
        Метод для сохранения документа DOCX с новым именем (по умолчанию перезапись текущего).
        :param new_file_name: Новое имя файла, если задано, иначе будет перезапись текущего файла.
        """
        self.file_name = new_file_name if new_file_name else self.file_name
        self.template_document.save(self.file_name)
