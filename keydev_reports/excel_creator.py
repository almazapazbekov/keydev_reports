import io
import os
from datetime import date
from typing import Iterable, NoReturn, Optional

import xlsxwriter

from .base import BaseCreator


class ExcelCreator(BaseCreator):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        """
        Конструктор класса. Создает объект для создания Excel-отчета.

        :param filename: Имя файла для сохранения отчета (по умолчанию - в памяти).
        """
        self.output = io.BytesIO()
        self.row_width = 20  # Ширина строк по умолчанию
        self.workbook = xlsxwriter.Workbook(self.file_path if self.file_path else self.output)
        self.date_format = self.workbook.add_format({'num_format': 'yyyy-mm-dd'})

    def add_data(
            self,
            sheet_name: str,
            data: Iterable[Iterable],
            start_row: int = 0,
            cell_format: Optional = None) -> NoReturn:
        """
        Метод для добавления данных на лист отчета.
        :param sheet_name: Имя листа, на котором добавляются данные.
        :param data: Данные для добавления.
        :param start_row: Начальная строка для добавления данных (по умолчанию - 0).
        :param cell_format: Формат ячеек (по умолчанию - None).
        """
        ws = self.workbook.add_worksheet(sheet_name)
        ws.set_default_row(self.row_width)
        new_format = self.workbook.add_format(cell_format) if cell_format else None

        for row_index, row_data in enumerate(data):
            for col_num, cell_data in enumerate(row_data):
                if isinstance(cell_data, date):
                    write_format = self.date_format
                else:
                    write_format = new_format

                ws.write(start_row + row_index, col_num, cell_data, write_format)

        ws.autofit()

    def save_excel(self) -> NoReturn:
        """
        Метод для сохранения книги.
        :return:
        """
        if self.filename:
            self.workbook.close()
        else:
            # Сохраняем отчет в памяти и сбрасываем позицию объекта BytesIO
            self.workbook.close()
            self.output.seek(0)

    def get_bytes_io(self) -> bytes:
        """
        Метод для получения содержимого Excel-файла в виде объекта BytesIO.

        :return: bytes с содержимым Excel-файла.
        """
        return self.output.read()

    def get_filepath(self) -> Optional[str]:
        """
        Метод для получения полного пути к файлу отчета.

        Если указано имя файла (filename), метод возвращает полный путь к файлу
        с учетом указанного аргумента (output_directory), если он предоставлен.
        В противном случае метод возвращает None.
        :return: Полный путь к файлу отчета или None.
        """
        if self.filename:
            return os.path.join(self.output_directory, self.filename) if self.output_directory else self.filename
        else:
            return None
