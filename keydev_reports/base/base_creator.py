import os
from typing import Optional


class BaseCreator:

    def __init__(self, filename: Optional[str] = None, output_directory: Optional[str] = None):
        self.filename = filename
        self.output_directory = output_directory
        self.file_path = self._build_file_path()

    def _build_file_path(self) -> str:
        """
        Внутренний метод для построения полного пути к файлу отчета.
        Если указана директория (output_directory), то метод построит полный путь к файлу.
        Если такая директория не существует, метод создаст директорию и построит путь к файлу.
        :return: Полный путь к файлу отчета.
        """
        if self.output_directory:
            if not os.path.exists(self.output_directory):
                os.makedirs(self.output_directory)
        file_path = os.path.join(self.output_directory, self.filename) if self.output_directory else self.filename
        return file_path
